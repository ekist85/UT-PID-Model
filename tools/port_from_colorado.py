#!/usr/bin/env python3
"""
port_from_colorado.py — re-port ``ut_pid_model`` from ``co_metro_model``.

The Utah model is a port of the Colorado metropolitan district model, not a
fork: the engines (revenue-wrap sizing, surplus fund, subordinate waterfall,
refunding, pricing, scenarios) and the whole reporting layer are shared, and
only the property tax framework differs.  Keeping that relationship mechanical
means a Colorado release can be pulled forward in one command instead of being
re-diffed by hand.

    python tools/port_from_colorado.py --source ../CO-Metro-District-Model
    python tools/port_from_colorado.py --source ../CO-Metro-District-Model --check

The port runs in three passes over each Colorado module:

  1. **Identifier renames** — whole-word, longest-first, for the attributes whose
     Colorado name names a Colorado mechanism (``tabor_current`` →
     ``resid_taxable_ratio``, ``oil_gas_*`` → ``centrally_assessed_*``, …).
  2. **Label substitutions** — the display strings that reach a workbook, a memo
     or the console ("Assessed Value" → "Taxable Value", "Vacant Land" →
     "Builder Lot Inventory", …).
  3. **Semantic patches** — the places where Utah law actually changes the
     arithmetic or the prose.

Every patch is checked.  If Colorado refactors the code a patch anchors to, the
port **fails loudly** rather than quietly emitting a Colorado-flavoured Utah
model — which is the whole point of doing it this way.

``config.py`` is ported too, so the Utah defaults and the statutory tables live
in the patch set below and survive a re-port.
"""

from __future__ import annotations

import argparse
import difflib
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent

#: Colorado modules that are ported wholesale.
MODULES = [
    "__init__.py", "config.py", "development.py", "summary.py",
    "debt_service.py", "subordinate.py", "sources_uses.py", "series_c.py",
    "refunding.py",
    "pricing.py", "inputs.py", "report.py", "residential_report.py",
    "forecast_report.py", "scenarios.py", "memo.py",
]
#: Top-level scripts ported alongside the package.
SCRIPTS = ["main.py", "build_notebook.py"]


# ── Pass 1: identifier renames (whole word, longest first) ────────────────────

IDENTIFIER_RENAMES = [
    ("co_metro_model_notebook", "ut_pid_model_notebook"),
    ("co_metro_model_output", "ut_pid_model_output"),
    ("co_metro_model_inputs", "ut_pid_model_inputs"),
    ("co_metro_model_memo", "ut_pid_model_memo"),
    ("co_metro_forecast_exhibits", "ut_pid_forecast_exhibits"),
    ("co_metro_model", "ut_pid_model"),
    ("metro_name", "pid_name"),
    ("tabor_service_plan", "resid_taxable_ratio_prior"),
    ("tabor_current", "resid_taxable_ratio"),
    ("_TABOR_SOURCES", "_RATE_SOURCES"),
    ("_TABOR_HISTORY", "_RATE_HISTORY"),
    ("_VACANT_LAND_HISTORY", "_INVENTORY_HISTORY"),
    ("_VACANT_TABLE_TITLE", "_INVENTORY_TABLE_TITLE"),
    ("_write_tabor_history_sheet", "_write_reference_sheet"),
    ("TABOR_HISTORY", "RESIDENTIAL_EXEMPTION_HISTORY"),
    ("_tabor_year_bounds", "_rate_year_bounds"),
    ("tabor_ratio", "residential_taxable_ratio_for"),
    ("VACANT_LAND_HISTORY", "BUILDER_INVENTORY_HISTORY"),
    ("vacant_land_assessment_schedule", "lot_inventory_taxable_schedule"),
    ("vacant_land_assessment_rate", "lot_inventory_taxable_rate"),
    ("vacant_land_rate_schedule", "lot_inventory_rate_schedule"),
    ("vacant_land_value_build", "lot_inventory_value_build"),
    ("_build_vacant_land_value_sheet", "_build_lot_inventory_value_sheet"),
    ("vacant_land_ratio", "lot_inventory_ratio"),
    ("developed_lot_value", "lot_inventory_taxable_ratio"),
    ("mill_levy_service_plan", "mill_levy_governing_doc"),
    ("series_c_biennial_reassess_rate", "series_c_reassess_rate"),
    ("SERIES_C_BIENNIAL_REASSESS_RATE", "SERIES_C_REASSESS_RATE"),
    ("biennial_reassess_comm_rate", "reassess_comm_rate"),
    ("biennial_reassess_rate", "reassess_rate"),
    ("strict_biennial_av", "hold_value_flat"),
    ("tax_collect_so_av_threshold", "uniform_fee_av_threshold"),
    ("tax_collect_so_prc", "uniform_fee_prc"),
    ("specific_ownership_tax", "uniform_fee_revenue"),
    ("county_treasurer_fee", "county_collection_fee"),
    ("oil_gas_assets", "centrally_assessed_value"),
    ("oil_gas_equipment", "centrally_assessed_equipment"),
    ("oil_gas_value", "centrally_assessed_ratio"),
    ("oil_gas_av", "centrally_assessed_av"),
    ("has_oil_gas", "has_centrally_assessed"),
    ("oil_gas", "centrally_assessed"),
    ("om_carveout_av_limit", "om_expense_av_limit"),
    ("om_carveout", "om_expense"),
    # Locals/keys that still carried Colorado vocabulary.  Utah has no specific
    # ownership tax (it is the § 59-2-405 uniform fee), no TABOR, and no county
    # treasurer's fee netted from the distribution (§ 59-2-1602 funds assessing
    # and collecting through a separate statewide levy).
    ("sot", "uniform_fee"),
    ("tot_sot", "tot_uniform_fee"),
    ("treasurer_fee", "collection_fee"),
    ("tabor", "resid_ratio"),
]


# ── Pass 2: display-label substitutions (plain text, longest first) ───────────

LABEL_SUBS = [
    # Range names first — before the generic "TABOR" → "Residential Exemption".
    ("TABOR_CURRENT", "RESID_TAXABLE_RATIO"),
    ("TABOR_SERVICE_PLAN", "RESID_TAXABLE_RATIO_PRIOR"),
    ("DEVELOPED_LOT_VALUE", "LOT_INVENTORY_TAXABLE_RATIO"),
    ("BIENNIAL_REASSESS_COMM_RATE", "REASSESS_COMM_RATE"),
    ("BIENNIAL_REASSESS_RATE", "REASSESS_RATE"),
    ("TAX_COLLECT_SO_AV_THRESHOLD", "UNIFORM_FEE_AV_THRESHOLD"),
    ("TAX_COLLECT_SO_PRC", "UNIFORM_FEE_PRC"),
    ("COUNTY_TREASURER_FEE", "COUNTY_COLLECTION_FEE"),
    ("MILL_LEVY_SERVICE_PLAN", "MILL_LEVY_GOVERNING_DOC"),
    ("OIL_GAS_ASSETS", "CENTRALLY_ASSESSED_VALUE"),
    ("OIL_GAS_EQUIPMENT", "CENTRALLY_ASSESSED_EQUIPMENT"),
    ("OIL_GAS_VALUE", "CENTRALLY_ASSESSED_RATIO"),
    ("OIL_GAS", "CENTRALLY_ASSESSED"),
    ("OM_CARVEOUT_AV_LIMIT", "OM_EXPENSE_AV_LIMIT"),
    ("OM_CARVEOUT", "OM_EXPENSE"),
    ("STRICT_BIENNIAL_AV", "HOLD_VALUE_FLAT"),
    # Excel date format — m/d/yyyy throughout the workbook (number formats and
    # the one place the format string is compared for alignment).
    ('fmt="MM/DD/YYYY"', 'fmt="m/d/yyyy"'),
    ("DATEFMT = 'mm/dd/yyyy'", "DATEFMT = 'm/d/yyyy'"),
    ("fmt == 'mm/dd/yyyy'", "fmt == 'm/d/yyyy'"),
    ("_DATEFMT = 'YYYY-MM-DD'", "_DATEFMT = 'm/d/yyyy'"),          # inputs template
    ('p.payment_date, "DD-MMM-YY"', 'p.payment_date, "m/d/yyyy"'),  # forecast exhibits
    ("Schedule of Estimated Assessed Valuation", "Schedule of Estimated Taxable Value"),
    ("Summary of Assessed Values and Net Tax Revenues",
     "Summary of Taxable Values and Net Tax Revenues"),
    ("Assessed Valuation", "Taxable Value"),
    ("Assessed Values", "Taxable Values"),
    ("assessed valuation", "taxable value"),
    ("Assessed Value", "Taxable Value"),
    ("assessed value", "taxable value"),
    # Header cells wrap, so the newline-split forms need their own entries.
    ("Assessment\\nRatio", "Taxable\\nRatio"),
    ("Assessment\\nRate", "Taxable\\nRatio"),
    ("Assessed\\nValue of Lots", "Taxable\\nValue of Lots"),
    ("Assessed\\nValue", "Taxable\\nValue"),
    ("Assessed\\nValuation", "Taxable\\nValue"),
    ("Taxable Value\\nAssessed Valuation", "Taxable Value"),
    ("Assessment Ratio", "Taxable Ratio"),
    ("Assessment Rate", "Taxable Ratio"),
    ("assessment ratio", "taxable ratio"),
    ("assessment rate", "taxable ratio"),
    ("Specific Ownership Taxes", "Personal Property Uniform Fee"),
    ("Specific Ownership Tax", "Personal Property Uniform Fee"),
    ("specific ownership tax", "personal property uniform fee"),
    ("SOT Tax", "Uniform Fee"),
    ("SOT AV Threshold", "Uniform Fee Taxable Value Threshold"),
    ("SOT\\n", "Uniform\\nFee\\n"),
    ("+ SOT", "+ Uniform Fee"),
    ("County\\nTreasurer Fee", "County\\nCollection Cost"),
    ("Specific\\nOwnership Taxes", "Personal Property\\nUniform Fee"),
    ("Specific\\nOwnership Tax", "Personal Property\\nUniform Fee"),
    ("Metro District Name", "Public Infrastructure District Name"),
    ("County Treasurer Fee", "County Collection Cost"),
    ("County treasurer fee", "County collection cost"),
    ("county treasurer fee", "county collection cost"),
    ("County Treas.", "County Collection"),
    ("County Treasurer", "County Collection"),
    ("Oil & Gas Producing (actual value)", "Centrally Assessed Property (market value)"),
    ("Oil & Gas Equipment (actual value)", "Centrally Assessed Equipment (market value)"),
    ("Oil & Gas Assessment Ratio", "Centrally Assessed Taxable Ratio"),
    ("Include Oil & Gas in Taxed AV", "Include Centrally Assessed in Taxed Value"),
    ("Oil & Gas / Commercial", "Centrally Assessed / Commercial"),
    ("Oil & Gas", "Centrally Assessed"),
    ("oil & gas", "centrally assessed"),
    ("Existing Vacant Land", "Existing Land Value"),
    ("existing vacant land", "existing land value"),
    ("Vacant Land Market Value", "Lot Inventory Market Value"),
    ("the Vacant Land Value / Residential Value tabs",
     "the Builder Lot Inventory Value / Residential Value tabs"),
    ("Vacant Land Value", "Builder Lot Inventory Value"),
    ("Assessed-value components", "Taxable-value components"),
    ("Assessed-value lag", "Taxable-value lag"),
    ("Vacant Land Value / Residential Value tabs",
     "Builder Lot Inventory Value / Residential Value tabs"),
    ("Colorado Vacant-Land / Nonresidential Taxable Ratios",
     "Utah Builder Lot Inventory Taxable Ratios"),
    ("Vacant Land", "Builder Lot Inventory"),
    ("Vacant-land", "Lot-inventory"),
    ("vacant-land", "lot-inventory"),
    ("vacant land", "builder lot inventory"),
    ("Vacant land", "Builder lot inventory"),
    ("Vacant Lot Value", "Lot Inventory Value"),
    ("Vacant Lot AV", "Lot Inventory\\nTaxable Value"),
    ("Vacant Lot Market Value", "Lot Inventory Market Value"),
    ("Vacant-lot", "Lot-inventory"),
    ("vacant-lot", "lot-inventory"),
    ("vacant lot value", "lot inventory value"),
    ("vacant lots + homes", "lot inventory + homes"),
    ("Historical TABOR Rates", "Utah Property Tax Reference"),
    ("TABOR Residential Taxable Ratio", "Primary Residential Taxable Ratio"),
    ("TABOR Rate at Time of Service Plan", "Prior Residential Taxable Ratio"),
    ("Assessment ratio (TABOR residential)", "Primary residential taxable ratio"),
    ("Mill Levy from Service Plan", "Mill Levy — Governing Document Cap"),
    ("Service Plan Gallagherization Base Rate", "Prior Residential Taxable Ratio"),
    ("Service Plan Mill Levy", "Governing Document Mill Levy Cap"),
    ("Service Plan cap", "governing document cap"),
    ("Service-Plan cap", "governing document cap"),
    ("the Service Plan", "the governing document"),
    ("Service Plan", "Governing Document"),
    ("@ TABOR", "@ Taxable Ratio"),
    ("TABOR gross-up", "taxable-value gross-up"),
    ("(TABOR)", "(Utah)"),
    ("TABOR ratio", "residential taxable ratio"),
    ("TABOR residential", "primary residential"),
    ("TABOR", "Residential Exemption"),
    ("Biennial\\nReassess.", "Reassessment"),
    ("Biennial\\nReassessment", "Annual\\nReassessment"),
    ("+ Biennial\\nReassess.", "+ Reassessment"),
    ("Biennial Reassessment", "Reassessment"),
    ("biennial reassessment", "reassessment"),
    ("(+ biennial reassess.)", "(+ reassessment)"),
    ("biennial reassess.", "reassessment"),
    ("Residential\\nAV", "Residential\\nTaxable Value"),
    ("Lot\\nAV", "Lot Inventory\\nTaxable Value"),
    ("Centrally Assessed\\nAV", "Centrally Assessed\\nTaxable Value"),
    ("Commercial\\nAV", "Commercial\\nTaxable Value"),
    ("Home AV\\n", "Home Taxable Value\\n"),
    ("Lot AV\\n", "Lot Inventory Taxable Value\\n"),
    ("Commercial AV\\n", "Commercial Taxable Value\\n"),
    ("Total AV", "Total Taxable Value"),
    ("lot AV", "lot taxable value"),
    ("home AV", "home taxable value"),
    ("Total Residential AV", "Total Residential Taxable Value"),
    ("AV-Set Year", "Roll Year"),
    ("AV Set", "Roll"),
    ("strict biennial level-of-value", "two-year level-of-value hold"),
    ("specific-ownership-tax", "uniform-fee"),
    ("Colorado Metropolitan District", "Utah Public Infrastructure District"),
    ("Colorado metro-district", "Utah PID"),
    ("Colorado metro district", "Utah PID"),
    ("metro-district", "PID"),
    ("Wolf Creek Run West", "Viridian Farm"),
    ("Wolf Creek", "Viridian Farm"),
    ("Silver Peaks", "Viridian Farm"),
    ("Adams County, CO", "Utah County, Utah"),
    (", Colorado", ", Utah"),
]


# ── Pass 3: semantic patches ─────────────────────────────────────────────────
#
# (old, new) applied in order.  Every entry must match unless it is listed in
# OPTIONAL_PATCHES — a miss means Colorado moved and this needs re-reading.

PATCHES: dict[str, list[tuple[str, str]]] = {}


def _p(module: str, old: str, new: str) -> None:
    PATCHES.setdefault(module, []).append((old, new))


# ── config.py ────────────────────────────────────────────────────────────────

_p("config.py", '''"""
config.py — All named-range inputs from the "Inputs - First" sheet.''',
   '''"""
config.py — All named-range inputs from the "Inputs" sheet.''')

_p("config.py", '''Every value here corresponds to a named range in the Colorado Metropolitan
District financial-analysis workbook (Viridian Farm / Viridian Farm
template).  Named range -> Python attribute mapping is intended to be 1-to-1.

To change an assumption, edit the value here, or pass overrides to
``ModelConfig(...)``.

Utah PID financings have several distinguishing features
that this model captures:

  * Property is assessed at a primary residential taxable ratio (~7.15%)
    rather than a flat 100% of market.
  * Assessed value is re-set on a **reassessment** cycle (every two
    years) which grows existing value by a fixed retaxable ratio.
  * Debt is pledged a **mill levy** (capped by the governing document) plus
    **Personal Property Uniform Fee (SOT)** rather than a tax rate per $100 of AV.
  * Financings are layered into a **senior lien** and a **subordinate /
    cash-flow** lien, and the senior new-money bonds are later **refunded**
    (refinanced) once they become callable — generating additional
    reimbursement ("new money").
"""''',
   '''Utah counterpart of ``co_metro_model/config.py``: same structure, same
attribute names wherever the concept survives, with the Colorado mechanics
replaced by their Utah equivalents.

To change an assumption, edit the value here, or pass overrides to
``ModelConfig(...)``.

What Utah does differently
--------------------------

  * **Taxable value, not an assessment ratio.**  Primary residential property
    is taxed on 55% of fair market value — the 45% primary residential
    exemption (Utah Const. art. XIII, § 3; Utah Code § 59-2-103), covering the
    dwelling and up to one acre of land.  Colorado's TABOR residential ratio
    (~6.7%) has no Utah analogue.
  * **Annual reassessment.**  County assessors revalue every year (§ 59-2-303.1),
    so there is no biennial hold and no odd/even reappraisal phase.
  * **Levy caps are fixed rates, not gallagherized.**  A PID's levy for all
    purposes may not exceed 0.015 per dollar of taxable value — 15.000 mills
    (§ 17D-4-303) — and the governing document and indentures usually set a
    lower rate.  The most restrictive controls, and none of them float with the
    residential exemption.
  * **Taxes are due 30 November in a single payment**, so debt service is
    structured with principal on 1 March rather than Colorado's 1 December.
  * **Personal property uniform fee**, not specific ownership tax (§ 59-2-405),
    distributed to taxing entities in the same proportion as real property tax.
  * **No county treasurer haircut.**  Assessing and collecting is funded by a
    separate statewide levy on property (§ 59-2-1602).
  * **Truth in Taxation, not TABOR.**  A rate above the certified tax rate needs
    notice and a hearing, but the PID Act exempts the required mill levy while
    it stays within the caps above.

Everything else — the layered senior / subordinate lien structure, the
revenue-wrap sizing, the capitalized-interest and debt-service-reserve
mechanics, and the senior refunding that generates "new money" — is the same as
Colorado, and is modeled the same way.

This file is generated by ``tools/port_from_colorado.py``; edit the patch set
there, not here.
"""''')

_p("config.py", '''    """
    Central parameter store.  Defaults mirror the Viridian Farm
    Metropolitan District workbook (Utah County, Utah).

    Override any field to run a what-if scenario.
    """''',
   '''    """
    Central parameter store.  Defaults mirror the Viridian Farm Public
    Infrastructure District No. 1 financing (Salem City, Utah County) priced
    17 September 2024 — the reference deal in this repository.

    Override any field to run a what-if scenario.
    """''')

_p("config.py", '''    pid_name: str = ""                                           # METRO
    city: str = ""                                                  # CITY
    county: str = ""                                               # COUNTY
    developer: str = ""                                            # DEVELOPER''',
   '''    pid_name: str = "Viridian Farm Public Infrastructure District No. 1"  # PID
    city: str = "Salem"                                            # CITY
    county: str = "Utah"                                           # COUNTY
    developer: str = "D.R. Horton"                                 # DEVELOPER''')

_p("config.py", '''    uwd_senior: float = 0.02            # UWD_SENIOR — underwriter's discount
    uwd_sub: float = 0.03               # UWD_SUB
    uwd_senior_refunding: float = 0.005 # UWD_SENIOR_REFUNDING
    county_collection_fee: float = 0.015 # COUNTY_COLLECTION_FEE''',
   '''    uwd_senior: float = 0.01            # UWD_SENIOR — underwriter's discount
    uwd_sub: float = 0.015              # UWD_SUB
    uwd_senior_refunding: float = 0.005 # UWD_SENIOR_REFUNDING
    # Utah funds assessing and collecting through a separate statewide levy on
    # property (§ 59-2-1602) rather than a haircut on the district's
    # distribution, so this is 0.00 where Colorado runs ~1.50%.  The indentures
    # still define pledged revenue net of county collection costs, so the input
    # stays available.
    county_collection_fee: float = 0.0  # COUNTY_COLLECTION_FEE''')

_p("config.py", '''    delivery: date = field(default_factory=lambda: date(2024, 12, 1))   # DELIVERY
    prin_maturity: int = 12             # PRIN_MATURITY (month principal is paid)
    int_maturity: int = 6               # INT_MATURITY  (derived: prin-6)''',
   '''    delivery: date = field(default_factory=lambda: date(2024, 9, 26))   # DELIVERY
    # Utah property taxes are due 30 November in a single payment, so principal
    # falls on 1 March — the first payment date after collections are in hand.
    # (Colorado's Feb/June collections put principal on 1 December.)
    prin_maturity: int = 3              # PRIN_MATURITY (month principal is paid)
    int_maturity: int = 9               # INT_MATURITY  (prin + 6)''')

_p("config.py", '''    capi_term: int = 24                 # CAPI_TERM (months of capitalized interest)''',
   '''    capi_term: int = 36                 # CAPI_TERM (months of capitalized interest)''')

_p("config.py", '''    senior_interest_rate: float = 0.06            # SENIOR_INTEREST_RATE (coupon)
    senior_refunding_interest_rate: float = 0.045 # SENIOR_REFUNDING_INTEREST_RATE
    sub_interest_rate: float = 0.08               # SUB_INTEREST_RATE''',
   '''    senior_interest_rate: float = 0.05875         # SENIOR_INTEREST_RATE (coupon)
    senior_refunding_interest_rate: float = 0.045 # SENIOR_REFUNDING_INTEREST_RATE
    sub_interest_rate: float = 0.08125            # SUB_INTEREST_RATE''')

_p("config.py", '''    dsc_senior: float = 1.25            # DSC_SENIOR_LIEN_BONDS''',
   '''    dsc_senior: float = 1.30            # DSC_SENIOR_LIEN_BONDS''')

_p("config.py", '''    interest_earn_rate: float = 0.01    # INTEREST_EARN_RATE (on DSRF / surplus)''',
   '''    interest_earn_rate: float = 0.025   # INTEREST_EARN_RATE (on DSRF / surplus)''')

_p("config.py", '''    first_year: int = 2023              # FIRST_YEAR (first "Summary" year)''',
   '''    first_year: int = 2022              # FIRST_YEAR (first "Summary" year)''')

_p("config.py", '''    inflation_start_year: int = 2024    # INFLATION_START_YEAR (home-price inflation''',
   '''    inflation_start_year: int = 2025    # INFLATION_START_YEAR (home-price inflation''')

_p("config.py", '''    reassess_rate: float = 0.02          # REASSESS_RATE (residential)
    reassess_comm_rate: float = 0.02     # REASSESS_COMM_RATE
    # Series C cash-flow bonds — sized against a SEPARATE assessment that biennially
    # reassesses the created AV at this rate (used only to size the Series C bonds).
    series_c_reassess_rate: float = 0.06  # SERIES_C_REASSESS_RATE
    size_series_c: str = "No"                      # SIZE_SERIES_C — Yes/No toggle
    series_c_interest_rate: float = 0.08           # SERIES_C_INTEREST_RATE (accreting)
    dsc_series_c: float = 1.0                      # DSC_SERIES_C — coverage
    series_c_final_mat_yrs: int = 40               # SERIES_C_FINAL_MAT_YRS (from delivery)
    # Developer cash contribution — a source in the first-financing Sources & Uses,
    # applied to a chosen series (Senior / Subordinate / Series C / Proportional).
    developer_contribution: float = 0.0            # DEVELOPER_CONTRIBUTION ($)
    developer_contribution_series: str = "Proportional"  # DEVELOPER_CONTRIBUTION_SERIES

    resid_taxable_ratio: float = 0.067        # RESID_TAXABLE_RATIO — residential taxable ratio
    resid_taxable_ratio_prior: float = 0.0715  # RESID_TAXABLE_RATIO_PRIOR — residential taxable ratio at time of service plan
    tax_collect_mill_prc: float = 0.995 # TAX_COLLECT_MILL_PRC
    uniform_fee_prc: float = 0.07    # UNIFORM_FEE_PRC (SOT as % of mill revenue)
    uniform_fee_av_threshold: float = 0  # UNIFORM_FEE_AV_THRESHOLD

    gallagherization: str = "Yes"       # GALLAGHERIZATION

    # ── Mill levies ──────────────────────────────────────────────────────────
    mill_levy_governing_doc: float = 50  # MILL_LEVY_GOVERNING_DOC
    mill_levy_ds_target: float = 50     # MILL_LEVY_DS_TARGET
    mill_levy_comm: Optional[float] = None  # MILL_LEVY_COMM (None ⇒ same as DS mill)
    mill_levy_ops_target: float = 5     # MILL_LEVY_OPS_TARGET

    # ── Lot / home valuation ratios ──────────────────────────────────────────
    lot_inventory_taxable_ratio: float = 0.29   # LOT_INVENTORY_TAXABLE_RATIO (% of market, dev. lots)''',
   '''    # Utah counties revalue ANNUALLY, so this growth is applied every year
    # rather than on Colorado's two-year cycle.
    reassess_rate: float = 0.01          # REASSESS_RATE (residential)
    reassess_comm_rate: float = 0.02     # REASSESS_COMM_RATE
    # "Annual" (Utah, § 59-2-303.1) or "Biennial" (the Colorado cadence).
    reassess_frequency: str = "Annual"   # REASSESS_FREQUENCY

    # ── Series C cash-flow bonds ─────────────────────────────────────────────
    # Colorado sizes an optional third lien against a SEPARATE assessment that
    # reassesses the created taxable value at its own rate.  Carried here inert:
    # the Utah authority for a second assessment on the same property has not
    # been worked through, so the toggle is off and the rows are kept off the
    # Inputs template rather than offered untested.
    series_c_reassess_rate: float = 0.06  # SERIES_C_REASSESS_RATE
    size_series_c: str = "No"                      # SIZE_SERIES_C — Yes/No toggle
    series_c_interest_rate: float = 0.08           # SERIES_C_INTEREST_RATE (accreting)
    dsc_series_c: float = 1.0                      # DSC_SERIES_C — coverage
    series_c_final_mat_yrs: int = 40               # SERIES_C_FINAL_MAT_YRS (from delivery)
    # Developer cash contribution — a source in the first-financing Sources &
    # Uses, applied to a chosen series (Senior / Subordinate / Proportional).
    developer_contribution: float = 0.0            # DEVELOPER_CONTRIBUTION ($)
    developer_contribution_series: str = "Proportional"  # DEVELOPER_CONTRIBUTION_SERIES

    # Primary residential exemption: taxable value is 55% of fair market value
    # (Utah Const. art. XIII, § 3; § 59-2-103).
    resid_taxable_ratio: float = 0.55         # RESID_TAXABLE_RATIO
    # Prior-period ratio, retained so a change in the exemption can be modeled;
    # unlike Colorado it does not adjust the mill levy.
    resid_taxable_ratio_prior: float = 0.55   # RESID_TAXABLE_RATIO_PRIOR
    tax_collect_mill_prc: float = 0.98  # TAX_COLLECT_MILL_PRC
    # Personal property uniform fee (§ 59-2-405) as a % of mill-levy revenue.
    uniform_fee_prc: float = 0.0        # UNIFORM_FEE_PRC
    uniform_fee_av_threshold: float = 0  # UNIFORM_FEE_AV_THRESHOLD

    # ── Mill levies ──────────────────────────────────────────────────────────
    # Rate per $1,000 of taxable value.  0.003 per dollar = 3.000 mills.
    mill_levy_governing_doc: float = 5.0  # MILL_LEVY_GOVERNING_DOC (governing document cap)
    mill_levy_indenture: Optional[float] = 3.0  # MILL_LEVY_INDENTURE (indenture cap)
    mill_levy_ds_target: float = 3.0    # MILL_LEVY_DS_TARGET
    mill_levy_comm: Optional[float] = None  # MILL_LEVY_COMM (None ⇒ same as DS mill)
    mill_levy_ops_target: float = 0     # MILL_LEVY_OPS_TARGET

    # ── Lot / home valuation ratios ──────────────────────────────────────────
    # Builder lot inventory.  Utah Admin. Code R884-24P-52 lets the residential
    # exemption reach unoccupied property (and property under construction) the
    # assessor determines will be a primary residence once occupied, so finished
    # lots are carried at the same 55%.  Set to 1.00 to tax inventory at full
    # fair market value instead (the conservative reading).
    lot_inventory_taxable_ratio: float = 0.55   # LOT_INVENTORY_TAXABLE_RATIO''')

_p("config.py", '''    om_expense: float = 0              # OM_EXPENSE
    om_expense_av_limit: float = 0     # OM_EXPENSE_AV_LIMIT
    om_growth_rate: float = 0.02        # OM_GROWTH_RATE''',
   '''    # District operations & maintenance — the single line for what the district
    # spends each year: administration (accounting, audit, legal, assessor and
    # continuing-disclosure filings) and operations (landscaping, parks and
    # trails, snow removal, street lighting, utilities on the improvements).
    # Colorado carves this out of pledged revenue against a separate operations
    # levy; a Utah PID rarely has one, so it is paid from the same revenue that
    # services the bonds and comes off the top — netted from the revenue
    # available to the senior AND the subordinate lien (see om_expense_for).
    om_expense: float = 53_060         # OM_EXPENSE (base year, $ per year)
    om_expense_av_limit: float = 0     # OM_EXPENSE_AV_LIMIT (0 ⇒ no limit)
    om_growth_rate: float = 0.02       # OM_GROWTH_RATE (inflates the base)
    # First collection year that carries district costs — O&M and the trustee
    # fees.  None ⇒ two years after closing: the first roll set with the bonds
    # outstanding is billed that November, so year 2 is the first with a full
    # year of collections to charge against.
    district_cost_start_year: Optional[int] = None   # DISTRICT_COST_START_YEAR''')

# Levy caps replace gallagherization.
_p("config.py", '''    @property
    def effective_ds_mill_levy(self) -> float:
        """
        Debt-service mill levy applied to taxable value.

        When ``gallagherization == "Yes"`` the Service-Plan mill levy is
        adjusted ("Gallagherized") to offset the change in the residential
        taxable ratio since the governing document was adopted:

            mill_levy_governing_doc / (resid_taxable_ratio / resid_taxable_ratio_prior)

        i.e. as the current residential taxable ratio falls relative to the ratio in effect
        when the governing document was adopted, the mill levy is grossed up to hold
        revenue.  Otherwise the DS target mill levy is used.
        """
        if (self.gallagherization == "Yes"
                and self.resid_taxable_ratio and self.resid_taxable_ratio_prior):
            return self.mill_levy_governing_doc / (self.resid_taxable_ratio / self.resid_taxable_ratio_prior)
        return self.mill_levy_ds_target

    @property
    def commercial_mill_levy(self) -> float:
        """
        Commercial debt-service mill levy.  Defaults to the (Gallagherized, if
        enabled) DS mill levy when not separately set, so the two start equal;
        override ``mill_levy_comm`` to tax commercial at a different rate.
        """
        return self.mill_levy_comm if self.mill_levy_comm is not None else self.effective_ds_mill_levy''',
   '''    def district_costs(self, collection_year: int) -> tuple[float, float]:
        """
        (senior trustee fee, subordinate trustee fee) charged against pledged
        revenue in ``collection_year``.  Nothing is charged before
        ``district_cost_start_year``; the fees are flat thereafter.

        District O&M is the other standing cost — see ``om_expense_for``, which
        is netted from both liens rather than the senior alone.
        """
        start = self.district_cost_start_year or (self.delivery.year + 2)
        if collection_year < start:
            return 0.0, 0.0
        return self.trustee_fee, self.trustee_fee_sub

    def om_expense_for(self, collection_year: int,
                       total_av: float | None = None) -> float:
        """
        District operations & maintenance charged against pledged revenue in
        ``collection_year`` — the single line covering district administration
        and operations alike.

        Nothing is charged before ``district_cost_start_year`` (the same start
        the trustee fees use); from then on the base inflates at
        ``om_growth_rate``.  If ``om_expense_av_limit`` is set and ``total_av``
        is above it, the charge stops — the district is assumed to fund itself
        from an operations levy once the base is large enough.
        """
        if not self.om_expense:
            return 0.0
        if (self.om_expense_av_limit and total_av is not None
                and total_av > self.om_expense_av_limit):
            return 0.0
        start = self.district_cost_start_year or (self.delivery.year + 2)
        if collection_year < start:
            return 0.0
        return self.om_expense * (1 + self.om_growth_rate) ** (collection_year - start)

    @property
    def mill_levy_cap(self) -> float:
        """
        The controlling levy cap, in mills — the most restrictive of the three
        that bind a Utah PID:

          1. **Statute** — § 17D-4-303 caps the levy for all purposes at 0.015
             per dollar of taxable value (15.000 mills).  It does not bind a levy
             to pay principal of and interest on a voted general obligation bond.
          2. **Governing document** — the rate fixed when the district was created.
          3. **Indentures** — the rate the district covenants to levy for the
             bonds, usually the tightest of the three.

        Unlike a Colorado service-plan cap this is a fixed rate per dollar of
        taxable value: it does not float with the residential exemption, so
        there is no gallagherization step.
        """
        caps = [PID_STATUTORY_LEVY_CAP * 1000.0, self.mill_levy_governing_doc]
        if self.mill_levy_indenture is not None:
            caps.append(self.mill_levy_indenture)
        return min(c for c in caps if c and c > 0)

    @property
    def effective_ds_mill_levy(self) -> float:
        """
        Debt-service mill levy actually applied to taxable value — the target
        rate, held down to the controlling cap.
        """
        return min(self.mill_levy_ds_target, self.mill_levy_cap)

    @property
    def commercial_mill_levy(self) -> float:
        """
        Commercial debt-service mill levy.  Defaults to the DS mill levy when not
        separately set, so the two start equal; override ``mill_levy_comm`` to
        tax commercial at a different rate.
        """
        return self.mill_levy_comm if self.mill_levy_comm is not None else self.effective_ds_mill_levy

    # ── Statutory checks ─────────────────────────────────────────────────────
    def validate(self) -> list[str]:
        """Statutory / structural warnings for this configuration (empty is good)."""
        out: list[str] = []
        statutory = PID_STATUTORY_LEVY_CAP * 1000.0
        total = self.mill_levy_ds_target + self.mill_levy_ops_target
        if total > statutory + 1e-9:
            out.append(
                f"Total district levy of {total:.3f} mills exceeds the "
                f"§ 17D-4-303 cap of {statutory:.3f} mills (0.015 per dollar of "
                f"taxable value).")
        if self.mill_levy_ds_target > self.mill_levy_cap + 1e-9:
            out.append(
                f"Debt service levy of {self.mill_levy_ds_target:.3f} mills exceeds "
                f"the controlling cap of {self.mill_levy_cap:.3f} mills; the model "
                f"levies the cap.")
        if not 0 < self.resid_taxable_ratio <= 1:
            out.append("Residential taxable ratio must be between 0 and 1 "
                       "(Utah taxes 55% of fair market value, not ~6.7%).")
        if self.lot_inventory_taxable_ratio > 1:
            out.append("Builder lot inventory cannot be taxed above 100% of fair "
                       "market value.")
        if self.reassess_frequency.strip().lower() not in ("annual", "biennial"):
            out.append('Reassessment frequency must be "Annual" or "Biennial".')
        return out''')

# Call / CAPI dates snap to a payment date (Colorado delivers on its payment date).
_p("config.py", '''    @property
    def capi_end_date(self) -> date:
        """CAPI_END_DATE — last capitalized-interest payment date."""
        y = self.delivery.year + self.capi_term // 12
        m = self.delivery.month
        return date(y, m, self.prin_maturity_day_senior)''',
   '''    def _snap_to_payment_date(self, d: date) -> date:
        """
        Snap a date back to the most recent principal payment date.

        Colorado delivers on 1 December and pays principal on 1 December, so
        this is a no-op there.  Utah delivers when the market allows and pays on
        1 March, so call dates and the capitalized-interest end date have to land
        on a payment date rather than an anniversary of closing.
        """
        candidate = date(d.year, self.prin_maturity, self.prin_maturity_day_senior)
        if candidate > d:
            candidate = date(d.year - 1, self.prin_maturity,
                             self.prin_maturity_day_senior)
        return candidate

    @property
    def capi_end_date(self) -> date:
        """CAPI_END_DATE — last capitalized-interest payment date."""
        return self._snap_to_payment_date(_edate(self.delivery, self.capi_term))''')

_p("config.py", '''        delivery date moves the first call date with it."""
        return _edate(self.delivery, 12 * self.premium_call_years)''',
   '''        delivery date moves the first call date with it.  The date is snapped to
        a principal payment date — bonds are redeemed on a payment date, not on
        an anniversary of closing."""
        return self._snap_to_payment_date(
            _edate(self.delivery, 12 * self.premium_call_years))''')

_p("config.py", '''        if self.hold_value_flat:
            # Hold value flat across the two-year cycle, anchored to the
            # reassessment-phase year (even by default).''',
   '''        if self.hold_value_flat:
            # Optional Colorado-style two-year hold; off for Utah, which
            # revalues annually.''')

_p("config.py", '''    # Colorado reassesses real property on a two-year cycle.  Value is set as of
    # the June-30 appraisal date in the year BEFORE the reappraisal year, held
    # flat across the two-year cycle, and the resulting taxes are collected the
    # following year — so the taxable value backing a year's mill-levy revenue
    # reflects market value from roughly two years earlier.  This lag is applied
    # in summary.py: residential AV for collection year C uses cumulative market
    # value from year (C - av_lag_years).''',
   '''    # Value is set as of 1 January, appears on that year's roll, and the taxes
    # are due 30 November of the SAME year — funding the following 1 March debt
    # service payment.  So value created during calendar year V lands on the roll
    # for V+1, is collected in V+1, and pays debt service in V+2: a two-year lag
    # from value creation to the debt service it supports.
    #
    # Colorado reaches the same two-year lag by a different route (value set from
    # a June-30 level of value the year before a biennial reappraisal, collected
    # the year AFTER the roll).  Same number of years; different mechanism, and
    # different behaviour in between — Utah revalues every year, so there is no
    # two-year hold.''')

# Replace the whole Colorado rate-history tail with Utah's.
_p("config.py", "@@TAIL@@", '''# ── Utah statutory constants ─────────────────────────────────────────────────

#: § 17D-4-303 — a public infrastructure district's property tax levy, for all
#: purposes including debt service on limited tax bonds, may not exceed this rate
#: per dollar of taxable value.  It does not bind a levy to pay principal of and
#: interest on a voted general obligation bond the district issues.
PID_STATUTORY_LEVY_CAP = 0.015          # = 15.000 mills

#: Utah Const. art. XIII, § 3 / § 59-2-103 — the primary residential exemption.
RESIDENTIAL_EXEMPTION = 0.45
DEFAULT_RESID_TAXABLE_RATIO = 1.0 - RESIDENTIAL_EXEMPTION   # 0.55


# ── Utah primary residential exemption history ───────────────────────────────
# Taxable share of fair market value for primary residential property, by tax
# year.  Utah voters added the exemption to the constitution in 1982 at 25% and
# the legislature stepped it up to the constitutional maximum of 45% effective
# 1 January 1995, where it has stayed.  Powers the "Utah Property Tax Reference"
# sheet and lets a certified taxable value be grossed up to market value
# (taxable ÷ ratio) for the year it was certified.
RESIDENTIAL_EXEMPTION_HISTORY = [
    ("1982–1984", 0.7500, "Exemption added by constitutional amendment at 25%"),
    ("1985–1994", 0.6800, "Stepped up in increments toward the 45% maximum"),
    ("1995 +",    0.5500, "45% exemption — the constitutional maximum "
                          "(art. XIII, § 3; § 59-2-103)"),
]


def _rate_year_bounds(label: str) -> tuple[int, int]:
    """(low, high) tax-year bounds for a history label ('1985–1994', '1995 +')."""
    open_ended = "+" in label
    s = label.replace("+", "").strip()
    for dash in ("–", "—", "-"):
        if dash in s:
            lo, hi = s.split(dash)
            return int(lo.strip()), int(hi.strip())
    y = int(s)
    return (y, 9999) if open_ended else (y, y)


def residential_taxable_ratio_for(year: int) -> float:
    """Primary residential taxable ratio in effect for ``year``."""
    fallback = (RESIDENTIAL_EXEMPTION_HISTORY[-1][1]
                if RESIDENTIAL_EXEMPTION_HISTORY else DEFAULT_RESID_TAXABLE_RATIO)
    for label, ratio, _ in RESIDENTIAL_EXEMPTION_HISTORY:
        lo, hi = _rate_year_bounds(label)
        if lo <= year <= hi:
            return ratio
    return fallback


# ── Builder lot inventory taxable ratio ──────────────────────────────────────
# Utah has no separate vacant-land assessment class: non-exempt property is
# taxed on 100% of fair market value.  What matters for a residential PID is
# whether finished lots and homes under construction, still owned by the
# builder, carry the primary residential exemption.  Utah Admin. Code
# R884-24P-52 says they can: on a written declaration, or where the assessor
# determines the property will qualify as a primary residence once occupied,
# the exemption applies while the property is unoccupied.  The default below
# follows that treatment (55%).  Enter 1.0000 for a roll year in the "Utah
# Property Tax Reference" tab to tax inventory at full market value instead.
#
# Keyed by ASSESSMENT ROLL / TAX YEAR.
BUILDER_INVENTORY_HISTORY = [
    ("1995 +", 0.5500, "Residential exemption applied to unoccupied builder "
                       "inventory (Utah Admin. Code R884-24P-52)"),
]


def lot_inventory_ratio(year: int) -> float:
    """Builder lot inventory taxable ratio for a roll/tax ``year``."""
    fallback = (BUILDER_INVENTORY_HISTORY[-1][1]
                if BUILDER_INVENTORY_HISTORY else DEFAULT_RESID_TAXABLE_RATIO)
    for label, ratio, _ in BUILDER_INVENTORY_HISTORY:
        lo, hi = _rate_year_bounds(label)
        if lo <= year <= hi:
            return ratio
    return fallback


# ── Utah property tax calendar (reference) ───────────────────────────────────
# Drives the reference sheet in the inputs workbook and the memo prose.  Dates
# are from Utah Code title 59, chapter 2, part 3 and part 13.
UTAH_TAX_CALENDAR = [
    ("January 1", "Lien date — property is valued as of this date"),
    ("May 1", "State Tax Commission assesses centrally assessed property"),
    ("May 22", "County assessors complete locally assessed property"),
    ("June 8", "Centrally assessed value apportioned to taxing entities"),
    ("June 22", "Taxing entity adopts a proposed or final tax rate"),
    ("July 22", "Valuation notices mailed; Truth in Taxation hearing noticed"),
    ("November 1", "Corrected rolls delivered; tax notices mailed"),
    ("November 30", "Taxes due — a single annual payment"),
    ("December 31", "Delinquency; penalty of the greater of 2.5% or $10"),
    ("January 1 (next)", "Interest runs at the federal funds target + 6%, "
                         "floored at 7% and capped at 10%"),
]
''')

#: Marker: everything from this line to the end of config.py is replaced.
CONFIG_TAIL_MARKER = "# ── Historical Colorado residential (Utah) taxable ratios ─"


_p("config.py", '''        Precedence: an explicit collection-year ``residential_assessment_schedule``
        wins; then the editable ROLL-year ``residential_rate_schedule`` read back
        from the "Utah Property Tax Reference" tab (so the template drives the rates —
        e.g. 2026+ = 6.80% per HB24B-1001); otherwise the built-in RESIDENTIAL_EXEMPTION_HISTORY
        rate for the ROLL year (= collection year − 1) applies.
        """
        if self.residential_assessment_schedule:
            return _schedule_lookup(
                self.residential_assessment_schedule, collection_year, self.resid_taxable_ratio)
        if self.residential_rate_schedule:
            return _schedule_lookup(
                self.residential_rate_schedule, collection_year - 1, self.resid_taxable_ratio)
        return residential_taxable_ratio_for(collection_year - 1)''',
   '''        Precedence: an explicit collection-year ``residential_assessment_schedule``
        wins; then the Inputs-page ``resid_taxable_ratio`` whenever it differs from
        the statutory 45% exemption (so overriding that cell actually changes the
        model); then the editable ROLL-year ``residential_rate_schedule`` read back
        from the "Utah Property Tax Reference" tab; otherwise the statutory table.

        Utah's exemption has been flat at 55% since 1995, so all four agree unless
        someone deliberately changes one.  Colorado needs the table on top because
        its ratio moves every reappraisal cycle.
        """
        if self.residential_assessment_schedule:
            return _schedule_lookup(
                self.residential_assessment_schedule, collection_year, self.resid_taxable_ratio)
        if abs(self.resid_taxable_ratio - DEFAULT_RESID_TAXABLE_RATIO) > 1e-12:
            return self.resid_taxable_ratio
        if self.residential_rate_schedule:
            return _schedule_lookup(
                self.residential_rate_schedule, collection_year - 1, self.resid_taxable_ratio)
        return residential_taxable_ratio_for(collection_year - 1)''')


# ── development.py ───────────────────────────────────────────────────────────

_p("development.py", '''  * **Aggregate** (default) — a single absorption stream via the dict fields
    below (reproduces the Viridian Farm program: 327 SFD lots).''',
   '''  * **Aggregate** (default) — a single absorption stream via the dict fields
    below (reproduces the Viridian Farm program: 716 residential units).''')

_p("development.py", '''    base_asp: float = 515_000
    asp_base_year: int = 2024

    # Aggregate drivers (used when ``products`` is empty).
    home_closings: dict[int, int] = field(default_factory=lambda: {
        2024: 48, 2025: 96, 2026: 96, 2027: 87,
    })
    lot_market_value: dict[int, float] = field(default_factory=lambda: {
        2024: 3_000_000, 2025: 9_150_000, 2026: 4_350_000,
    })
    lot_deliveries: dict[int, int] = field(default_factory=lambda: {
        2024: 108, 2025: 219,
    })''',
   '''    base_asp: float = 449_000
    asp_base_year: int = 2024

    # Aggregate drivers (used when ``products`` is empty) — the Viridian Farm
    # PID No. 1 program: 716 units closing 2024-2029, lots delivered a year ahead.
    home_closings: dict[int, int] = field(default_factory=lambda: {
        2024: 4, 2025: 172, 2026: 202, 2027: 151, 2028: 96, 2029: 91,
    })
    lot_deliveries: dict[int, int] = field(default_factory=lambda: {
        2023: 4, 2024: 172, 2025: 202, 2026: 151, 2027: 96, 2028: 91,
    })
    lot_market_value: dict[int, float] = field(default_factory=lambda: {
        2023: 4 * 44_900.0, 2024: 172 * 44_900.0, 2025: 202 * 44_900.0,
        2026: 151 * 44_900.0, 2027: 96 * 44_900.0, 2028: 91 * 44_900.0,
    })''')

_p("development.py", '''    total_lots: int = 327
    pace_factor: float = 1.0''',
   '''    total_lots: int = 716
    pace_factor: float = 1.0''')

_p("development.py", '''    Aggregate defaults match Viridian Farm (327 SFD lots, base ASP
    $515,000 (2024) appreciating at the inflation rate).  Supply ``products``
    to drive the model from per-builder/per-product detail instead.''',
   '''    Aggregate defaults match Viridian Farm PID No. 1 (716 units, weighted-average
    base ASP $449,000 (2024) appreciating at the inflation rate).  Supply
    ``products`` to drive the model from per-builder/per-product detail instead.''')

_p("development.py", '''        # (> $100k), gross it up to MARKET value using the Historical residential taxable ratio
        # for the certification year (assessed ÷ ratio).''',
   '''        # (> $100k), gross it up to MARKET value using the primary residential
        # taxable ratio for the certification year (taxable ÷ ratio).''')

# Annual reassessment.
_p("development.py", '''        Colorado reappraises property on a two-year cycle, in the ODD assessment
        ROLL years (2025, 2027, ...).  The reappraised level of value is fixed as
        of June 30 of the prior (even) year, so in this model — where ``year`` is
        the level-of-value / market-value index that backs the roll of ``year + 1``
        (and is collected in ``year + 2``) — the reassessment step is booked in the
        even level-of-value year that feeds the odd roll year.  Parity is therefore
        evaluated on the ROLL year (``year + 1``), NOT on the level-of-value index
        or the tax-collection year.

        No reassessment is applied on the certified roll — the county-certified
        value already reflects the actual roll for that year, so the model should
        not layer a reassessment on top of it.
        """
        if year < cfg.first_year + 2:
            return False
        roll_year = year + 1
        cert_date = getattr(cfg, "certification_date", None)
        if cert_date is not None and roll_year == cert_date.year:
            return False
        return (roll_year % 2 == 0) if getattr(cfg, "reassess_on_even_years", True) else (roll_year % 2 == 1)''',
   '''        Utah county assessors update values **annually** based on a systematic
        review of current market data (§ 59-2-303.1), with a detailed review of
        each parcel at least every five years — so the answer is normally "every
        year".  Set ``reassess_frequency = "Biennial"`` to model Colorado's
        two-year reappraisal cycle instead, in which case the step is booked on
        the level-of-value year that feeds the odd (or even) roll year.

        No growth is applied on a certified roll: the county-certified value
        already IS the roll for that year, so the model must not layer a
        reassessment on top of it.
        """
        if year < cfg.first_year + 2:
            return False
        roll_year = year + 1
        cert_date = getattr(cfg, "certification_date", None)
        if cert_date is not None and roll_year == cert_date.year:
            return False
        if getattr(cfg, "reassess_frequency", "Annual").strip().lower() != "biennial":
            return True
        return ((roll_year % 2 == 0) if getattr(cfg, "reassess_on_even_years", False)
                else (roll_year % 2 == 1))''')

# The residential value build reassesses annually too.
_p("development.py", '''        # Biennial reassessment occurs only in ODD years of the AV set (the value
        # build is keyed by AV-set/roll year directly, so no level-of-value lag),
        # suppressed in the certification year and before the first collectible roll.
        def reassess_on(av):
            if av < first + 2:
                return False
            if cfg.certification_date is not None and av == cfg.certification_date.year:
                return False
            return (av % 2 == 0) if getattr(cfg, "reassess_on_even_years", False) else (av % 2 == 1)''',
   '''        # Utah revalues annually, so value already on the roll grows every year
        # (§ 59-2-303.1) — suppressed in the certification year and before the
        # first collectible roll.  With ``reassess_frequency = "Biennial"`` the
        # step falls only on odd (or even) roll years, the Colorado cadence.
        def reassess_on(av):
            if av < first + 2:
                return False
            if cfg.certification_date is not None and av == cfg.certification_date.year:
                return False
            if getattr(cfg, "reassess_frequency", "Annual").strip().lower() != "biennial":
                return True
            return (av % 2 == 0) if getattr(cfg, "reassess_on_even_years", False) else (av % 2 == 1)''')


# ── summary.py ───────────────────────────────────────────────────────────────

_p("summary.py", '''  AI — district debt-service mill-levy collections
  AJ — specific-ownership tax collected''',
   '''  AI — district debt-service mill-levy collections
  AJ — personal property uniform fee allocated to the district (§ 59-2-405)''')

_p("summary.py", '''All series here are keyed by **collection year** (the year revenue is received,
= the year the matching debt-service payment is due on December 1).''',
   '''All series here are keyed by **collection year** — the year the matching debt
service is due.  In Utah the roll is set 1 January of the prior year and the
taxes are due 30 November of that prior year, so the money is in hand before the
1 March payment it supports.''')

_p("summary.py", '''            # Net revenue available for SENIOR lien debt service (AX):
            #   mill + Uniform Fee - county collection cost - senior trustee fee - O&M carveout
            collection_fee = mill_revenue * cfg.county_collection_fee
            net_senior_revenue = (
                mill_revenue + uniform_fee - collection_fee - cfg.trustee_fee - cfg.om_expense
            )''',
   '''            # Net revenue available for SENIOR lien debt service (AX):
            #   mill + uniform fee - county collection cost - senior trustee fee
            #   - district O&M (inflated, and not charged before the district is
            #     up and running).  O&M is netted from the SUBORDINATE side too —
            #     see ModelConfig.om_expense_for.
            collection_fee = mill_revenue * cfg.county_collection_fee
            trustee_fee, trustee_fee_sub = cfg.district_costs(collect)
            om_expense = cfg.om_expense_for(collect, total_av)
            net_senior_revenue = (
                mill_revenue + uniform_fee - collection_fee - trustee_fee - om_expense
            )''')

_p("summary.py", '''            net_sub_revenue = mill_revenue + uniform_fee - cfg.trustee_fee_sub''',
   '''            net_sub_revenue = mill_revenue + uniform_fee - trustee_fee_sub - om_expense''')


# ── subordinate.py ───────────────────────────────────────────────────────────

# The par-rounding fix that used to live here (round → floor, so the solved par
# never lands above the largest amount the residual cashflow retires) was adopted
# upstream in ba72e6b.  Nothing to patch.


# Upstream loops collection years to ``dev.last_year`` (2067) while the value
# builds stop at the senior final maturity, so the Summary tail shows a decade of
# zero taxable value and negative revenue.  Bound the loop to the builds.  No
# effect on sizing (nothing is sized past final maturity) — worth pushing back to
# co_metro_model.
_p("summary.py", '''        # Collection years run from first_year+2 (first AV actually collected)
        for collect in range(cfg.first_year + 2, dev.last_year + 1):''',
   '''        # Collection years run from first_year+2 (the first roll actually
        # collected) through the last year the value builds cover — past the
        # senior final maturity there is no build to read, and a tail of zero
        # taxable value with fee-only negative revenue is not a projection.
        last_collect = min(dev.last_year,
                           max(vac_build) if vac_build else dev.last_year,
                           max(res_build) if res_build else dev.last_year)
        for collect in range(cfg.first_year + 2, last_collect + 1):''')


# Upstream measures the first (stub) coupon period only when the payment date
# falls in the SAME calendar year as the dated date.  That holds in Colorado
# (dated 1 December, sub pays 15 December) but not in Utah, where a September
# delivery's first sub payment is the following 15 March — upstream charges that
# a full year of interest against ~5.6 months of elapsed time.  Anchor the stub
# to the first payment date AFTER dating instead, which is correct in both
# states.  Worth pushing back to co_metro_model.
_p("subordinate.py", '''            pay_date = date(y, cfg.prin_maturity, cfg.prin_maturity_day_sub)
            if y < dated.year:
                year_frac = 0.0
            elif y == dated.year:
                year_frac = max(0.0, _yearfrac_30360(dated, pay_date))
            else:
                year_frac = 1.0''',
   '''            pay_date = date(y, cfg.prin_maturity, cfg.prin_maturity_day_sub)
            prior_pay = date(y - 1, cfg.prin_maturity, cfg.prin_maturity_day_sub)
            if pay_date <= dated:
                year_frac = 0.0                       # bonds not yet dated
            elif prior_pay <= dated:
                # First payment after dating — a stub, however many months it is.
                year_frac = max(0.0, _yearfrac_30360(dated, pay_date))
            else:
                year_frac = 1.0''')


# ── report.py ────────────────────────────────────────────────────────────────

_p("report.py", '''                "net revenue for sizing.  Lot-inventory and home taxable value flow from the Vacant "
                "Land Value / Residential Value tabs."], last_col)''',
   '''                "net revenue for sizing.  Lot-inventory and home taxable value flow from "
                "the Builder Lot Inventory Value / Residential Value tabs."], last_col)''')

_p("report.py", '''    base = f"{cfg.county} County, Utah" if cfg.county else "Colorado"''',
   '''    base = f"{cfg.county} County, Utah" if cfg.county else "Utah"''')

_p("report.py", '''        trust = -cfg.trustee_fee if r.total_av > 0 else 0.0
        subtrust = -cfg.trustee_fee_sub if r.total_av > 0 else 0.0
        om = -cfg.om_expense if cfg.om_expense else 0.0''',
   '''        # District costs come from the same helpers the revenue engine uses,
        # so the report and the sizing can never disagree.
        _trustee, _subtrustee = cfg.district_costs(r.collection_year)
        trust = -_trustee
        subtrust = -_subtrustee
        om = -r.om_expense''')

_p("report.py", '''        ("om", "− O&M\\nCarveout", 11),''',
   '''        ("om", "− District\\nO&M", 11),''')

_p("report.py", '''                 + (" (adjusted)" if cfg.gallagherization == "Yes" else ""),''',
   '''                 + f" (cap {cfg.mill_levy_cap:.3f})",''')


# ── residential_report.py ────────────────────────────────────────────────────

_p("residential_report.py", '''    # Builder lot inventory is a NONRESIDENTIAL subclass — fixed at 29% under Gallagher,
    # phasing to 25% under SB24-233.  Show the rate the model applies to vacant-
    # lot value by roll year (see the "Utah Property Tax Reference" reference tab),
    # with the residential ratio alongside for comparison (held flat at input).''',
   '''    # Utah has no separate lot-inventory class: what matters is whether the
    # primary residential exemption reaches builder-held inventory (Utah Admin.
    # Code R884-24P-52 says it can).  Show the ratio the model applies to''')

_p("residential_report.py", '''        "Builder lot inventory is nonresidential: fixed at 29% under Gallagher, phasing to 25% under "
        "SB24-233. The model applies this rate to lot-inventory value by roll year (source: the "
        "Utah Property Tax Reference tab). Residential shown for comparison (held at the input ratio)."))''',
   '''        "Utah taxes non-exempt property at 100% of fair market value; builder lot inventory "
        "carries the 45% primary residential exemption where the assessor determines the "
        "property will be a primary residence once occupied (Utah Admin. Code R884-24P-52), "
        "so it is taxed at the same 55%. Source: the Utah Property Tax Reference tab."))''')


# ── inputs.py ────────────────────────────────────────────────────────────────

_p("inputs.py", '''    ("Tax & Valuation", "Adjust the Rate", "GALLAGHERIZATION", "gallagherization", "yesno", "Yes/No"),''',
   '''    ("Tax & Valuation", "Reassessment Frequency", "REASSESS_FREQUENCY", "reassess_frequency", "text", "Annual (Utah, § 59-2-303.1) or Biennial (Colorado cadence)"),''')

_p("inputs.py", '''    ("Mill Levies", "Mill Levy — Governing Document Cap", "MILL_LEVY_GOVERNING_DOC", "mill_levy_governing_doc", "float", "mills"),''',
   '''    ("Mill Levies", "Mill Levy — Governing Document Cap", "MILL_LEVY_GOVERNING_DOC", "mill_levy_governing_doc", "float", "mills"),
    ("Mill Levies", "Mill Levy — Indenture Cap", "MILL_LEVY_INDENTURE", "mill_levy_indenture", "float", "mills; blank ⇒ no indenture cap"),''')

_p("inputs.py", '''    ("Centrally Assessed / Commercial", "Centrally Assessed Property (market value)", "CENTRALLY_ASSESSED_VALUE", "centrally_assessed_value", "float", "$ actual"),
    ("Centrally Assessed / Commercial", "Centrally Assessed Equipment (market value)", "CENTRALLY_ASSESSED_EQUIPMENT", "centrally_assessed_equipment", "float", "$ actual"),
    ("Centrally Assessed / Commercial", "Centrally Assessed Taxable Ratio", "CENTRALLY_ASSESSED_RATIO", "centrally_assessed_ratio", "pct", "statutory ~87.5%"),
    ("Centrally Assessed / Commercial", "Include Centrally Assessed in Taxed Value", "CENTRALLY_ASSESSED", "centrally_assessed", "yesno", "Yes/No"),''',
   '''    ("Centrally Assessed / Commercial", "Centrally Assessed Property (market value)", "CENTRALLY_ASSESSED_VALUE", "centrally_assessed_value", "float", "$ market — § 59-2-201"),
    ("Centrally Assessed / Commercial", "Centrally Assessed Equipment (market value)", "CENTRALLY_ASSESSED_EQUIPMENT", "centrally_assessed_equipment", "float", "$ market"),
    ("Centrally Assessed / Commercial", "Centrally Assessed Taxable Ratio", "CENTRALLY_ASSESSED_RATIO", "centrally_assessed_ratio", "pct", "market value × ratio"),
    ("Centrally Assessed / Commercial", "Include Centrally Assessed in Taxed Value", "CENTRALLY_ASSESSED", "centrally_assessed", "yesno", "Yes/No"),''')

_p("inputs.py", '''    ("Fees", "County Collection Cost", "COUNTY_COLLECTION_FEE", "county_collection_fee", "pct", "% of mill revenue"),''',
   '''    ("Fees", "County Collection Cost", "COUNTY_COLLECTION_FEE", "county_collection_fee", "pct", "% of mill revenue; 0% in Utah (§ 59-2-1602)"),''')

_p("inputs.py", '''    ("Tax & Valuation", "Personal Property Uniform Fee %", "UNIFORM_FEE_PRC", "uniform_fee_prc", "pct", "% of mill revenue"),
    ("Tax & Valuation", "Uniform Fee Taxable Value Threshold", "UNIFORM_FEE_AV_THRESHOLD", "uniform_fee_av_threshold", "float", "$"),''',
   '''    ("Tax & Valuation", "Personal Property Uniform Fee %", "UNIFORM_FEE_PRC", "uniform_fee_prc", "pct", "§ 59-2-405; % of mill revenue"),
    ("Tax & Valuation", "Uniform Fee Taxable Value Threshold", "UNIFORM_FEE_AV_THRESHOLD", "uniform_fee_av_threshold", "float", "$"),''')

_p("inputs.py", '''    ("Tax & Valuation", "Primary Residential Taxable Ratio", "RESID_TAXABLE_RATIO", "resid_taxable_ratio", "pct", ""),
    ("Tax & Valuation", "Prior Residential Taxable Ratio", "RESID_TAXABLE_RATIO_PRIOR", "resid_taxable_ratio_prior", "pct", "residential taxable ratio when the governing document was adopted"),''',
   '''    ("Tax & Valuation", "Primary Residential Taxable Ratio", "RESID_TAXABLE_RATIO", "resid_taxable_ratio", "pct", "55% — the 45% exemption, § 59-2-103"),
    ("Tax & Valuation", "Prior Residential Taxable Ratio", "RESID_TAXABLE_RATIO_PRIOR", "resid_taxable_ratio_prior", "pct", "ratio before the current exemption"),''')

_p("inputs.py", '''    ("O&M", "O&M Carveout", "OM_EXPENSE", "om_expense", "float", "$ carved out of pledged revenue each year for operations"),
]''',
   '''    ("District Costs", "Starting O&M Expense", "OM_EXPENSE", "om_expense", "float", "$ per year of district operations & administration, netted from the revenue available to both liens"),
    ("District Costs", "O&M Expense Growth Rate", "OM_GROWTH_RATE", "om_growth_rate", "pct", "annual inflation on the O&M base"),
    ("District Costs", "O&M Expense Taxable Value Limit", "OM_EXPENSE_AV_LIMIT", "om_expense_av_limit", "float", "$ — above this taxable value the charge stops; 0 ⇒ no limit"),
    ("District Costs", "First Year District Costs Are Charged", "DISTRICT_COST_START_YEAR", "district_cost_start_year", "int", "blank ⇒ two years after closing"),
]''')

_p("inputs.py", '''    ("Tax & Valuation", "Reassessment - Residential", "REASSESS_RATE", "reassess_rate", "pct", ""),
    ("Tax & Valuation", "Reassessment - Commercial", "REASSESS_COMM_RATE", "reassess_comm_rate", "pct", ""),''',
   '''    ("Tax & Valuation", "Reassessment - Residential", "REASSESS_RATE", "reassess_rate", "pct", "applied annually in Utah"),
    ("Tax & Valuation", "Reassessment - Commercial", "REASSESS_COMM_RATE", "reassess_comm_rate", "pct", ""),''')

_p("inputs.py", '''    ("Lot / Home Valuation", "Developed Lot Value (% of market)", "LOT_INVENTORY_TAXABLE_RATIO", "lot_inventory_taxable_ratio", "pct", ""),''',
   '''    ("Lot / Home Ratios", "Builder Lot Inventory Taxable Ratio", "LOT_INVENTORY_TAXABLE_RATIO", "lot_inventory_taxable_ratio", "pct", "55% under Utah Admin. Code R884-24P-52; 100% to tax at full market"),''')

_p("inputs.py", '''_RATE_SOURCES = [
    "Colorado DPT — Historical Taxable Ratios (dpt.colorado.gov/historical-assessment-rates)",
    "Colorado DPT — Residential Taxable Ratio Study (dpt.colorado.gov/residential-assessment-rate-study)",
    "Colorado ARL Ch. 6 — Property Classification & Assessment Percentages (builder lot inventory / nonresidential)",
    "SB19-255; SB21-293; SB22-238; SB23B-001; SB24-233; HB24B-1001",
]''',
   '''_RATE_SOURCES = [
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
]''')

_p("inputs.py", '''_RESIDENTIAL_TABLE_TITLE = "Historical Residential Exemption / Colorado Residential Taxable Ratios"''',
   '''_RESIDENTIAL_TABLE_TITLE = "Utah Primary Residential Exemption — Taxable Share of Fair Market Value"''')

_p("inputs.py", '''        ("EDITABLE — residential (Utah) taxable ratio by tax/roll year. A metro "
         "district is a local government; use the local-government rate. The model "
         "READS these yellow rate cells and applies them to residential value by roll "
         "year (carry-forward for years past the last row). Edit a rate to retune it "
         "(e.g. 2026+ = 6.80% per HB24B-1001)."),''',
   '''        ("EDITABLE — Utah taxes primary residential property on 55% of fair market value "
         "— the 45% exemption of Utah Const. art. XIII, § 3 and Utah Code § 59-2-103, "
         "covering the dwelling and up to one acre of land. The rate has been flat at "
         "55% since 1995. The model READS these yellow cells and applies them to "
         "residential value by roll year (carry-forward past the last row); the "
         "RESID_TAXABLE_RATIO cell on the Inputs tab overrides them when changed."),''')


_p("inputs.py", '''    r += 2
    sc = ws.cell(row=r, column=2, value="Sources")''',
   '''    # ── Utah property tax calendar ───────────────────────────────────────────
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
    sc = ws.cell(row=r, column=2, value="Sources")''')

_p("inputs.py", '''        ("EDITABLE — builder lot inventory is a NONRESIDENTIAL subclass (fixed at 29% under "
         "Gallagher; phasing to 25% under SB24-233). The model READS these yellow "
         "rate cells and applies them to lot-inventory value by roll year (carry-forward "
         "for years past the last row). Edit a rate to retune the phase-down."),
        "Vacant-Land Taxable Ratio", _INVENTORY_HISTORY, editable=True)''',
   '''        ("EDITABLE — Utah taxes non-exempt property at 100% of fair market value, but "
         "Utah Admin. Code R884-24P-52 lets the primary residential exemption reach "
         "unoccupied property and property under construction that the assessor determines "
         "will be a primary residence once occupied — so builder lot inventory is carried "
         "at the same 55%. The model READS these yellow rate cells and applies them to "
         "lot-inventory value by roll year (carry-forward past the last row). Enter 100% "
         "to tax inventory at full market value instead."),
        "Lot Inventory Taxable Ratio", _INVENTORY_HISTORY, editable=True)''')


# ── memo.py ──────────────────────────────────────────────────────────────────

_p("memo.py", '''    gall = cfg.gallagherization.strip().lower() == "yes"''', "")

_p("memo.py", '''    _mill_bullet = (
        f"A total mill levy of {mill_total:.3f} mills is assumed &mdash; {mill_ds:.3f} mills to "
        f"debt service (the governing document cap) and {mill_ops:.3f} mills to operations &amp; "
        f"maintenance."
        + (f" The debt-service mill is &ldquo;adjusted&rdquo; to ~{eff_mill:.3f} effective "
           f"mills to offset the decline in the residential taxable ratio since the Service "
           f"Plan was adopted ({cfg.resid_taxable_ratio_prior:.2%} &rarr; {cfg.resid_taxable_ratio:.2%})."
           if gall else ""))''',
   '''    from .config import PID_STATUTORY_LEVY_CAP
    statutory_mills = PID_STATUTORY_LEVY_CAP * 1000.0
    _mill_bullet = (
        f"A district levy of {mill_total:.3f} mills is assumed &mdash; {mill_ds:.3f} mills for "
        f"debt service"
        + (f" and {mill_ops:.3f} mills for operations" if mill_ops else "")
        + f". The levy is capped at the most restrictive of {statutory_mills:.3f} mills "
          f"(&sect;&nbsp;17D-4-303), the governing document, and the indentures &mdash; "
          f"{cfg.mill_levy_cap:.3f} mills here, so {eff_mill:.3f} mills is levied. Utah caps are "
          f"fixed rates per dollar of taxable value and do not float with the residential "
          f"exemption, so there is no Colorado-style &ldquo;Gallagherization&rdquo; adjustment. "
          f"Because the required mill levy stays within those caps, the District need not give "
          f"Truth in Taxation notice or hold a hearing to impose it.")''')

_p("memo.py", '''         f"{_vl_txt} on builder lot inventory / nonresidential (SB24-233 phase-down, by roll year)."),''',
   '''         f"{_vl_txt} on builder lot inventory, which carries the same exemption where the "
         f"assessor determines the property will be a primary residence once occupied "
         f"(Utah Admin. Code R884-24P-52)."),''')

_p("memo.py", '''        (f"Colorado taxable ratios: {cfg.resid_taxable_ratio:.2%} residential (Utah) and "''',
   '''        (f"Utah taxable ratios: {cfg.resid_taxable_ratio:.2%} on primary residential value "
         f"(the 45% exemption, &sect;&nbsp;59-2-103) and "''')

_p("memo.py", '''        (f"A {cfg.reassess_rate:.1%} reassessment is applied in odd-numbered "
         f"(re-valuation) years; value is set at the June&nbsp;30 level of value the year before "
         f"reappraisal (~{cfg.av_lag_years}-year lag to collection)."),''',
   '''        (f"County assessors revalue annually (&sect;&nbsp;59-2-303.1); value already on the roll "
         f"grows at {cfg.reassess_rate:.1%} a year. Value created in a calendar year lands on the "
         f"following 1&nbsp;January roll, is billed that November, and pays debt service the next "
         f"1&nbsp;March &mdash; a {cfg.av_lag_years}-year lag from creation to the payment it "
         f"supports."),''')

_p("memo.py", '''         f"sales price &asymp; {_money(wasp)}); value phases onto the assessed-value roll as homes close."),''',
   '''         f"sales price &asymp; {_money(wasp)}); value phases onto the tax roll as homes close."),''')

_p("memo.py", '''    if cfg.capi_term > 0:
        assumptions.append(''',
   '''    assumptions.append(
        f"Utah property taxes are levied on value as of 1&nbsp;January and are due in a single "
        f"payment on 30&nbsp;November, so district principal is structured on 1&nbsp;March with "
        f"interest on 1&nbsp;March and 1&nbsp;September &mdash; each year&rsquo;s collections are "
        f"in hand before the payment they support. Delinquent taxes carry a penalty of the greater "
        f"of 2.5% or $10, and the Public Infrastructure District Act permits the District to impose "
        f"an additional 7% annual penalty; no penalty revenue is credited here.")
    if cfg.uniform_fee_prc:
        assumptions.append(
            f"Registered personal property pays a uniform fee in lieu of ad valorem tax "
            f"(&sect;&nbsp;59-2-405), distributed to taxing entities in the same proportion as real "
            f"property tax; {cfg.uniform_fee_prc:.2%} of the levy is credited.")
    else:
        assumptions.append(
            "No credit is taken for the personal property uniform fee "
            "(&sect;&nbsp;59-2-405), which would otherwise be distributed to the District in the "
            "same proportion as real property tax.")
    assumptions.append(
        f"County assessing and collecting is funded by a separate statewide levy on property "
        f"(&sect;&nbsp;59-2-1602) rather than a deduction from the District&rsquo;s distribution, "
        f"so {cfg.county_collection_fee:.2%} is netted from pledged revenue "
        f"(a Colorado metropolitan district would lose roughly 1.50% to the county treasurer).")
    if cfg.capi_term > 0:
        assumptions.append(''')

_p("memo.py", '''Colorado metropolitan-district financing and a subsequent senior refunding, held to the
Service-Plan mill-levy limit. The following is a summary of the assumptions used in the
analysis:</p>''',
   '''Utah public infrastructure district financing and a subsequent senior refunding, held within the
levy caps that bind a public infrastructure district under the Public Infrastructure District
Act, Title&nbsp;17D, Chapter&nbsp;4, Utah Code. The following is a summary of the assumptions used
in the analysis:</p>''')

_p("memo.py", '''Peak AV &asymp; {_money(peak_av)}''', '''Peak Taxable Value &asymp; {_money(peak_av)}''')

# Utah's fixed levy caps are tight enough that a refunding can produce NEGATIVE
# new money (see docs/utah-vs-colorado.md).  Upstream renders that as "$-914,919";
# show it in accounting parentheses so the client memo reads correctly.
_p("memo.py", '''def _money(x):
    return f"${x:,.0f}"''',
   '''def _money(x):
    # Accounting style: negatives in parentheses.  A refunding under Utah's
    # fixed levy caps can genuinely return less than it costs, so this is a
    # real case for a PID rather than a defensive flourish.
    return f"(${abs(x):,.0f})" if x < 0 else f"${x:,.0f}"''')


# Two prose strings wrap across source lines, so the label pass cannot see them
# whole; patch them explicitly.
_p("report.py", '''                "Aggregate per year: market value (beginning & ending) → total assessed "
                "value → gross & net revenue for sizing."''',
   '''                "Aggregate per year: market value (beginning & ending) → total taxable "
                "value → gross & net revenue for sizing."''')

_p("report.py", '''    O&M revenue projection: the operations mill levy applied to total assessed
    value, collected at the collection rate, plus the specific-ownership tax —''',
   '''    O&M revenue projection: the operations mill levy applied to total taxable
    value, collected at the collection rate, plus the personal property uniform fee —''')



_p("report.py", '''                "certified-value adjustments → cumulative 100% lot value → assessed."], 10)''',
   '''                "certified-value adjustments → cumulative 100% lot value → taxable "
                "value @ the residential exemption."], 10)''')

_p("report.py", '''                "Beginning market value → + new home value added to rolls → + biennial "
                "reassessment → + certified-value adjustments → gross market value → assessed."], 10)''',
   '''                "Beginning market value → + new home value added to rolls → + annual "
                "reassessment → + certified-value adjustments → gross market value → "
                "taxable value @ the residential exemption."], 10)''')

_p("report.py", '''        ("note", "Biennial reassessment (residential / commercial)",''',
   '''        ("note", f"{cfg.reassess_frequency} reassessment (residential / commercial)",''')

_p("inputs.py", '''"Strict Biennial Level-of-Value", "HOLD_VALUE_FLAT"''',
   '''"Hold Value Flat Between Reassessments", "HOLD_VALUE_FLAT"''')



# ── District operations & maintenance ────────────────────────────────────────
# Colorado's single "O&M Carveout" is the district's standing annual cost, and
# it stays a single line here: administration (accounting, audit, legal,
# assessor, continuing disclosure) and operations (landscaping, parks, snow
# removal, lighting) are the same budget for a Utah PID, which rarely carries a
# separate operations levy to fund either.  Three Utah refinements:
#
#  1. The base ACTUALLY INFLATES.  Colorado's summary.py charges a flat
#     `om_carveout` and never reads `om_growth_rate`; here the growth rate
#     applies, which is what ties the model to the reference workbook.
#  2. Nothing is charged before `district_cost_start_year`.
#  3. It is netted from the revenue available to BOTH liens.  The subordinate
#     lien's own revenue is measured as `net_sub_revenue - net_senior_revenue`,
#     so a cost netted from the senior side alone is handed straight to the sub
#     — which would make an operating expense *raise* subordinate capacity.
#     Money the district actually spends is available to neither bond.

_p("summary.py", '''    net_senior_revenue: float      # AX
    net_sub_revenue: float         # BO''',
   '''    net_senior_revenue: float      # AX
    net_sub_revenue: float         # BO
    om_expense: float = 0.0        # district O&M, netted from both liens''')

_p("summary.py", '''                net_sub_revenue=net_sub_revenue,
            )''',
   '''                net_sub_revenue=net_sub_revenue,
                om_expense=om_expense,
            )''')

_p("report.py", '''             ("om", "− District\\nO&M", 11),
             ("net", "Net Revenue\\n(senior sizing)", 15)]''',
   '''             ("om", "− District\\nO&M", 11),
             ("net", "Net Revenue\\n(senior sizing)", 15)]''')



# The O&M tab showed only the operations-levy revenue.  Now that the district
# carries a modelled O&M expense, show it alongside — and the surplus/(deficit),
# which is the number that says whether the operations levy actually covers the
# operating budget or whether the debt levy is carrying it.
_p("report.py", '''    ops_mill = cfg.mill_levy_ops_target
    coll = cfg.tax_collect_mill_prc
    _title(ws, cfg, [cfg.pid_name, "Operations & Maintenance (O&M) Revenue Projection",
                f"Operations mill levy {ops_mill:.3f} mills @ {coll:.1%} collection"], 6)
    hdrs = [(1, "Collection\\nYear", 12), (2, "Total\\nTaxable Value", 16),
            (3, "Operations\\nMill Levy", 13),
            (4, f"Total Collections\\n@ {coll:.1%}", 16),
            (5, f"Uniform Fee\\n@ {cfg.uniform_fee_prc:.0%}", 14),
            (6, "Total Available\\nfor O&M", 16)]''',
   '''    ops_mill = cfg.mill_levy_ops_target
    coll = cfg.tax_collect_mill_prc
    _title(ws, cfg, [cfg.pid_name, "Operations & Maintenance (O&M) Revenue and Expense",
                f"Operations mill levy {ops_mill:.3f} mills @ {coll:.1%} collection"
                + (f"  ·  O&M expense ${cfg.om_expense:,.0f} base, inflating at "
                   f"{cfg.om_growth_rate:.1%}" if cfg.om_expense else
                   "  ·  no O&M expense entered")], 8)
    hdrs = [(1, "Collection\\nYear", 12), (2, "Total\\nTaxable Value", 16),
            (3, "Operations\\nMill Levy", 13),
            (4, f"Total Collections\\n@ {coll:.1%}", 16),
            (5, f"Uniform Fee\\n@ {cfg.uniform_fee_prc:.0%}", 14),
            (6, "Total Available\\nfor O&M", 16),
            (7, "− O&M\\nExpense", 14),
            (8, "O&M Surplus /\\n(Deficit)", 16)]''')

_p("report.py", '''    tot_coll = tot_uniform_fee = tot_avail = 0.0
    for i, r in enumerate(sm.rows):''',
   '''    tot_coll = tot_uniform_fee = tot_avail = 0.0
    tot_om = tot_net_om = 0.0
    for i, r in enumerate(sm.rows):''')

_p("report.py", '''        tot_coll += collections; tot_uniform_fee += uniform_fee; tot_avail += avail''',
   '''        om_expense = r.om_expense
        net_om = avail - om_expense
        tot_coll += collections; tot_uniform_fee += uniform_fee; tot_avail += avail
        tot_om += om_expense; tot_net_om += net_om''')

_p("report.py", '''        _cell(ws, rw, 6, round(avail, 0) or None, fill, fmt=_DOLLAR)

    rw = 6 + len(sm.rows)''',
   '''        _cell(ws, rw, 6, round(avail, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 7, -round(om_expense, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 8, round(net_om, 0) or None, fill, font=_BOLD, fmt=_DOLLAR)

    rw = 6 + len(sm.rows)''')

_p("report.py", '''    _cell(ws, rw, 6, round(tot_avail, 0) or None, _TOTAL, font=_TOTAL_FONT, fmt=_DOLLAR)
    ws.freeze_panes = "A6"''',
   '''    _cell(ws, rw, 6, round(tot_avail, 0) or None, _TOTAL, font=_TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, rw, 7, -round(tot_om, 0) or None, _TOTAL, font=_TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, rw, 8, round(tot_net_om, 0) or None, _TOTAL, font=_TOTAL_FONT, fmt=_DOLLAR)
    ws.freeze_panes = "A6"''')

_p("report.py", '''        ("note", "County collection cost", _pct(cfg.county_collection_fee)),''',
   '''        ("note", "County collection cost", _pct(cfg.county_collection_fee)),
        ("note", "District O&M expense (base / growth)",
            (f"${cfg.om_expense:,.0f} / {_pct(cfg.om_growth_rate)}"
             if cfg.om_expense else "none entered")),''')



_p("report.py", '''    O&M revenue projection: the operations mill levy applied to total taxable
    value, collected at the collection rate, plus the personal property uniform fee —
    the total available each year for operations & maintenance.
    """''',
   '''    O&M revenue and expense: the operations mill levy applied to total taxable
    value, collected at the collection rate, plus the personal property uniform
    fee — the total available each year for operations & maintenance — against
    the district's modelled O&M expense, and the surplus or deficit between
    them.  A Utah PID usually runs no operations levy, so the deficit shown here
    is what the debt-service levy is carrying.
    """''')



# ── Series C is carried inert ────────────────────────────────────────────────
# Colorado's optional third lien is sized against a SEPARATE assessment that
# reassesses the created value at its own rate.  Whether a Utah PID can levy
# against a second assessment on the same property is a statutory question this
# model has not worked through, so Series C stays off (SIZE_SERIES_C = "No") and
# its rows are kept OFF the Inputs template — an untested toggle in a
# client-facing template is worse than no toggle.  The code ports across
# unchanged, so turning it on is a one-line change once the authority is
# settled.  The developer-contribution rows are generic and stay.
_p("inputs.py", '''    ("Financing Toggles", "Size Series C Cash-Flow Bonds", "SIZE_SERIES_C", "size_series_c", "yesno", "Yes/No"),\n''', "")

_p("inputs.py", '''    ("Tax & Valuation", "Reassessment - Series C", "SERIES_C_REASSESS_RATE", "series_c_reassess_rate", "pct", "separate assessment used only to size the Series C cash-flow bonds"),
    ("Series C Bonds", "Series C Interest Rate", "SERIES_C_INTEREST_RATE", "series_c_interest_rate", "pct", "accreting cash-flow coupon"),
    ("Series C Bonds", "Series C Coverage", "DSC_SERIES_C", "dsc_series_c", "float", "debt-service coverage"),
    ("Series C Bonds", "Series C Final Maturity (years)", "SERIES_C_FINAL_MAT_YRS", "series_c_final_mat_yrs", "int", "years from delivery; same 12/15 payment dates as the sub lien"),
''', "")

_p("inputs.py", '''    ("Developer Contribution", "Applied To (Senior / Subordinate / Series C / Proportional)", "DEVELOPER_CONTRIBUTION_SERIES", "developer_contribution_series", "str", "which series the contribution funds; Proportional spreads it by par"),''',
   '''    ("Developer Contribution", "Applied To (Senior / Subordinate / Proportional)", "DEVELOPER_CONTRIBUTION_SERIES", "developer_contribution_series", "str", "which series the contribution funds; Proportional spreads it by par"),''')

_p("inputs.py", '''        type="list", formula1='"Senior,Subordinate,Series C,Proportional"',''',
   '''        type="list", formula1='"Senior,Subordinate,Proportional"',''')



# ── Dates read m/d/yyyy across the Excel output ──────────────────────────────
# The number formats are handled in the label pass; these are the dates rendered
# as text — the title band on every tab and the prose subtitles on the CAPI and
# Call Schedule sheets.
#
# File names are the exception: m/d/yyyy contains path separators, and
# deliverable_basename strips them, which would turn 9/24/2026 into "9242026".
# Saved files therefore carry m.d.yyyy, the same substitution Colorado settled on
# in fd4a687.
_p("report.py", '''def _today_str() -> str:
    from datetime import date as _dt
    _t = _dt.today()
    return f"{_t.strftime('%B')} {_t.day}, {_t.year}"''',
   '''def _today_str() -> str:
    from datetime import date as _dt
    return _d(_dt.today())


def _d(d) -> str:
    """A date as m/d/yyyy — the workbook's date format, unpadded."""
    return f"{d.month}/{d.day}/{d.year}"


def _d_file(d) -> str:
    """
    A date for a FILE NAME — m.d.yyyy.

    m/d/yyyy carries path separators, and deliverable_basename strips the
    characters a filesystem will not take, which would leave "9242026".
    """
    return f"{d.month}.{d.day}.{d.year}"


def _today_file_str() -> str:
    from datetime import date as _dt
    return _d_file(_dt.today())


def _ym(d) -> str:
    """A month as m/yyyy."""
    return f"{d.month}/{d.year}"''')

_p("report.py", '''    parts = [_today_str(), "Reimbursement Analysis",''',
   '''    parts = [_today_file_str(), "Reimbursement Analysis",''')

_p("report.py", '''                f"{cfg.capi_end_date:%b %Y}; balance earns {cfg.interest_earn_rate:.2%}/yr"], 6)''',
   '''                f"{_ym(cfg.capi_end_date)}; balance earns {cfg.interest_earn_rate:.2%}/yr"], 6)''')

_p("report.py", '''        f"Delivered {cfg.delivery:%B %d, %Y}; callable on and after "
        f"{cfg.premium_call_date:%B %d, %Y} at par plus accrued interest and a premium of:",''',
   '''        f"Delivered {_d(cfg.delivery)}; callable on and after "
        f"{_d(cfg.premium_call_date)} at par plus accrued interest and a premium of:",''')

_p("report.py", '''            f"Delivered {rd:%B %d, %Y}; callable on and after {ref_par:%B %d, %Y} "''',
   '''            f"Delivered {_d(rd)}; callable on and after {_d(ref_par)} "''')



# ── Builder lot inventory is driven by lot DELIVERIES, not home closings ─────
# Upstream sets the "Value of New Lots" column to l2h(y) — the value of lots
# CONVERTING TO HOMES in year y — so the whole tab keys off home closings.  A
# builder can hold 100 delivered lots for two years and the tab reports $0 of
# inventory, which is wrong on any reading: the lots exist on the 1 January
# lien date and the county assesses them (§ 59-2-103, § 59-2-303.1).
#
# It looks right on the Viridian No. 1 reference deal only by coincidence —
# closings there equal the prior year's deliveries exactly, so "converted in y"
# and "delivered in y-1" are the same series.  Any deal where homes lag lots
# breaks it.
#
# The column headers already describe the correct construction: new lots ADD
# value, lots rolled into homes REMOVE it, the net accumulates.  Feeding the
# first column the prior year's DELIVERIES makes the arithmetic match the
# headers, and the cumulative then telescopes exactly to
# ``vacant_lot_market_value(y - 1)`` — lots delivered but not yet closed, lagged
# one year onto the roll.  Worth pushing back to co_metro_model.
_p("development.py", '''        for y in range(first, (max(trued) + 1) if trued else first):
            net = l2h(y) - l2h(y - 1)''',
   '''        for y in range(first, (max(trued) + 1) if trued else first):
            net = self.lot_market_value.get(y - 1, 0.0) - l2h(y - 1)''')

_p("development.py", '''        net_years = [y for y in range(first, horizon_end + 1)
                     if abs(l2h(y) - l2h(y - 1)) > 1e-6]''',
   '''        net_years = [y for y in range(first, horizon_end + 1)
                     if abs(self.lot_market_value.get(y - 1, 0.0) - l2h(y - 1)) > 1e-6]''')

_p("development.py", '''        for y in range(first, horizon_end + 1):
            new_lots = l2h(y)
            lots_to_homes = -l2h(y - 1)''',
   '''        for y in range(first, horizon_end + 1):
            # Lots delivered in y-1 land on the roll set 1 January of year y.
            new_lots = self.lot_market_value.get(y - 1, 0.0)
            lots_to_homes = -l2h(y - 1)''')

_p("development.py", '''        Columns: Value of New Lots (lot value rolling into homes), − Lots to Homes
        (prior year, the lag), Net Value with Lag, Adjustments (recognition that''',
   '''        Columns: Value of New Lots (delivered the prior year — a lot platted
        during year y-1 is on the roll set 1 January of year y), − Lots to Homes
        (the value leaving inventory as those lots are built out), Net Value with
        Lag (the change in builder inventory, so the running Cumulative equals
        ``vacant_lot_market_value(y - 1)``), Adjustments (recognition that''')



# ── Debt Structure: refuse to read a pre-16b4edd sheet ───────────────────────
# The tab gained a "Par Amount" column, so every column after B shifted right.
# An Inputs workbook saved against the OLD layout puts its coupons where par is
# now read (0.06 becomes a $0.06 par) and its term-final-maturity year where the
# yield is now read (2054 becomes a 205,400% yield).  Both parse as numbers, so
# nothing would complain — the model would just size a nonsense structure.
# Check the magnitudes and say plainly what happened.
_p("inputs.py", '''        if not rows:''',
   '''        for _y, _par, _cpn, _yld, _typ in rows:
            if _par is not None and 0 < _par < 1:
                raise ValueError(
                    f"Debt Structure: {_y} has a Par Amount of {_par:g}, which is a "
                    f"rate, not a dollar amount.  This sheet looks like the older "
                    f"layout (Maturity | Coupon | Yield | Term Final Maturity).  "
                    f"The tab now reads Maturity | Par Amount | Coupon | Yield | "
                    f"Type, so every entry after the maturity year sits one column "
                    f"left of where the model reads it.  Regenerate the sheet with "
                    f"write_inputs_workbook() and re-enter the scale.")
            for _label, _v in (("Coupon", _cpn), ("Yield", _yld)):
                if _v is not None and _v > 1:
                    raise ValueError(
                        f"Debt Structure: {_y} has a {_label} of {_v:g}.  Coupons and "
                        f"yields are decimals (0.06 = 6.00%); a year here means the "
                        f"sheet is in the older column layout.  Regenerate it with "
                        f"write_inputs_workbook() and re-enter the scale.")
        if not rows:''')



# ── Pricing uses the entered coupon, not the flat sizing rate ────────────────
# `price_for` passed `self.rate` — the Inputs-page flat rate the structure is
# SIZED with — as the coupon, so the Coupon column on the Debt Structure tab
# never reached the price.  A 2056 term at 6.250% / 6.625% off a 5.875% flat
# rate priced at 90.334 instead of 95.167: a 4.8-point error on the OID, which
# is a Source of Funds and therefore moves the reimbursement.
#
# Also make `coupon_for` respect term-bond membership.  The coupon is entered on
# the term's FINAL maturity row, and `_schedule_lookup` carries forward, so an
# installment inside a LATER term would otherwise pick up the earlier term's
# coupon — wrong for pricing and for the interest `_apply_coupon_scale` derives.
# Worth pushing back to co_metro_model.
_p("debt_service.py", '''    def coupon_for(self, year: int) -> float:
        """Coupon for a maturity — the per-maturity scale (carry-forward) or the flat rate."""
        if self.coupon_scale:
            from .config import _schedule_lookup
            return _schedule_lookup(self.coupon_scale, year, self.rate)
        return self.rate''',
   '''    def coupon_for(self, year: int) -> float:
        """
        Coupon for a maturity — the per-maturity scale (carry-forward) or the
        flat rate.

        A term bond carries one coupon, entered on its FINAL maturity row, so an
        installment inside a term looks that row up rather than carrying forward
        from whatever precedes it.
        """
        if self.coupon_scale:
            from .config import _schedule_lookup
            term = self._term_for(year)
            lookup_year = term[1] if term is not None else year
            return _schedule_lookup(self.coupon_scale, lookup_year, self.rate)
        return self.rate''')

_p("debt_service.py", '''            maturity = date(last, self.prin_month, self.prin_day)
            return price_to_worst(self.delivery, maturity, self.rate, ty,
                                  self._call_scenarios())
        maturity = date(year, self.prin_month, self.prin_day)
        return price_to_worst(self.delivery, maturity, self.rate,
                              self.yield_for(year), self._call_scenarios())''',
   '''            maturity = date(last, self.prin_month, self.prin_day)
            return price_to_worst(self.delivery, maturity, self.coupon_for(last), ty,
                                  self._call_scenarios())
        maturity = date(year, self.prin_month, self.prin_day)
        return price_to_worst(self.delivery, maturity, self.coupon_for(year),
                              self.yield_for(year), self._call_scenarios())''')



# ── The revenue wrap sizes at the entered coupons ────────────────────────────
# `size_for_par` charged interest as `balance * rate` — the flat Inputs-page rate
# the deal is SIZED with — while `_apply_coupon_scale` afterwards restated the
# actual interest at the per-maturity coupons.  Enter a 6.250% scale against a
# 5.875% Inputs rate and the deal is sized as though it pays 5.875% but really
# pays 6.250%, so achieved coverage lands under target (1.29x against a 1.30x
# target on the reference deal, minimum 0.76x against 0.82x).
#
# Charge the wrap the same interest the bonds will actually pay: in year y,
# sum(P_m * coupon_m) over the maturities still outstanding.  That is circular —
# the interest depends on principal not yet allocated — so solve it as a fixed
# point, seeded with the flat-rate schedule.  An error of e in the future
# principal moves interest by coupon * e, so the map contracts by roughly the
# coupon (~0.06) each pass and settles in a handful of iterations; the $5,000
# rounding means it either lands exactly or wobbles one step, so the loop is
# capped.  Worth pushing back to co_metro_model.
_p("debt_service.py", '''    def coupon_for(self, year: int) -> float:
        """
        Coupon for a maturity — the per-maturity scale (carry-forward) or the
        flat rate.

        A term bond carries one coupon, entered on its FINAL maturity row, so an
        installment inside a term looks that row up rather than carrying forward
        from whatever precedes it.
        """
        if self.coupon_scale:
            from .config import _schedule_lookup
            term = self._term_for(year)
            lookup_year = term[1] if term is not None else year
            return _schedule_lookup(self.coupon_scale, lookup_year, self.rate)
        return self.rate''',
   '''    def coupon_for(self, year: int) -> float:
        """
        Coupon for a maturity — the per-maturity scale (carry-forward) or the
        flat rate.

        A term bond carries one coupon, entered on its FINAL maturity row, so an
        installment inside a term looks that row up rather than carrying forward
        from whatever precedes it.
        """
        return coupon_at(self.coupon_scale, self.term_bonds, self.rate, year)''')

_p("debt_service.py", '''@dataclass
class BondTranche:''',
   '''def coupon_at(coupon_scale, term_bonds, rate: float, year: int) -> float:
    """
    Coupon for a maturity, independent of any tranche — the sizer needs this
    before a tranche exists, and it must agree with ``BondTranche.coupon_for``.

    A term bond carries a single coupon, entered on its FINAL maturity row, so a
    year inside a term looks that row up instead of carrying forward from
    whatever precedes it.
    """
    if not coupon_scale:
        return rate
    from .config import _schedule_lookup
    lookup_year = year
    for first, last, _ty in (term_bonds or []):
        if first <= year <= last:
            lookup_year = last
            break
    return _schedule_lookup(coupon_scale, lookup_year, rate)


@dataclass
class BondTranche:''')

_p("debt_service.py", '''        def size_for_par(par: float) -> dict[int, float]:
            """Apply the wrap formula for an assumed par; returns the principal schedule."""
            principals: dict[int, float] = {}
            balance = par
            for y in principal_years:
                if capi_end_year is not None and y <= capi_end_year:
                    # Interest capitalized — no principal sized during the CAPI period.
                    principals[y] = 0.0
                    continue
                net_rev = self.sm.net_senior_revenue(y)
                # DSRF interest earnings offset debt service dollar-for-dollar
                # (net_total already subtracts them), so credit them to the DS
                # target directly rather than grossing them up by coverage.  This
                # keeps the resulting net-DS coverage exactly at the target.
                target = net_rev / coverage + dsrf_earn
                avail = target - balance * rate
                if release_surplus and y == final_year:
                    avail += dsrf_deposit  # released DSRF pays down the final maturity
                p = max(0.0, math.floor(avail / 5000.0) * 5000.0)
                principals[y] = p
                balance -= p
            return principals''',
   '''        def _wrap(par: float, interest_of) -> dict[int, float]:
            """The wrap formula for an assumed par, charging ``interest_of(year,
            balance)`` as that year's interest; returns the principal schedule."""
            principals: dict[int, float] = {}
            balance = par
            for y in principal_years:
                if capi_end_year is not None and y <= capi_end_year:
                    # Interest capitalized — no principal sized during the CAPI period.
                    principals[y] = 0.0
                    continue
                net_rev = self.sm.net_senior_revenue(y)
                # DSRF interest earnings offset debt service dollar-for-dollar
                # (net_total already subtracts them), so credit them to the DS
                # target directly rather than grossing them up by coverage.  This
                # keeps the resulting net-DS coverage exactly at the target.
                target = net_rev / coverage + dsrf_earn
                avail = target - interest_of(y, balance)
                if release_surplus and y == final_year:
                    avail += dsrf_deposit  # released DSRF pays down the final maturity
                p = max(0.0, math.floor(avail / 5000.0) * 5000.0)
                principals[y] = p
                balance -= p
            return principals

        def size_for_par(par: float) -> dict[int, float]:
            """
            Principal schedule for an assumed par.

            With no coupon scale every maturity pays the flat rate and the wrap
            is a single pass.  With a scale the interest in a year is
            ``sum(P_m * coupon_m)`` over the maturities still outstanding, which
            depends on principal the pass has not allocated yet — so iterate to a
            fixed point from the flat-rate schedule.
            """
            flat = _wrap(par, lambda y, balance: balance * rate)
            if not coupon_scale:
                return flat

            def scaled_interest(principals):
                """
                Charge the running balance at the weighted-average coupon of the
                maturities still outstanding.

                Weighting the BALANCE (rather than summing P_m * coupon_m
                directly) keeps the flat case exact: when every coupon equals the
                flat rate the average is that rate and this reduces to
                ``balance * rate``, so entering a scale equal to the Inputs rate
                does not move the par.  Summing the schedule directly would
                charge interest on the schedule's total rather than on the par
                being tested, which differ by the rounding residual while the
                outer bisection is still searching.
                """
                def of(y, balance):
                    num = den = 0.0
                    for m, P in principals.items():
                        if m >= y:
                            num += P * coupon_at(coupon_scale, term_bonds, rate, m)
                            den += P
                    return balance * (num / den if den else rate)
                return of

            principals, history = flat, []
            for _ in range(50):
                if principals in history:
                    # The $5,000 rounding can leave the iteration in a short
                    # cycle rather than at an exact fixed point.  Take the
                    # SMALLEST schedule in the cycle: under-sizing by a rounding
                    # step leaves coverage at or above target, over-sizing pushes
                    # it below — which is the defect this whole change fixes.
                    cycle = history[history.index(principals):]
                    return min(cycle, key=lambda s: sum(s.values()))
                history.append(principals)
                nxt = _wrap(par, scaled_interest(principals))
                if nxt == principals:
                    return principals
                principals = nxt
            return principals''')



# ── Period count is day-accurate; Excel PRICE() available alongside ─────────
# `semiannual_periods` counted MONTHS ONLY — the day of the dated date was
# discarded, so bonds dated 1 July, 5 July and 31 July all priced identically.
# Count 30/360 days instead.
#
# The discounting stays a present value from the DATED date with no accrued
# subtracted, which is right for a new issue: settlement IS the dated date, no
# accrued interest changes hands, and the proceeds are price x par.  It also
# keeps the module's invariant that a bond reoffered at its coupon prices at
# exactly 100 — PV(full coupons at y == c) is 100 for any period count, whole or
# fractional, so an off-cycle dated date cannot inject a phantom OID.
#
# `clean_price` is added beside it for quoting on a 30/360 clean basis (Excel
# PRICE, accrued subtracted), which is what a secondary-market quote means.  It
# agrees with Excel to 1e-14.  Worth pushing back to co_metro_model.
_p("pricing.py", '''def semiannual_periods(settlement: date, redemption: date, freq: int = 2) -> float:
    """Number of coupon periods between two coupon-cycle dates."""
    months = (redemption.year - settlement.year) * 12 + (redemption.month - settlement.month)
    return months / (12.0 / freq)''',
   '''def days_30_360(d1: date, d2: date) -> int:
    """Day count on the 30/360 (US) basis — Excel's basis 0."""
    day1 = min(d1.day, 30)
    day2 = d2.day
    if day1 == 30 and day2 == 31:
        day2 = 30
    return (d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + (day2 - day1)


def shift_months(d: date, months: int) -> date:
    """Move a date by whole months, clamping the day to the target month."""
    m = d.month - 1 + months
    y = d.year + m // 12
    m = m % 12 + 1
    leap = y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)
    last = [31, 29 if leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    return date(y, m, min(d.day, last))


def semiannual_periods(settlement: date, redemption: date, freq: int = 2) -> float:
    """
    Coupon periods between settlement and redemption, on a 30/360 basis.

    Counted in DAYS, not whole months: a bond dated 5 July and one dated 31 July
    are not the same bond, and the fractional first period is what makes the
    difference.
    """
    return days_30_360(settlement, redemption) / (360.0 / freq)


def clean_price(settlement: date, redemption_date: date, coupon: float, ytm: float,
                redemption: float = 100.0, freq: int = 2) -> float:
    """
    Clean (quoted) price per 100 of face — Excel ``PRICE()`` on a 30/360 basis.

    This is the secondary-market convention: the buyer pays this plus accrued
    interest.  A new issue settling on its dated date pays no accrued, so the
    sizing engine uses the present value from the dated date instead — see
    ``price_to_worst``.
    """
    step = 12 // freq
    dates, c_date = [], redemption_date
    while c_date > settlement:
        dates.append(c_date)
        c_date = shift_months(c_date, -step)
    if not dates:
        return redemption
    dates.reverse()
    n = len(dates)
    nxt = dates[0]
    prev = shift_months(nxt, -step)
    e = float(days_30_360(prev, nxt)) or (360.0 / freq)
    dsc = days_30_360(settlement, nxt) / e       # fraction of the period remaining
    accrued = 1.0 - dsc
    i = ytm / freq
    c = coupon / freq * 100.0
    if abs(i) < 1e-12:
        return redemption + c * n - c * accrued
    price = redemption / (1.0 + i) ** (n - 1 + dsc)
    price += sum(c / (1.0 + i) ** (k + dsc) for k in range(n))
    return price - c * accrued''')



# ── Debt Structure: an entered Price overrides the calculated one ────────────
# The model's price agrees with DBC on convention — a bond reoffered at its
# coupon prices at exactly 100 on an off-cycle dated date, which Excel's clean
# PRICE() does not — but the two differ by ~0.024 on a long discount term bond
# because of the odd first period (9/30/2026 dated, first coupon 3/1/2027).
# On $7.18MM that is ~$1,700 of OID.
#
# On pricing day the underwriter's price is the price.  Column G takes it, and
# it overrides the calculation outright.  Added at the END so B-F keep their
# meaning — the tab has already moved its columns once.
_p("config.py", '''    senior_par_schedule: Optional[dict] = None          # {maturity_year: par} — manual amortization override''',
   '''    senior_par_schedule: Optional[dict] = None          # {maturity_year: par} — manual amortization override
    senior_price_scale: Optional[dict] = None          # {maturity_year: price per 100} — entered, overrides the calc''')

_p("config.py", '''    senior_refunding_par_schedule: Optional[dict] = None''',
   '''    senior_refunding_par_schedule: Optional[dict] = None
    senior_refunding_price_scale: Optional[dict] = None''')

_p("debt_service.py", '''    term_bonds: Optional[list] = None            # [(first_year, last_year, term_yield), ...]''',
   '''    term_bonds: Optional[list] = None            # [(first_year, last_year, term_yield), ...]
    price_scale: Optional[dict] = None           # {maturity_year: price} — entered on pricing day''')

_p("debt_service.py", '''        from .pricing import price_to_worst
        term = self._term_for(year)
        if term is not None:
            first, last, ty = term''',
   '''        from .pricing import price_to_worst
        term = self._term_for(year)
        # An entered price is the price — on pricing day the underwriter's
        # number governs, not a convention.  A term bond carries the price on
        # its FINAL maturity row, as it carries the coupon and the yield.
        if self.price_scale:
            keyed = term[1] if term is not None else year
            if keyed in self.price_scale:
                return float(self.price_scale[keyed])
            if year in self.price_scale:
                return float(self.price_scale[year])
        if term is not None:
            first, last, ty = term''')

# inputs.py — write and read the Price column.
_p("inputs.py", '''    for col in "BCDEF":
        ws.column_dimensions[col].width = 18
    ws.merge_cells("B1:F1")''',
   '''    for col in "BCDEFG":
        ws.column_dimensions[col].width = 18
    ws.merge_cells("B1:G1")''')

_p("inputs.py", '''    ws.merge_cells("B2:F2")''', '''    ws.merge_cells("B2:G2")''')

_p("inputs.py", '''        "structure one amortizing term bond priced to that maturity. Prices run to worst call.")''',
   '''        "structure one amortizing term bond priced to that maturity. Prices run to worst call. "
        "Price is optional: leave it blank and the model computes it from the coupon and yield; "
        "enter the underwriter's price (e.g. 95.148) and that price governs the OID outright.")''')

_p("inputs.py", '''        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)''',
   '''        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)''')

_p("inputs.py", '''        for i, h in enumerate(["Maturity Year", "Par Amount", "Coupon", "Yield", "Type (Serial/Term)"]):''',
   '''        for i, h in enumerate(["Maturity Year", "Par Amount", "Coupon", "Yield",
                               "Type (Serial/Term)", "Price (optional)"]):''')

_p("inputs.py", '''    def _block(r, marker, years, coupon_scale, yield_scale, term_bonds, par_schedule):''',
   '''    def _block(r, marker, years, coupon_scale, yield_scale, term_bonds, par_schedule,
               price_scale=None):''')

_p("inputs.py", '''            vals = [y, par, cpn, yld, typ]
            for i in range(5):
                cell = ws.cell(row=r, column=2 + i)
                if vals[i] is not None:
                    cell.value = vals[i]
                cell.fill = _INPUT; cell.border = _BORDER; cell.alignment = _C
                if i == 1:
                    cell.number_format = _DOLLAR
                if i in (2, 3):
                    cell.number_format = _PCT''',
   '''            vals = [y, par, cpn, yld, typ, (price_scale or {}).get(y)]
            for i in range(6):
                cell = ws.cell(row=r, column=2 + i)
                if vals[i] is not None:
                    cell.value = vals[i]
                cell.fill = _INPUT; cell.border = _BORDER; cell.alignment = _C
                if i == 1:
                    cell.number_format = _DOLLAR
                if i in (2, 3):
                    cell.number_format = _PCT
                if i == 5:
                    cell.number_format = "0.000"''')

_p("inputs.py", '''    r = _block(4, "SENIOR BONDS", senior_years, cfg.senior_coupon_scale,
               cfg.senior_yield_scale, cfg.senior_term_bonds, cfg.senior_par_schedule)''',
   '''    r = _block(4, "SENIOR BONDS", senior_years, cfg.senior_coupon_scale,
               cfg.senior_yield_scale, cfg.senior_term_bonds, cfg.senior_par_schedule,
               cfg.senior_price_scale)''')

_p("inputs.py", '''    _block(r, "REFUNDING BONDS", ref_years, cfg.senior_refunding_coupon_scale,
           cfg.senior_refunding_yield_scale, cfg.senior_refunding_term_bonds,
           cfg.senior_refunding_par_schedule)''',
   '''    _block(r, "REFUNDING BONDS", ref_years, cfg.senior_refunding_coupon_scale,
           cfg.senior_refunding_yield_scale, cfg.senior_refunding_term_bonds,
           cfg.senior_refunding_par_schedule, cfg.senior_refunding_price_scale)''')

_p("inputs.py", '''                rows.append((int(y),
                             _num(ds.cell(row=r, column=3).value),   # par
                             _num(ds.cell(row=r, column=4).value),   # coupon
                             _num(ds.cell(row=r, column=5).value),   # yield
                             ds.cell(row=r, column=6).value))        # type''',
   '''                rows.append((int(y),
                             _num(ds.cell(row=r, column=3).value),   # par
                             _num(ds.cell(row=r, column=4).value),   # coupon
                             _num(ds.cell(row=r, column=5).value),   # yield
                             ds.cell(row=r, column=6).value,         # type
                             _num(ds.cell(row=r, column=7).value)))  # price''')

_p("inputs.py", '''        for _y, _par, _cpn, _yld, _typ in rows:''',
   '''        for _y, _par, _cpn, _yld, _typ, _prc in rows:''')

_p("inputs.py", '''        coupon = {y: c for y, p, c, yl, t in rows if c is not None}
        par = {y: p for y, p, c, yl, t in rows if p is not None}
        yield_of = {y: yl for y, p, c, yl, t in rows}''',
   '''        coupon = {y: c for y, p, c, yl, t, pr in rows if c is not None}
        par = {y: p for y, p, c, yl, t, pr in rows if p is not None}
        price = {y: pr for y, p, c, yl, t, pr in rows if pr is not None}
        yield_of = {y: yl for y, p, c, yl, t, pr in rows}''')

_p("inputs.py", '''        for f in sorted(y for y, p, c, yl, t in rows if _is_term(t)):''',
   '''        for f in sorted(y for y, p, c, yl, t, pr in rows if _is_term(t)):''')

_p("inputs.py", '''        serial_yield = {y: yl for y, p, c, yl, t in rows
                        if yl is not None and y not in covered and not _is_term(t)}
        return {"coupon": coupon or None, "yield": serial_yield or None,
                "term": term_bonds or None, "par": par or None}''',
   '''        serial_yield = {y: yl for y, p, c, yl, t, pr in rows
                        if yl is not None and y not in covered and not _is_term(t)}
        return {"coupon": coupon or None, "yield": serial_yield or None,
                "term": term_bonds or None, "par": par or None,
                "price": price or None}''')

_p("inputs.py", '''        if block.get("par"):
            out[f"{pre}_par_schedule"] = block["par"]''',
   '''        if block.get("par"):
            out[f"{pre}_par_schedule"] = block["par"]
        if block.get("price"):
            out[f"{pre}_price_scale"] = block["price"]''')



# Thread the entered price through the sizer and every call site.
_p("debt_service.py", '''        par_schedule: Optional[dict] = None,
    ) -> BondTranche:''',
   '''        par_schedule: Optional[dict] = None,
        price_scale: Optional[dict] = None,
    ) -> BondTranche:''')

_p("debt_service.py", '''            coupon_scale=coupon_scale, yield_scale=yield_scale, term_bonds=term_bonds,
        )''',
   '''            coupon_scale=coupon_scale, yield_scale=yield_scale, term_bonds=term_bonds,
            price_scale=price_scale,
        )''')

_p("refunding.py", '''            par_schedule=cfg.senior_refunding_par_schedule,''',
   '''            par_schedule=cfg.senior_refunding_par_schedule,
            price_scale=cfg.senior_refunding_price_scale,''')

_p("scenarios.py", '''                par_schedule=cfg.senior_par_schedule,''',
   '''                par_schedule=cfg.senior_par_schedule,
                price_scale=cfg.senior_price_scale,''')

_p("main.py", '''        par_schedule=cfg.senior_par_schedule,''',
   '''        par_schedule=cfg.senior_par_schedule,
        price_scale=cfg.senior_price_scale,''')

_p("build_notebook.py", '''    par_schedule=cfg.senior_par_schedule,''',
   '''    par_schedule=cfg.senior_par_schedule,
    price_scale=cfg.senior_price_scale,''')



# ── Interest accrues from the DATED date; payments run in date order ────────
# Two defects, both exposed by the Wells Fargo / DBC run for Viridian Farm PID
# No. 2 (dated 9/30/2026, coupons 3/1 and 9/1):
#
#  1. Every coupon was a full half-year, `balance * rate / 2`, including the
#     first.  Interest has to accrue from the DATED date, so a bond not dated on
#     a coupon date owes a STUB first coupon.  DBC's first payment is $188,226 —
#     151/360 of a year on $7,180,000 at 6.250% — where the model charged the
#     full $224,375.  That flows into debt service, coverage, the CAPI deposit
#     and the price.
#  2. The two rows in a year were emitted mid-year coupon first, then the
#     principal date.  That is date order in Colorado (June, then December) but
#     REVERSED in Utah, where principal falls in March and the other coupon in
#     September — so every Utah schedule ran 9/1 before 3/1, and
#     `_apply_coupon_scale` charged September interest on a balance that March
#     had already paid down.
#
# Build the dates in order and carry a running balance and accrual date through.
# Worth pushing back to co_metro_model.
_p("debt_service.py", '''        # Interest accrues from the first interest date (June after delivery)
        # through final maturity.  Principal is paid each December.
        for y in range(t.delivery.year + 1, t.final_year + 1):
            # June coupon (interest only)
            jun = date(y, cfg.int_maturity, t.prin_day)
            jun_int = balance * rate / 2.0
            jun_capi = jun_int if (t.capi_end_year and y <= t.capi_end_year) else 0.0
            rows.append(PaymentRow(jun, 0.0, jun_int, capitalized_interest=jun_capi))

            # December coupon (interest + principal)
            dec = date(y, t.prin_month, t.prin_day)
            dec_int = balance * rate / 2.0
            dec_capi = dec_int if (t.capi_end_year and y <= t.capi_end_year) else 0.0
            p = principals.get(y, 0.0)
            surplus_rel = t.dsrf_deposit if (release_surplus and y == t.final_year) else 0.0
            # DSRF interest-earnings credit applies every year the fund is held,
            # INCLUDING the final year — the reserve earns interest through the
            # last year before its principal is released at maturity.  (This is
            # also what the sizer assumes, so final-year coverage ties to target.)
            earn = dsrf_earn
            rows.append(PaymentRow(
                dec, p, dec_int,
                capitalized_interest=dec_capi,
                dsrf_earnings=earn,
                surplus_release=surplus_rel,
            ))
            balance -= p

        return rows''',
   '''        from .pricing import days_30_360

        # Payment dates in DATE order.  Utah pays principal in March and its
        # other coupon in September, so the mid-year coupon falls AFTER the
        # principal date within a calendar year — the reverse of Colorado's
        # June/December.  Sorting is what makes both states right.
        dates = sorted({date(y, m, t.prin_day)
                        for y in range(t.delivery.year, t.final_year + 1)
                        for m in (cfg.int_maturity, t.prin_month)
                        if date(y, m, t.prin_day) > t.delivery})

        # Interest accrues from the DATED date, so a bond that is not dated on a
        # coupon date owes a stub first coupon rather than a full half-year.
        prev = t.delivery
        for d in dates:
            interest = balance * rate * days_30_360(prev, d) / 360.0
            capi = interest if (t.capi_end_year and d.year <= t.capi_end_year) else 0.0
            is_prin = d.month == t.prin_month
            p = principals.get(d.year, 0.0) if is_prin else 0.0
            surplus_rel = (t.dsrf_deposit if (release_surplus and is_prin
                                              and d.year == t.final_year) else 0.0)
            # DSRF interest-earnings credit applies every year the fund is held,
            # INCLUDING the final year — the reserve earns interest through the
            # last year before its principal is released at maturity.  (This is
            # also what the sizer assumes, so final-year coverage ties to target.)
            rows.append(PaymentRow(
                d, p, interest,
                capitalized_interest=capi,
                dsrf_earnings=(dsrf_earn if is_prin else 0.0),
                surplus_release=surplus_rel,
            ))
            balance -= p
            prev = d

        return rows''')

_p("debt_service.py", '''        principal_by_year = {p.payment_date.year: p.principal
                             for p in tranche.schedule if p.principal}
        for p in tranche.schedule:
            t = p.payment_date.year
            annual_int = sum(P * tranche.coupon_for(y)
                             for y, P in principal_by_year.items() if y >= t)
            p.interest = annual_int / 2.0
            in_capi = capi_end_year is not None and t <= capi_end_year
            p.capitalized_interest = p.interest if in_capi else 0.0''',
   '''        from .pricing import days_30_360

        # Maturities still outstanding, retired as the schedule pays them, so a
        # coupon that falls AFTER a principal date in the same year is charged on
        # the reduced balance.  Accrual runs from the last payment date — the
        # dated date for the first coupon, which is therefore a stub.
        outstanding = {p.payment_date.year: p.principal
                       for p in tranche.schedule if p.principal}
        prev = tranche.delivery
        for p in tranche.schedule:
            annual_int = sum(P * tranche.coupon_for(y) for y, P in outstanding.items())
            p.interest = annual_int * days_30_360(prev, p.payment_date) / 360.0
            in_capi = (capi_end_year is not None
                       and p.payment_date.year <= capi_end_year)
            p.capitalized_interest = p.interest if in_capi else 0.0
            if p.principal:
                outstanding.pop(p.payment_date.year, None)
            prev = p.payment_date''')



# The wrap must charge the same interest the schedule now pays.  A coupon that
# falls AFTER the principal date within a calendar year is charged on the
# balance that principal payment has already reduced.  Colorado pays principal
# on the LAST coupon of the year (December), so nothing follows it and the old
# `balance * rate` was right; Utah pays principal in March with September still
# to come, so half the year's interest is charged on the reduced balance.
# Solving target = p + balance*r - p*r*f for p gives the extra denominator.
_p("debt_service.py", '''        principal_years = list(range(first_principal_year, final_year + 1))''',
   '''        principal_years = list(range(first_principal_year, final_year + 1))
        # Share of a year's interest accruing AFTER the principal date.
        post_prin = 0.5 if cfg.int_maturity > cfg.prin_maturity else 0.0''')

_p("debt_service.py", '''                target = net_rev / coverage + dsrf_earn
                avail = target - interest_of(y, balance)
                if release_surplus and y == final_year:
                    avail += dsrf_deposit  # released DSRF pays down the final maturity
                p = max(0.0, math.floor(avail / 5000.0) * 5000.0)''',
   '''                target = net_rev / coverage + dsrf_earn
                annual_int = interest_of(y, balance)
                avail = target - annual_int
                if release_surplus and y == final_year:
                    avail += dsrf_deposit  # released DSRF pays down the final maturity
                # Principal retired on the principal date stops earning the
                # coupons that follow it later in the same year.
                eff_rate = (annual_int / balance) if balance else rate
                avail /= max(1e-9, 1.0 - eff_rate * post_prin)
                p = max(0.0, math.floor(avail / 5000.0) * 5000.0)''')



# Price the cash flows the bond actually pays.  Now that interest accrues from
# the dated date, the first coupon is a STUB whenever the bonds are not dated on
# a coupon date, and the price has to discount that stub rather than a full
# half-year.  On Viridian Farm PID No. 2 (dated 9/30/2026, 3/1/2056 maturity,
# 6.250%/6.625%) this is 95.1516 against DBC's 95.148 — the last 0.004.
# A bond reoffered at its coupon still prices at exactly 100: accrual and
# discounting share the same 30/360 clock, so the identity holds through a stub.
_p("pricing.py", '''    prices = []
    for red_date, red_price in scenarios:
        n = semiannual_periods(settlement, red_date, freq)
        if n > 0:
            prices.append(bond_price(n, coupon, ytm, red_price, freq))
    return min(prices) if prices else 100.0''',
   '''    prices = [price_from_dated_date(settlement, red_date, coupon, ytm, red_price, freq)
              for red_date, red_price in scenarios if red_date > settlement]
    return min(prices) if prices else 100.0


def price_from_dated_date(dated: date, redemption_date: date, coupon: float,
                          ytm: float, redemption: float = 100.0,
                          freq: int = 2) -> float:
    """
    Price per 100 for a NEW ISSUE settling on its dated date.

    Discounts the cash flows the bond actually pays: interest accrues from the
    dated date on a 30/360 basis, so the first coupon is a stub whenever the
    bonds are not dated on a coupon date, and every later coupon is a full
    period.  No accrued interest is subtracted — none changes hands when
    settlement is the dating.
    """
    step = 12 // freq
    dates, c_date = [], redemption_date
    while c_date > dated:
        dates.append(c_date)
        c_date = shift_months(c_date, -step)
    if not dates:
        return redemption
    dates.reverse()
    i = ytm / freq
    per = 360.0 / freq
    stub = days_30_360(dated, dates[0]) / per          # first period, in periods
    # The odd first period discounts at SIMPLE interest (1 + i*stub), the
    # convention Excel's ODDFPRICE uses.  It is not a nicety: interest accrues
    # simply across the stub, so discounting it compound would break the
    # identity that a bond reoffered at its coupon prices at exactly 100.
    def df(d):
        return 1.0 / ((1.0 + i * stub)
                      * (1.0 + i) ** (days_30_360(dated, d) / per - stub))
    price, prev = 0.0, dated
    for d in dates:
        price += (100.0 * coupon * days_30_360(prev, d) / 360.0) * df(d)
        prev = d
    return price + redemption * df(redemption_date)''')



# DBC truncates the price to three decimals and computes the OID from the
# TRUNCATED price, not the full-precision one: -348,373.60 is exactly
# 7,180,000 x (95.148 - 100) / 100.  Truncate (not round) so the printed price
# and the OID always agree, and so an entered price is used verbatim.
_p("debt_service.py", '''        maturity = date(year, self.prin_month, self.prin_day)
        return price_to_worst(self.delivery, maturity, self.coupon_for(year),
                              self.yield_for(year), self._call_scenarios())''',
   '''        maturity = date(year, self.prin_month, self.prin_day)
        return _truncate3(price_to_worst(self.delivery, maturity, self.coupon_for(year),
                                         self.yield_for(year), self._call_scenarios()))''')

_p("debt_service.py", '''            return price_to_worst(self.delivery, maturity, self.coupon_for(last), ty,
                                  self._call_scenarios())''',
   '''            return _truncate3(price_to_worst(self.delivery, maturity,
                                             self.coupon_for(last), ty,
                                             self._call_scenarios()))''')

_p("debt_service.py", '''def coupon_at(coupon_scale, term_bonds, rate: float, year: int) -> float:''',
   '''def _truncate3(price: float) -> float:
    """
    Price truncated — not rounded — to three decimals.

    The convention DBC prints and, more to the point, computes the OID from:
    its -348,373.60 is exactly 7,180,000 x (95.148 - 100) / 100.  Truncating
    here keeps the printed price and the premium/OID consistent, and leaves a
    price entered to three decimals untouched.
    """
    from decimal import Decimal, ROUND_DOWN
    # Clean floating-point noise before truncating.  A par bond's present value
    # lands on 99.999999999999 as often as 100.0, and truncating THAT would
    # invent a 0.001 discount on a bond reoffered at its coupon.
    cleaned = round(price, 9)
    return float(Decimal(str(cleaned)).quantize(Decimal("0.001"), rounding=ROUND_DOWN))


def coupon_at(coupon_scale, term_bonds, rate: float, year: int) -> float:''')



# ── Stale Colorado vocabulary in docstrings / section comments ───────────────

_p("memo.py", '''sources & uses), but states Colorado assumptions — mill levy (governing document cap +
Gallagher adjustment), primary residential / lot-inventory taxable ratios, the
reassessment on odd years, capitalized interest, the 3-prong DSRF, the''',
   '''sources & uses), but states Utah assumptions — the levy held under the
§ 17D-4-303 cap, the 45% primary residential exemption applied to homes and to
builder lot inventory, annual reassessment, capitalized interest, the 3-prong DSRF, the''')

_p("inputs.py", '''# ── Historical Residential Exemption (Colorado residential assessment) rates — reference ─────
# Residential taxable ratio by tax year / reassessment cycle.''',
   '''# ── Historical Utah residential exemption rates — reference ─────────────────
# Residential taxable ratio by tax year.''')

_p("report.py", '''    Values are shown in the year the AV is created; the model collects it on the
    Colorado lag (see Notes).''',
   '''    Values are shown in the year the value is created; the model collects it on
    the Utah lag — created in a calendar year, on the following 1 January roll,
    billed that 30 November, paying debt service the next 1 March (see Notes).''')


# ── main.py ──────────────────────────────────────────────────────────────────

_p("main.py", '''main.py — Entry point for the Utah Public Infrastructure District financial model.''',
   '''main.py — Entry point for the Utah Public Infrastructure District financial model
           (UCA 17D-4).''')

_p("main.py", '''    print("  Utah Public Infrastructure District Financial Model")
    print("  Python conversion (senior / subordinate liens + refunding)")''',
   '''    print("  Utah Public Infrastructure District Financial Model")
    print("  UCA 17D-4  |  senior / subordinate liens + refunding")''')

_p("main.py", '''        cfg, dev = ModelConfig(), DeveloperProjections()''',
   '''        from ut_pid_model import viridian_farm_projections
        cfg, dev = ModelConfig(), viridian_farm_projections()''')

_p("main.py", '''    print(f"    Mill levy (DS): {cfg.mill_levy_ds_target:.0f}  |  "
          f"residential taxable ratio: {cfg.resid_taxable_ratio:.2%}  |  "
          f"Biennial reassessment: {cfg.reassess_rate:.1%}")''',
   '''    print(f"    {cfg.city} City, {cfg.county} County, Utah")
    print(f"    DS mill levy: {cfg.effective_ds_mill_levy:.3f} "
          f"(cap {cfg.mill_levy_cap:.3f})  |  "
          f"residential taxable ratio: {cfg.resid_taxable_ratio:.2%}  |  "
          f"{cfg.reassess_frequency.lower()} reassessment: {cfg.reassess_rate:.1%}")
    for w in cfg.validate():
        print(f"    ! {w}")''')


# ── __init__.py ──────────────────────────────────────────────────────────────

_p("__init__.py", '''from .development import DeveloperProjections, ProductLine''',
   '''from .config import (PID_STATUTORY_LEVY_CAP, RESIDENTIAL_EXEMPTION,
                     UTAH_TAX_CALENDAR, lot_inventory_ratio,
                     residential_taxable_ratio_for)
from .development import (DeveloperProjections, ProductLine,
                          viridian_farm_projections)''')

_p("__init__.py", '''__all__ = [
    "ModelConfig",''',
   '''__all__ = [
    "ModelConfig",
    "PID_STATUTORY_LEVY_CAP",
    "RESIDENTIAL_EXEMPTION",
    "UTAH_TAX_CALENDAR",
    "residential_taxable_ratio_for",
    "lot_inventory_ratio",
    "viridian_farm_projections",''')


# ── Appended Utah-only code ──────────────────────────────────────────────────

APPENDS = {
    "development.py": '''

# ── The Viridian Farm PID No. 1 program (this repo's reference deal) ──────────

#: Product, units, base ASP, and the lot-delivery schedule priced 17 Sept 2024.
#: Homes close the year after their finished lots are delivered.
VIRIDIAN_FARM_PRODUCTS = [
    ("Rear-Load Townhome",  139, 365_620.0, {2024: 38, 2025: 60, 2026: 41}),
    ("Front-Load Townhome", 327, 394_910.0,
     {2024: 28, 2025: 48, 2026: 64, 2027: 96, 2028: 91}),
    ("Alley-Load Cottages",  30, 434_350.0, {2024: 24, 2025: 6}),
    ("Front-Load Cottages",  88, 492_150.0, {2024: 20, 2025: 36, 2026: 32}),
    ("8,000 Lots",           23, 563_750.0, {2023: 2, 2024: 21}),
    ("12,000 Lots",          67, 635_500.0, {2023: 2, 2024: 28, 2025: 30, 2026: 7}),
    ("18,000 Lots",          41, 709_813.0, {2024: 12, 2025: 22, 2026: 7}),
    ("21,000 Lots",           1, 761_063.0, {2024: 1}),
]


def viridian_farm_projections(asp_base_year: int = 2024) -> "DeveloperProjections":
    """Per-product absorption for Viridian Farm PID No. 1, exactly as priced."""
    products = [
        ProductLine(
            name=name, existing_units=units,
            lot_deliveries={y: u for y, u in sched.items() if u},
            home_closings={y + 1: u for y, u in sched.items() if u},
            asp_base=asp, asp_base_year=asp_base_year,
        )
        for name, units, asp, sched in VIRIDIAN_FARM_PRODUCTS
    ]
    return DeveloperProjections(products=products, total_lots=716)
''',
}

#: Patches allowed to miss (Colorado may have already changed them).
OPTIONAL_PATCHES: set[tuple[str, int]] = set()


# ── Port machinery ───────────────────────────────────────────────────────────

def rename_identifiers(text: str) -> str:
    for old, new in IDENTIFIER_RENAMES:
        text = re.sub(rf"\b{re.escape(old)}\b", new, text)
    return text


def substitute_labels(text: str) -> str:
    for old, new in LABEL_SUBS:
        text = text.replace(old, new)
    return text


def apply_patches(module: str, text: str, failures: list[str]) -> str:
    for i, (old, new) in enumerate(PATCHES.get(module, [])):
        if old == "@@TAIL@@":
            idx = text.find(CONFIG_TAIL_MARKER)
            if idx < 0:
                failures.append(f"{module}: config tail marker not found")
                continue
            text = text[:idx] + new
            continue
        if old not in text:
            if (module, i) not in OPTIONAL_PATCHES:
                failures.append(f"{module}: patch #{i} did not match:\n    "
                                + old.strip().splitlines()[0][:100])
            continue
        text = text.replace(old, new, 1)
    return text + APPENDS.get(module, "")


def port_file(src: Path, module: str, failures: list[str]) -> str:
    text = src.read_text()
    text = rename_identifiers(text)
    text = substitute_labels(text)
    text = apply_patches(module, text, failures)
    return text


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--source", required=True,
                    help="path to the CO-Metro-District-Model checkout")
    ap.add_argument("--check", action="store_true",
                    help="report differences without writing")
    args = ap.parse_args(argv)

    src_root = Path(args.source).expanduser().resolve()
    src_pkg = src_root / "co_metro_model"
    if not src_pkg.is_dir():
        print(f"error: {src_pkg} not found", file=sys.stderr)
        return 2

    dst_pkg = REPO / "ut_pid_model"
    failures: list[str] = []
    staged: list[tuple[Path, str, str]] = []

    for module in MODULES:
        staged.append((dst_pkg / module, f"ut_pid_model/{module}",
                       port_file(src_pkg / module, module, failures)))
    for script in SCRIPTS:
        staged.append((REPO / script, script,
                       port_file(src_root / script, script, failures)))

    # Nothing is written unless every patch applied — a half-ported model that
    # still reads "Gallagherized" is worse than no port at all.
    if failures:
        print("PATCHES THAT DID NOT APPLY — Colorado has moved, re-read these:",
              file=sys.stderr)
        for f in failures:
            print(f"  • {f}", file=sys.stderr)
        print("\nNothing was written.", file=sys.stderr)
        return 1

    changed: list[str] = []
    for dst, name, ported in staged:
        if dst.exists() and dst.read_text() == ported:
            continue
        changed.append(name)
        if args.check:
            if dst.exists():
                diff = difflib.unified_diff(
                    dst.read_text().splitlines(), ported.splitlines(),
                    fromfile=f"a/{name}", tofile=f"b/{name}", lineterm="", n=2)
                print("\n".join(list(diff)[:60]))
        else:
            dst.write_text(ported)

    if changed:
        print(("would update " if args.check else "updated ") + f"{len(changed)} file(s):")
        for c in changed:
            print(f"  {c}")
    else:
        print("already up to date with Colorado")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
