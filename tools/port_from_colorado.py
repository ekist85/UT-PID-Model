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
    "debt_service.py", "subordinate.py", "sources_uses.py", "refunding.py",
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
    ("om_carveout_av_limit", "admin_cost_av_limit"),
    ("om_carveout", "admin_cost"),
    ("om_growth_rate", "admin_growth_rate"),
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
    ("OM_CARVEOUT_AV_LIMIT", "ADMIN_COST_AV_LIMIT"),
    ("OM_CARVEOUT", "ADMIN_COST"),
    ("OM_GROWTH_RATE", "ADMIN_GROWTH_RATE"),
    ("STRICT_BIENNIAL_AV", "HOLD_VALUE_FLAT"),
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

_p("config.py", '''    first_year: int = 2023              # FIRST_YEAR (first "Summary" year)
    inflation_start_year: int = 2024    # INFLATION_START_YEAR (home-price inflation''',
   '''    first_year: int = 2022              # FIRST_YEAR (first "Summary" year)
    inflation_start_year: int = 2025    # INFLATION_START_YEAR (home-price inflation''')

_p("config.py", '''    reassess_rate: float = 0.02          # REASSESS_RATE (residential)
    reassess_comm_rate: float = 0.02     # REASSESS_COMM_RATE

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

_p("config.py", '''    admin_cost: float = 0              # ADMIN_COST
    admin_cost_av_limit: float = 0     # ADMIN_COST_AV_LIMIT
    admin_growth_rate: float = 0.02        # ADMIN_GROWTH_RATE''',
   '''    # Annual district administration — accounting, audit, legal, assessor and
    # continuing-disclosure filings — charged against pledged revenue.  Colorado
    # books this as the O&M carveout against a separate operations levy; a Utah
    # PID typically has no operations levy, so the cost lands here.
    admin_cost: float = 53_060         # ADMIN_COST (base year)
    admin_cost_av_limit: float = 0     # ADMIN_COST_AV_LIMIT
    admin_growth_rate: float = 0.02    # ADMIN_GROWTH_RATE (inflates the base)
    # First collection year that carries district costs — administration and the
    # trustee fees.  None ⇒ two years after closing: the first roll set with the
    # bonds outstanding is billed that November, so year 2 is the first with a
    # full year of collections to charge against.
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
   '''    def district_costs(self, collection_year: int) -> tuple[float, float, float]:
        """
        (administration, senior trustee fee, subordinate trustee fee) charged
        against pledged revenue in ``collection_year``.

        Nothing is charged before ``district_cost_start_year``; from then on the
        administration base inflates at ``admin_growth_rate`` and the trustee
        fees are flat.
        """
        start = self.district_cost_start_year or (self.delivery.year + 2)
        if collection_year < start:
            return 0.0, 0.0, 0.0
        admin = self.admin_cost * (1 + self.admin_growth_rate) ** (collection_year - start)
        return admin, self.trustee_fee, self.trustee_fee_sub

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
                mill_revenue + uniform_fee - collection_fee - cfg.trustee_fee - cfg.admin_cost
            )''',
   '''            # Net revenue available for SENIOR lien debt service (AX):
            #   mill + uniform fee - county collection cost - senior trustee fee
            #   - annual district administration (inflated, and not charged
            #     before the district is up and running)
            collection_fee = mill_revenue * cfg.county_collection_fee
            admin_cost, trustee_fee, trustee_fee_sub = cfg.district_costs(collect)
            if cfg.admin_cost_av_limit and total_av > cfg.admin_cost_av_limit:
                admin_cost = 0.0
            net_senior_revenue = (
                mill_revenue + uniform_fee - collection_fee - trustee_fee - admin_cost
            )''')

_p("summary.py", '''            net_sub_revenue = mill_revenue + uniform_fee - cfg.trustee_fee_sub''',
   '''            net_sub_revenue = mill_revenue + uniform_fee - trustee_fee_sub''')


# ── subordinate.py ───────────────────────────────────────────────────────────

# Upstream rounds the solved par to the NEAREST $1,000, which can round up past
# the largest amount the residual cashflow actually retires — the base case then
# reports "fully repaid: False" by a few thousand dollars.  Round down instead;
# worth pushing back to co_metro_model.
_p("subordinate.py", '''        lo, hi = 0.0, max_par
        if repaid(hi):
            return round(hi / 1000.0) * 1000.0
        for _ in range(40):
            mid = (lo + hi) / 2.0
            if repaid(mid):
                lo = mid
            else:
                hi = mid
        return round(lo / 1000.0) * 1000.0''',
   '''        lo, hi = 0.0, max_par
        if repaid(hi):
            return math.floor(hi / 1000.0) * 1000.0
        for _ in range(40):
            mid = (lo + hi) / 2.0
            if repaid(mid):
                lo = mid
            else:
                hi = mid
        # Round DOWN: rounding to the nearest $1,000 can land above the largest
        # par the residual surplus retires, which would leave the note short at
        # final maturity.
        return math.floor(lo / 1000.0) * 1000.0''')

_p("subordinate.py", '''from __future__ import annotations

from dataclasses import dataclass, field''',
   '''from __future__ import annotations

import math
from dataclasses import dataclass, field''')


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
            dated = cfg.delivery
            if y < dated.year:
                year_frac = 0.0
            elif y == dated.year:
                year_frac = max(0.0, _yearfrac_30360(dated, pay_date))
            else:
                year_frac = 1.0''',
   '''            pay_date = date(y, cfg.prin_maturity, cfg.prin_maturity_day_sub)
            prior_pay = date(y - 1, cfg.prin_maturity, cfg.prin_maturity_day_sub)
            dated = cfg.delivery
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
        om = -cfg.admin_cost if cfg.admin_cost else 0.0''',
   '''        # District costs come from the same helper the revenue engine uses,
        # so the report and the sizing can never disagree.
        _admin, _trustee, _subtrustee = cfg.district_costs(r.collection_year)
        trust = -_trustee
        subtrust = -_subtrustee
        om = -_admin''')

_p("report.py", '''        ("om", "− O&M\\nCarveout", 11),''',
   '''        ("om", "− District\\nAdmin", 11),''')

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

_p("inputs.py", '''    ("O&M", "O&M Carveout", "ADMIN_COST", "admin_cost", "float", "$ carved out of pledged revenue each year for operations"),
]''',
   '''    ("District Costs", "Annual District Administration", "ADMIN_COST", "admin_cost", "float", "$ per year, charged against pledged revenue"),
    ("District Costs", "District Administration Growth Rate", "ADMIN_GROWTH_RATE", "admin_growth_rate", "pct", "annual inflation on the administration base"),
    ("District Costs", "District Administration Taxable Value Limit", "ADMIN_COST_AV_LIMIT", "admin_cost_av_limit", "float", "$ — above this taxable value the charge stops; 0 ⇒ no limit"),
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



# ── District operations & maintenance expense ────────────────────────────────
# Colorado's template exposes a single "O&M Carveout" (renamed here to the
# district administration line).  A Utah PID typically has no separate
# operations levy, so its operating budget has to be funded out of the same
# pledged revenue that services the bonds — which makes the O&M expense a real
# input to sizing, not a footnote.  Two rows: a starting expense and the
# inflation that grows it.
#
# Unlike the administration carveout, this is netted from the revenue available
# to BOTH liens.  The subordinate lien's own revenue is measured as
# `net_sub_revenue - net_senior_revenue`, so a cost netted from the senior side
# alone is handed straight to the sub — which would make an O&M expense *raise*
# subordinate capacity.  Money the district actually spends is available to
# neither bond.

_p("config.py", '''    admin_growth_rate: float = 0.02    # ADMIN_GROWTH_RATE (inflates the base)''',
   '''    admin_growth_rate: float = 0.02    # ADMIN_GROWTH_RATE (inflates the base)
    # District operations & maintenance — landscaping, parks and trails, snow
    # removal, street lighting, utilities on the district improvements.  A Utah
    # PID rarely carries a separate operations levy, so this is paid out of the
    # same pledged revenue as debt service and comes off the top: it is netted
    # from the revenue available to the senior AND the subordinate lien.
    # Defaults to zero — an operating budget is a district-specific number, not
    # something to assume.
    om_expense: float = 0.0            # OM_EXPENSE (base year, $ per year)
    om_growth_rate: float = 0.03       # OM_GROWTH_RATE (inflates the base)''')

_p("config.py", '''    @property
    def mill_levy_cap(self) -> float:''',
   '''    def om_expense_for(self, collection_year: int) -> float:
        """
        District operations & maintenance charged against pledged revenue in
        ``collection_year``.

        Nothing is charged before ``district_cost_start_year`` — the same start
        the administration and trustee fees use — and from then on the base
        inflates at ``om_growth_rate``.
        """
        if not self.om_expense:
            return 0.0
        start = self.district_cost_start_year or (self.delivery.year + 2)
        if collection_year < start:
            return 0.0
        return self.om_expense * (1 + self.om_growth_rate) ** (collection_year - start)

    @property
    def mill_levy_cap(self) -> float:''')

_p("summary.py", '''    net_senior_revenue: float      # AX
    net_sub_revenue: float         # BO''',
   '''    net_senior_revenue: float      # AX
    net_sub_revenue: float         # BO
    om_expense: float = 0.0        # district O&M, netted from both liens''')

_p("summary.py", '''            collection_fee = mill_revenue * cfg.county_collection_fee
            admin_cost, trustee_fee, trustee_fee_sub = cfg.district_costs(collect)''',
   '''            collection_fee = mill_revenue * cfg.county_collection_fee
            admin_cost, trustee_fee, trustee_fee_sub = cfg.district_costs(collect)
            # District O&M comes off the top — see ModelConfig.om_expense_for.
            om_expense = cfg.om_expense_for(collect)''')

_p("summary.py", '''            net_senior_revenue = (
                mill_revenue + uniform_fee - collection_fee - trustee_fee - admin_cost
            )''',
   '''            net_senior_revenue = (
                mill_revenue + uniform_fee - collection_fee - trustee_fee - admin_cost
                - om_expense
            )''')

_p("summary.py", '''            net_sub_revenue = mill_revenue + uniform_fee - trustee_fee_sub''',
   '''            net_sub_revenue = mill_revenue + uniform_fee - trustee_fee_sub - om_expense''')

_p("summary.py", '''                net_sub_revenue=net_sub_revenue,
            )''',
   '''                net_sub_revenue=net_sub_revenue,
                om_expense=om_expense,
            )''')

_p("inputs.py", '''    ("District Costs", "First Year District Costs Are Charged", "DISTRICT_COST_START_YEAR",''',
   '''    ("District Costs", "Starting O&M Expense", "OM_EXPENSE", "om_expense", "float", "$ per year of district operations & maintenance, netted from the revenue available to both liens"),
    ("District Costs", "O&M Expense Growth Rate", "OM_GROWTH_RATE", "om_growth_rate", "pct", "annual inflation on the O&M base"),
    ("District Costs", "First Year District Costs Are Charged", "DISTRICT_COST_START_YEAR",''')

_p("report.py", '''             ("om", "− District\\nAdmin", 11),
             ("net", "Net Revenue\\n(senior sizing)", 15)]''',
   '''             ("om", "− District\\nAdmin", 11),
             ("omexp", "− District\\nO&M", 11),
             ("net", "Net Revenue\\n(senior sizing)", 15)]''')

_p("report.py", '''            "trust": trust, "subtrust": subtrust, "om": om, "net": r.net_senior_revenue,''',
   '''            "trust": trust, "subtrust": subtrust, "om": om,
            "omexp": -r.om_expense, "net": r.net_senior_revenue,''')



# The O&M tab showed only the operations-levy revenue.  Now that the district
# carries a modelled O&M expense, show it alongside — and the surplus/(deficit),
# which is the number that says whether the operations levy actually covers the
# operating budget or whether the debt levy is carrying it.
_p("report.py", '''    ops_mill = cfg.mill_levy_ops_target
    coll = cfg.tax_collect_mill_prc
    _title(ws, [cfg.pid_name, "Operations & Maintenance (O&M) Revenue Projection",
                f"Operations mill levy {ops_mill:.3f} mills @ {coll:.1%} collection"], 6)
    hdrs = [(1, "Collection\\nYear", 12), (2, "Total\\nTaxable Value", 16),
            (3, "Operations\\nMill Levy", 13),
            (4, f"Total Collections\\n@ {coll:.1%}", 16),
            (5, f"Uniform Fee\\n@ {cfg.uniform_fee_prc:.0%}", 14),
            (6, "Total Available\\nfor O&M", 16)]''',
   '''    ops_mill = cfg.mill_levy_ops_target
    coll = cfg.tax_collect_mill_prc
    _title(ws, [cfg.pid_name, "Operations & Maintenance (O&M) Revenue and Expense",
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
        ("note", "District administration (base / growth)",
            f"${cfg.admin_cost:,.0f} / {_pct(cfg.admin_growth_rate)}"),
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
