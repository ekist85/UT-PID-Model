"""
config.py — every input on the "Inputs - First" tab, as a dataclass.

Layout parity
-------------
The row numbers in `INPUT_ROWS` below are the *actual* rows on the
"Inputs - First" tab, and they match the Colorado metro district template
one-for-one.  `ut_pid_model.workbook` writes the tab straight from this table,
so the Python model and the Excel output can never drift apart.

Colorado → Utah
---------------
Three Colorado concepts have no Utah analogue and are re-labelled (the range
names carry the Utah term, with the Colorado name kept as an alias so the two
workbooks stay diff-able):

    Colorado                                Utah
    ------------------------------------    ------------------------------------
    Gallagher residential assessment rate    Primary Residential Exemption
      (TABOR_PRIOR / TABOR_CURRENT)            (RESID_TAXABLE_RATIO = 55%)
    Gallagherization of the service plan     n/a — Utah caps are stated as a
      mill levy                                fixed rate per dollar of value
    Specific Ownership Tax                   Personal property uniform fee
      (motor vehicles, Art. X §6)              (UCA 59-2-405, distributed pro rata)

Defaults are the Viridian Farm PID No. 1 (Salem City, Utah County) financing
priced 9/17/2024 — the deal in this repo — so a bare `ModelConfig()` reproduces
that transaction.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from .xlfin import edate


# ── Statutory constants (Utah) ────────────────────────────────────────────────

#: UCA 17D-4-303 — a PID's levy for all purposes may not exceed this rate per
#: dollar of taxable value.  Does not bind a levy for voted GO bonds.
PID_STATUTORY_LEVY_CAP = 0.015          # 15.000 mills

#: UCA 59-2-103 — primary residential property is taxed on 55% of fair market
#: value (a 45% exemption), for up to one acre of land per residential unit.
RESIDENTIAL_EXEMPTION = 0.45
DEFAULT_RESID_TAXABLE_RATIO = 1.0 - RESIDENTIAL_EXEMPTION   # 0.55


@dataclass
class ModelConfig:
    """Central parameter store.  Override any field to run a what-if."""

    # ── Basic inputs (rows 6-17) ──────────────────────────────────────────────
    district_name: str = "Viridian Farm Public Infrastructure District No. 1"
    city: str = "Salem"
    county: str = "Utah"
    developer: str = "D.R. Horton"
    scenario_label: str = ""
    second_financing: str = "No"
    refund_financing: str = "No"

    # ── Fees (rows 19-26) ─────────────────────────────────────────────────────
    coi: float = 400_000.0
    coi_refunding: float = 200_000.0
    uwd_senior: float = 0.010
    uwd_sub: float = 0.015
    uwd_senior_refunding: float = 0.005
    #: Utah counties recover assessing/collecting through a separate statewide
    #: levy on property (UCA 59-2-1602), not a haircut on the district's
    #: distribution, so this is 0.00 for Utah where Colorado runs ~1.50%.
    county_treasurer_fee: float = 0.0
    trustee_fee: float = 4_000.0
    trustee_fee_sub: float = 3_000.0

    # ── Structuring assumptions (rows 30-69) ──────────────────────────────────
    delivery: date = date(2024, 9, 26)
    capi: str = "Yes"
    capi_term_months: int = 36
    surplus_on_off: str = "Yes"
    surplus_release_sizing: str = "Yes"
    surplus_target_multiple: float = 0.0      # × MADS held in the surplus fund
    sub_sizing_threshold: float = 10_000.0
    #: Force a subordinate par (e.g. a round $1,000,000) instead of letting the
    #: model size the largest amount the residual cashflow retires in full.
    sub_par_override: Optional[float] = None
    premium_call_years: int = 5
    premium_call_price: float = 103.0
    par_call_years: int = 7
    final_mat_yrs: int = 29
    final_mat_sub_yrs: int = 30
    final_mat_yrs_refunding: int = 30
    #: Utah PID debt service is structured around a 30 November tax due date, so
    #: principal falls on 1 March (Colorado's Feb/June collections → 1 December).
    prin_maturity: int = 3
    prin_maturity_day_senior: int = 1
    prin_maturity_day_sub: int = 15
    ig_rated: str = "No"
    senior_rate_ig: float = 0.05000
    senior_rate_nr: float = 0.05875
    sub_rate_ig: float = 0.07000
    sub_rate_nr: float = 0.08125
    senior_refunding_interest_rate: float = 0.045
    dsc_senior_lien_bonds: float = 1.30
    dsc_sub_lien_bonds: float = 1.00
    dsc_refunding_bonds: float = 1.20

    # ── Projection assumptions (rows 71-106) ──────────────────────────────────
    first_year: int = 2023
    resid_delivery_year: date = date(2024, 1, 1)
    inflation_rate: float = 0.03
    inflation_rate_comm_sales: float = 0.01
    inflation_step_years: int = 1
    inflation_step_start_year: date = date(2021, 12, 1)
    #: Utah county assessors revalue annually (UCA 59-2-303.1), so the default
    #: reassessment cadence is "Annual".  Set to "Biennial" to mirror Colorado.
    reassess_frequency: str = "Annual"
    reassess_rate_resid: float = 0.01
    reassess_rate_resid_sub: float = 0.01
    reassess_rate_comm: float = 0.02
    #: Utah primary residential exemption.  `resid_taxable_ratio_prior` exists
    #: only so the Colorado "gallagherization" arithmetic stays available; with
    #: `gallagherization = "No"` (the Utah default) it is inert.
    resid_taxable_ratio_prior: float = 0.55
    resid_taxable_ratio: float = DEFAULT_RESID_TAXABLE_RATIO
    tax_collect_mill_prc: float = 0.98
    #: Personal-property uniform fee (UCA 59-2-405) allocated to the district in
    #: the same proportion as real property tax.  Viridian excluded it — 0.00.
    uniform_fee_prc: float = 0.0
    uniform_fee_av_threshold: float = 0.0
    interest_earn_rate: float = 0.025
    gallagherization: str = "No"
    #: Rate per $1,000 of taxable value.  0.003/dollar = 3.000 mills.
    mill_levy_governing_doc: float = 3.0
    mill_levy_comm: float = 0.0
    mill_levy_ops_target: float = 0.0
    mill_levy_cap_total: float = 0.0
    contribution_rate: float = 0.0
    #: Annual district administration (accounting, audit, legal, assessor
    #: filings), inflated at `admin_cost_growth`.  Colorado books this as the
    #: "O&M carveout"; a Utah PID has no separate operations levy, so the cost
    #: is charged against pledged revenue instead.
    admin_cost_base: float = 53_060.0
    admin_cost_av_limit: float = 0.0
    admin_cost_growth: float = 0.02
    #: First assessment year that carries district costs (admin + trustee).
    #: Defaults to two years after closing: the first tax roll set after the
    #: bonds are outstanding is billed the following November, so year 2 is the
    #: first one with a full year of collections to charge against.
    district_cost_start_year: Optional[int] = None
    #: System Development Fee ("SDF") per residential unit, if any.
    system_development_fee: float = 0.0
    state_assessed: float = 0.0
    centrally_assessed_value: float = 0.0
    resid_new_value_add: Optional[str] = None      # defaults from `ig_rated`
    resid_new_value_add_sub: str = "Yes"
    comm_new_value_add: str = "No"
    comm_new_value_add_sub: str = "No"
    centrally_assessed_senior: str = "No"
    centrally_assessed_sub: str = "Yes"

    # ── Developer assumptions (rows 113-136) ──────────────────────────────────
    asp: list[float] = field(default_factory=lambda: [
        365_620.0, 394_910.0, 434_350.0, 492_150.0,
        563_750.0, 635_500.0, 709_813.0, 761_063.0,
        0.0, 0.0, 0.0, 0.0,
    ])
    product_labels: list[str] = field(default_factory=lambda: [
        "Rear-Load Townhome", "Front-Load Townhome", "Alley-Load Cottages",
        "Front-Load Cottages", "8,000 Lots", "12,000 Lots",
        "18,000 Lots", "21,000 Lots", "", "", "", "",
    ])
    hypothetical_scenario: str = "No"
    lot_delivery_scenario: float = 1.0
    absorption_scenario: float = 1.0
    home_first_close_months: int = 6               # after delivery
    commercial_lag_years: int = 1
    home_lot_delivery_lead_months: int = 6         # before first closing
    platted_comm_lot_value: float = 0.10
    #: Finished-lot market value as a share of the eventual home ASP.
    platted_lot_value: float = 0.10
    #: Taxable-value ratio applied to developer lot inventory.  Utah allows the
    #: residential exemption on unoccupied property the assessor determines will
    #: become a primary residence (Utah Admin. Code R884-24P-52), so this is the
    #: 55% residential ratio.  Colorado uses its 29% vacant-land rate.
    developed_lot_value: float = DEFAULT_RESID_TAXABLE_RATIO
    centrally_assessed_ratio: float = 0.875

    # ── Value-lag convention ──────────────────────────────────────────────────
    #: Years between value creation and the tax roll it first appears on.
    #: Utah assesses on 1 January and bills the same year (due 30 November), so
    #: one year.  Colorado's biennial cycle effectively runs two.
    value_lag_years: int = 1

    # ── Derived structuring dates ─────────────────────────────────────────────

    def __post_init__(self) -> None:
        if self.resid_new_value_add is None:
            self.resid_new_value_add = "No" if self.ig_rated == "Yes" else "Yes"
        if self.district_cost_start_year is None:
            self.district_cost_start_year = self.delivery.year + 2

    @property
    def int_maturity(self) -> int:
        """Interest-only month, six months off the principal month."""
        return self.prin_maturity - 6 if self.prin_maturity > 6 else self.prin_maturity + 6

    @property
    def senior_interest_rate(self) -> float:
        return self.senior_rate_ig if self.ig_rated == "Yes" else self.senior_rate_nr

    @property
    def sub_interest_rate(self) -> float:
        return self.sub_rate_ig if self.ig_rated == "Yes" else self.sub_rate_nr

    @property
    def first_int(self) -> date:
        """First senior interest date — the earlier coupon date after delivery."""
        d = self.delivery
        prin = date(d.year, self.prin_maturity, self.prin_maturity_day_senior)
        intr = date(d.year, self.int_maturity, self.prin_maturity_day_senior)
        if min(prin, intr) > d:
            nxt = edate(d, 6)
            return min(date(nxt.year, self.prin_maturity, self.prin_maturity_day_senior),
                       date(nxt.year, self.int_maturity, self.prin_maturity_day_senior))
        return max(prin, intr) if max(prin, intr) > d else edate(min(prin, intr), 12)

    @property
    def first_int_sub(self) -> date:
        d = self.delivery
        prin = date(d.year, self.prin_maturity, self.prin_maturity_day_sub)
        intr = date(d.year, self.int_maturity, self.prin_maturity_day_sub)
        if min(prin, intr) > d:
            nxt = edate(d, 6)
            return min(date(nxt.year, self.prin_maturity, self.prin_maturity_day_sub),
                       date(nxt.year, self.int_maturity, self.prin_maturity_day_sub))
        return max(prin, intr) if max(prin, intr) > d else edate(min(prin, intr), 12)

    @property
    def premium_call_first(self) -> date:
        d = self.delivery
        base = min(date(d.year, self.prin_maturity, self.prin_maturity_day_senior),
                   date(d.year, self.int_maturity, self.prin_maturity_day_senior))
        return edate(base, self.premium_call_years * 12)

    @property
    def par_call_first(self) -> date:
        d = self.delivery
        base = min(date(d.year, self.prin_maturity, self.prin_maturity_day_senior),
                   date(d.year, self.int_maturity, self.prin_maturity_day_senior))
        return edate(base, self.par_call_years * 12)

    @property
    def delivery_refunding(self) -> date:
        return self.premium_call_first

    @property
    def first_int_refunding(self) -> date:
        d = self.delivery_refunding
        prin = date(d.year, self.prin_maturity, self.prin_maturity_day_senior)
        intr = date(d.year, self.int_maturity, self.prin_maturity_day_senior)
        if min(prin, intr) <= d:
            nxt = edate(d, 6)
            return min(date(nxt.year, self.prin_maturity, self.prin_maturity_day_senior),
                       date(nxt.year, self.int_maturity, self.prin_maturity_day_senior))
        return min(prin, intr)

    @property
    def capi_first_draw(self) -> date:
        return self.first_int

    @property
    def capi_end_date(self) -> date:
        e = edate(self.delivery, self.capi_term_months)
        return date(e.year, min(self.prin_maturity, e.month), self.prin_maturity_day_senior)

    @property
    def surplus_fund_release_date(self) -> date:
        a = edate(date(self.delivery.year, self.prin_maturity,
                       self.prin_maturity_day_senior), 12 * self.final_mat_yrs)
        b = edate(date(self.first_int.year, self.prin_maturity,
                       self.prin_maturity_day_senior), 12 * self.final_mat_yrs)
        return max(a, b)

    @property
    def surplus_fund_release_date_refunding(self) -> date:
        a = edate(date(self.delivery_refunding.year, self.prin_maturity,
                       self.prin_maturity_day_senior), 12 * self.final_mat_yrs_refunding)
        b = edate(date(self.first_int.year, self.prin_maturity,
                       self.prin_maturity_day_senior), 12 * self.final_mat_yrs_refunding)
        return max(a, b)

    @property
    def end_bal_accrued_sub_bonds(self) -> date:
        return edate(date(self.delivery.year, self.prin_maturity,
                          self.prin_maturity_day_sub), 12 * self.final_mat_sub_yrs)

    @property
    def home_first_close_date(self) -> date:
        return edate(self.delivery, self.home_first_close_months)

    @property
    def home_lot_delivery_date(self) -> date:
        return edate(self.home_first_close_date, -self.home_lot_delivery_lead_months)

    @property
    def mill_levy_ds_cap(self) -> float:
        if self.gallagherization == "Yes" and self.resid_taxable_ratio:
            return self.mill_levy_governing_doc / (
                self.resid_taxable_ratio / self.resid_taxable_ratio_prior)
        return self.mill_levy_governing_doc

    @property
    def mill_levy_ds_target(self) -> float:
        return self.mill_levy_ds_cap

    @property
    def analysis(self) -> str:
        return (f"Development Projections at {self.mill_levy_ds_target:.3f} "
                f"Mills for Debt Service")

    @property
    def senior_bonds_series(self) -> str:
        return f"Series {self.delivery.year}A"

    @property
    def sub_bonds_series(self) -> str:
        return f"Series {self.delivery.year}B"

    @property
    def refund_bonds_series(self) -> str:
        return f"Series {self.premium_call_first.year}"

    @property
    def frac(self) -> float:
        from .xlfin import yearfrac
        return yearfrac(self.delivery, self.first_int)

    @property
    def frac_sub(self) -> float:
        from .xlfin import yearfrac
        return yearfrac(self.delivery, self.first_int_sub)

    def validate(self) -> list[str]:
        """Return a list of statutory / structural warnings (empty is good)."""
        warnings: list[str] = []
        total_levy = (self.mill_levy_ds_target + self.mill_levy_ops_target
                      + self.contribution_rate)
        statutory_mills = PID_STATUTORY_LEVY_CAP * 1000.0
        if total_levy > statutory_mills:
            warnings.append(
                f"Total district levy of {total_levy:.3f} mills exceeds the "
                f"UCA 17D-4-303 cap of {statutory_mills:.3f} mills "
                f"(0.015 per dollar of taxable value).")
        if self.mill_levy_cap_total and total_levy > self.mill_levy_cap_total:
            warnings.append(
                f"Total district levy of {total_levy:.3f} mills exceeds the "
                f"governing-document cap of {self.mill_levy_cap_total:.3f} mills.")
        if self.gallagherization == "Yes":
            warnings.append(
                "Gallagherization is a Colorado mechanism; Utah levy caps are "
                "stated as a fixed rate per dollar of taxable value and are not "
                "adjusted for changes in the residential exemption.")
        if not 0 < self.resid_taxable_ratio <= 1:
            warnings.append("Residential taxable ratio must be between 0 and 1.")
        return warnings
