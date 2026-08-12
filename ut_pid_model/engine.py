"""
engine.py — the calculation core.

Reproduces, in order, the workbook's dependency chain:

    Residential / Comm Development  →  Summary taxable value (cols E–AE)
                                    →  Summary revenue      (cols AG–BO)
                                    →  Senior Lien DS sizing
                                    →  Surplus fund / CAPI
                                    →  Sub Lien DS sizing
                                    →  Sources and Uses, Ops, DBC Output

The workbook resolves the loop between the surplus-fund deposit, capitalized
interest and the senior sizing with Excel's iterative calculation.  Here that is
an explicit fixed-point loop in `Model.run`, which converges in a handful of
passes and is checked rather than assumed.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from .config import ModelConfig
from .development import DevelopmentProjections
from .xlfin import (days360, edate, npv, price, round_down_to, safe_div,
                    solve_tic, truncate, yearfrac)

PRINCIPAL_INCREMENT = 5_000.0
SUB_PRINCIPAL_INCREMENT = 1_000.0


# ── Row containers ────────────────────────────────────────────────────────────

@dataclass
class SummaryRow:
    """One assessment year — the Summary tab, row by row."""
    index: int
    assessment_date: date                 # col C
    tax_revenue_date: date                # col D

    # Residential taxable value (cols E–Q)
    existing_reassessment: float = 0.0    # E
    existing_market_value: float = 0.0    # F
    existing_taxable_value: float = 0.0   # G
    lot_units: float = 0.0                # H
    lot_market_value: float = 0.0         # I
    lot_taxable_value: float = 0.0        # J
    residential_units: float = 0.0        # K
    new_home_reassessment: float = 0.0    # L
    new_home_market_value: float = 0.0    # M
    new_home_taxable_value: float = 0.0   # N
    prior_roll_value: float = 0.0         # P
    residential_taxable_value: float = 0.0  # Q

    # Commercial (cols R–Z)
    comm_sf_delivered: float = 0.0        # R
    comm_reassessment_platted: float = 0.0  # S
    comm_market_value_platted: float = 0.0  # T
    comm_taxable_value_platted: float = 0.0  # U
    comm_sf_sold: float = 0.0             # V
    comm_reassessment_sold: float = 0.0   # W
    comm_market_value_sold: float = 0.0   # X
    comm_taxable_value_sold: float = 0.0  # Y
    comm_taxable_value: float = 0.0       # Z

    # Centrally / state assessed (cols AA–AE)
    centrally_assessed_market: float = 0.0   # AA
    centrally_assessed_taxable: float = 0.0  # AD
    state_assessed_taxable: float = 0.0      # AE

    # Senior revenue (cols AG–AX)
    senior_taxable_value: float = 0.0     # AG
    senior_mill_levy: float = 0.0         # AH
    senior_levy_collections: float = 0.0  # AI
    senior_uniform_fee: float = 0.0       # AJ
    senior_comm_taxable_value: float = 0.0    # AK
    senior_comm_mill_levy: float = 0.0        # AL
    senior_comm_collections: float = 0.0      # AM
    senior_comm_uniform_fee: float = 0.0      # AN
    taxable_sales: float = 0.0            # AO
    pif_revenue: float = 0.0              # AP
    system_development_fee: float = 0.0   # AQ
    senior_total_revenue: float = 0.0     # AR
    county_fee: float = 0.0               # AT
    trustee_fee: float = 0.0              # AU
    admin_costs: float = 0.0              # AV
    senior_net_revenue: float = 0.0       # AX

    # Subordinate revenue (cols AZ–BO)
    sub_taxable_value: float = 0.0        # AZ
    sub_levy_collections: float = 0.0     # BB
    sub_uniform_fee: float = 0.0          # BC
    sub_comm_collections: float = 0.0     # BF
    sub_comm_uniform_fee: float = 0.0     # BG
    sub_total_revenue: float = 0.0        # BK
    sub_trustee_fee: float = 0.0          # BM
    sub_net_revenue: float = 0.0          # BO

    # Debt service and surplus (cols BQ–CE)
    senior_net_debt_service: float = 0.0  # BQ
    refunding_net_debt_service: float = 0.0  # BR
    total_net_debt_service: float = 0.0   # BS
    coverage: float = 0.0                 # BT
    senior_surplus: float = 0.0           # BU
    funds_on_hand: float = 0.0            # BW
    annual_surplus: float = 0.0           # BX
    surplus_release: float = 0.0          # BY
    surplus_fund_balance: float = 0.0     # BZ
    cumulative_surplus: float = 0.0       # CA
    debt_to_taxable_value: Optional[float] = None    # CB
    debt_to_market_value: Optional[float] = None     # CC

    # Subordinate waterfall (cols CH–CY)
    sub_available: float = 0.0            # CH
    prior_year_surplus_applied: float = 0.0  # CI
    sub_total_available: float = 0.0      # CJ
    sub_interest_due: float = 0.0         # CK
    sub_interest_paid: float = 0.0        # CL
    sub_accrued_added: float = 0.0        # CM
    sub_accrued_paid: float = 0.0         # CN
    sub_accrued_balance: float = 0.0      # CO
    sub_principal_issued: float = 0.0     # CQ
    sub_principal_paid: float = 0.0       # CR
    sub_principal_balance: float = 0.0    # CS
    sub_total_payments: float = 0.0       # CU
    sub_surplus_cashflow: float = 0.0     # CW
    sub_cumulative_surplus: float = 0.0   # CY


@dataclass
class DebtServiceRow:
    """One semi-annual period on a debt-service tab."""
    payment_date: date                    # A
    rate: float = 0.0                     # B
    yld: float = 0.0                      # C
    price: float = 0.0                    # D  (as a decimal of par)
    premium_oid: float = 0.0              # F
    principal: float = 0.0                # G
    interest: float = 0.0                 # H
    total: float = 0.0                    # I
    annual_gross: float = 0.0             # J
    capitalized_interest: float = 0.0     # K
    surplus_release: float = 0.0          # L
    interest_earnings: float = 0.0        # M
    net_total: float = 0.0                # N
    annual_net: float = 0.0               # O
    bond_value: float = 0.0               # P
    revenue: float = 0.0                  # T
    actual_coverage: float = 0.0          # U
    coverage_target: float = 0.0          # V
    bond_years: float = 0.0               # X


@dataclass
class SubDebtServiceRow(DebtServiceRow):
    """Sub-lien tabs carry the accrued-interest waterfall alongside the coupon."""
    annual_interest: float = 0.0          # I
    interest_paid: float = 0.0            # J
    accrued_added: float = 0.0            # K
    accrued_paid: float = 0.0             # L
    accrued_balance: float = 0.0          # M
    aggregate_debt_service: float = 0.0   # N


@dataclass
class CapiRow:
    draw_date: date
    construction_amount: float = 0.0
    capitalized_interest: float = 0.0
    combined_draws: float = 0.0
    beginning_balance: float = 0.0
    days: int = 0
    periodic_rate: float = 0.0
    interest: float = 0.0
    draws: float = 0.0
    ending_balance: float = 0.0


@dataclass
class BondStatistics:
    par: float = 0.0
    premium: float = 0.0
    bond_years: float = 0.0
    average_life: float = 0.0
    arbitrage_tic: float = 0.0
    all_in_tic: float = 0.0
    max_annual_debt_service: float = 0.0
    total_debt_service: float = 0.0
    first_maturity: Optional[date] = None
    final_maturity: Optional[date] = None


@dataclass
class Results:
    cfg: ModelConfig
    dev: DevelopmentProjections
    summary: list[SummaryRow] = field(default_factory=list)
    senior: list[DebtServiceRow] = field(default_factory=list)
    sub_sa: list[SubDebtServiceRow] = field(default_factory=list)
    sub_annual: list[SubDebtServiceRow] = field(default_factory=list)
    refunding: list[DebtServiceRow] = field(default_factory=list)
    capi: list[CapiRow] = field(default_factory=list)
    senior_stats: BondStatistics = field(default_factory=BondStatistics)
    sub_stats: BondStatistics = field(default_factory=BondStatistics)
    refunding_stats: BondStatistics = field(default_factory=BondStatistics)
    surplus_fund_deposit: float = 0.0
    surplus_fund_target: float = 0.0
    capitalized_interest_deposit: float = 0.0
    iterations: int = 0
    converged: bool = False
    warnings: list[str] = field(default_factory=list)

    # ── Sources and Uses ──────────────────────────────────────────────────────

    @property
    def senior_par(self) -> float:
        return self.senior_stats.par

    @property
    def sub_par(self) -> float:
        return self.sub_stats.par

    @property
    def total_par(self) -> float:
        return self.senior_par + self.sub_par

    @property
    def senior_uwd(self) -> float:
        return self.cfg.uwd_senior * self.senior_par

    @property
    def sub_uwd(self) -> float:
        return self.cfg.uwd_sub * self.sub_par

    @property
    def senior_reimbursement(self) -> float:
        return (self.senior_par + self.senior_stats.premium
                - self.surplus_fund_deposit
                - self.capitalized_interest_deposit
                - self.senior_uwd - self.cfg.coi)

    @property
    def sub_reimbursement(self) -> float:
        return self.sub_par - self.sub_uwd

    @property
    def total_reimbursement(self) -> float:
        return self.senior_reimbursement + self.sub_reimbursement

    @property
    def reimbursement_per_lot(self) -> float:
        units = self.dev.total_units()
        return safe_div(self.total_reimbursement, units)

    @property
    def repayment_ratio(self) -> float:
        total_ds = (self.senior_stats.total_debt_service
                    + self.sub_stats.total_debt_service)
        return safe_div(total_ds, self.total_reimbursement)


# ── The model ─────────────────────────────────────────────────────────────────

class Model:
    """Runs a Utah PID financing analysis end to end."""

    def __init__(self, cfg: ModelConfig, dev: DevelopmentProjections):
        self.cfg = cfg
        self.dev = dev

    # -- public API ------------------------------------------------------------

    def run(self, max_iterations: int = 40, tolerance: float = 1.0) -> Results:
        cfg = self.cfg
        res = Results(cfg=cfg, dev=self.dev)
        res.warnings = cfg.validate()

        summary = self._build_taxable_value()
        self._apply_revenue(summary)

        surplus_deposit = 0.0
        surplus_target = 0.0
        senior: list[DebtServiceRow] = []

        for i in range(max_iterations):
            senior = self._size_senior(summary, surplus_deposit)
            new_deposit = self._surplus_fund_deposit(senior, summary)
            new_target = self._surplus_fund_target(senior, new_deposit)
            converged = (abs(new_deposit - surplus_deposit) < tolerance
                         and abs(new_target - surplus_target) < tolerance)
            surplus_deposit, surplus_target = new_deposit, new_target
            res.iterations = i + 1
            if converged and i > 0:
                res.converged = True
                break

        senior = self._size_senior(summary, surplus_deposit)
        res.surplus_fund_deposit = surplus_deposit
        res.surplus_fund_target = surplus_target
        res.senior = senior

        # Push senior debt service onto the Summary, then run the sub waterfall.
        self._apply_senior_debt_service(summary, senior, surplus_deposit, surplus_target)
        self._surplus_target_cache = surplus_target
        sub_par = self._size_sub_par(summary)
        sub_sa, sub_annual = self._build_sub_schedules(summary, sub_par)
        res.sub_sa, res.sub_annual = sub_sa, sub_annual

        res.capi = self._build_capi(senior)
        res.capitalized_interest_deposit = res.capi[0].beginning_balance if res.capi else 0.0

        res.senior_stats = self._statistics(senior, cfg.senior_interest_rate,
                                            res.capitalized_interest_deposit,
                                            cfg.uwd_senior, cfg.coi)
        res.sub_stats = self._sub_statistics(sub_annual)
        res.refunding = self._build_refunding(senior)
        res.refunding_stats = self._statistics(res.refunding,
                                               cfg.senior_refunding_interest_rate,
                                               0.0, cfg.uwd_senior_refunding,
                                               cfg.coi_refunding)

        self._finalise_summary(summary, res)
        res.summary = summary
        return res

    # -- taxable value ---------------------------------------------------------

    def _summary_years(self) -> list[int]:
        """
        The Summary axis runs past the development table: absorption finishes
        long before the bonds do, and the later rows still carry reassessment
        growth and debt service.  It ends one year after the last maturity.
        """
        cfg, dev = self.cfg, self.dev
        start = dev.start_year
        last_maturity = max(
            edate(date(cfg.first_int.year, cfg.prin_maturity, 1),
                  12 * cfg.final_mat_yrs).year,
            edate(date(cfg.first_int_sub.year, cfg.prin_maturity, 1),
                  12 * cfg.final_mat_sub_yrs).year,
        )
        end = max(dev.years()[-1], last_maturity)
        return list(range(start, end + 1))

    def _build_taxable_value(self) -> list[SummaryRow]:
        cfg, dev = self.cfg, self.dev
        years = self._summary_years()
        lag = cfg.value_lag_years

        def pad(series: list[float]) -> list[float]:
            return series + [0.0] * (len(years) - len(series))

        lot_value_created = pad([sum(row) for row in dev.lot_value(cfg)])
        av_created_lagged = pad(dev.total_av_creation_lagged_by_year(cfg))
        lot_units = pad(dev.total_lot_units_by_year(cfg))
        closings = pad(dev.total_closings_by_year(cfg))
        comm_value = pad(dev.total_comm_value_by_year(cfg))
        comm_sf = pad(dev.total_comm_sf_by_year(cfg))

        annual = cfg.reassess_frequency.strip().lower() == "annual"
        rows: list[SummaryRow] = []

        for r, y in enumerate(years):
            prev = rows[r - 1] if r else None
            row = SummaryRow(
                index=r,
                assessment_date=date(y, cfg.prin_maturity, 1),
                tax_revenue_date=date(y + 1, cfg.prin_maturity, 1),
            )

            # Existing homes already on the roll.
            prev_existing = prev.existing_market_value if prev else dev.existing_home_market_value
            if prev is None:
                row.existing_reassessment = 0.0
            elif annual:
                row.existing_reassessment = prev_existing * cfg.reassess_rate_resid
            else:
                row.existing_reassessment = (prev_existing * cfg.reassess_rate_resid
                                             if prev.existing_reassessment <= 0 else 0.0)
            row.existing_market_value = prev_existing + row.existing_reassessment
            row.existing_taxable_value = row.existing_market_value * cfg.resid_taxable_ratio

            # Developer lot inventory: value created `lag` years earlier, which
            # is the inventory still standing when the assessor calls on 1 Jan.
            row.lot_units = lot_units[r]
            row.lot_market_value = lot_value_created[r - lag] if r - lag >= 0 else 0.0
            row.lot_taxable_value = row.lot_market_value * cfg.developed_lot_value

            # Closed homes.
            row.residential_units = closings[r]
            prev_new = prev.new_home_market_value if prev else 0.0
            if prev is None:
                row.new_home_reassessment = 0.0
            elif annual:
                row.new_home_reassessment = prev_new * cfg.reassess_rate_resid_sub
            else:
                row.new_home_reassessment = (prev_new * cfg.reassess_rate_resid_sub
                                             if prev.new_home_reassessment <= 0 else 0.0)
            row.new_home_market_value = (av_created_lagged[r] + prev_new
                                         + row.new_home_reassessment)
            row.new_home_taxable_value = row.new_home_market_value * cfg.resid_taxable_ratio

            row.prior_roll_value = dev.existing_lot_value if r == 0 else 0.0
            row.residential_taxable_value = (row.existing_taxable_value
                                             + row.lot_taxable_value
                                             + row.new_home_taxable_value
                                             + row.prior_roll_value)

            # Commercial.
            row.comm_sf_delivered = comm_sf[r]
            prev_comm = prev.comm_market_value_platted if prev else 0.0
            row.comm_reassessment_platted = (prev_comm * cfg.reassess_rate_comm
                                             if (annual or not prev
                                                 or prev.comm_reassessment_platted <= 0)
                                             else 0.0)
            row.comm_market_value_platted = (comm_value[r] + prev_comm
                                             + row.comm_reassessment_platted)
            lag_c = cfg.commercial_lag_years
            row.comm_taxable_value_platted = (
                (rows[r - lag_c].comm_market_value_platted if r - lag_c >= 0 else 0.0)
                * cfg.platted_comm_lot_value)
            row.comm_sf_sold = comm_sf[r]
            row.comm_market_value_sold = row.comm_market_value_platted
            row.comm_taxable_value_sold = (
                (rows[r - lag_c].comm_market_value_sold if r - lag_c >= 0 else 0.0)
                * cfg.developed_lot_value)
            row.comm_taxable_value = (row.comm_taxable_value_platted
                                      + row.comm_taxable_value_sold)

            row.centrally_assessed_market = cfg.centrally_assessed_value
            row.centrally_assessed_taxable = (row.centrally_assessed_market
                                              * cfg.centrally_assessed_ratio)
            row.state_assessed_taxable = cfg.state_assessed

            rows.append(row)
        return rows

    # -- revenue ---------------------------------------------------------------

    def _apply_revenue(self, rows: list[SummaryRow]) -> None:
        cfg = self.cfg
        for r, row in enumerate(rows):
            prev = rows[r - 1] if r else None

            # Senior pledged base.  Utah pledges the whole residential roll to
            # both liens; the Colorado "new value add" toggles are retained so a
            # senior lien can be limited to existing value if a rating requires.
            senior_resid = row.residential_taxable_value
            if cfg.resid_new_value_add != "Yes":
                senior_resid -= row.new_home_taxable_value
            row.senior_taxable_value = (
                senior_resid
                + (row.centrally_assessed_taxable
                   if cfg.centrally_assessed_senior == "Yes" else 0.0)
                + row.state_assessed_taxable)

            row.senior_mill_levy = cfg.mill_levy_ds_target
            row.senior_levy_collections = (row.senior_taxable_value / 1000.0
                                           * row.senior_mill_levy
                                           * cfg.tax_collect_mill_prc)
            fee_rate = (cfg.uniform_fee_prc / 2.0
                        if row.senior_taxable_value < cfg.uniform_fee_av_threshold
                        else cfg.uniform_fee_prc)
            row.senior_uniform_fee = row.senior_levy_collections * fee_rate

            row.senior_comm_taxable_value = (row.comm_taxable_value
                                             if cfg.comm_new_value_add == "Yes" else 0.0)
            row.senior_comm_mill_levy = cfg.mill_levy_comm
            row.senior_comm_collections = (row.senior_comm_taxable_value / 1000.0
                                           * row.senior_comm_mill_levy
                                           * cfg.tax_collect_mill_prc)
            row.senior_comm_uniform_fee = row.senior_comm_collections * cfg.uniform_fee_prc
            row.pif_revenue = row.taxable_sales * cfg.inflation_rate
            row.system_development_fee = cfg.system_development_fee * row.residential_units

            row.senior_total_revenue = (row.senior_levy_collections
                                        + row.senior_uniform_fee
                                        + row.senior_comm_collections
                                        + row.senior_comm_uniform_fee
                                        + row.pif_revenue
                                        + row.system_development_fee)
            row.county_fee = -(row.senior_levy_collections
                               + row.senior_comm_collections) * cfg.county_treasurer_fee

            charging = row.assessment_date.year >= (cfg.district_cost_start_year or 0)
            row.trustee_fee = (-cfg.trustee_fee
                               if charging and row.senior_levy_collections else 0.0)
            if not charging:
                row.admin_costs = 0.0
            elif cfg.admin_cost_av_limit and row.senior_taxable_value > cfg.admin_cost_av_limit:
                row.admin_costs = 0.0
            elif prev is None or prev.admin_costs == 0.0:
                row.admin_costs = -cfg.admin_cost_base
            else:
                row.admin_costs = prev.admin_costs * (1 + cfg.admin_cost_growth)
            row.senior_net_revenue = (row.senior_total_revenue + row.county_fee
                                      + row.trustee_fee + row.admin_costs)

            # Subordinate pledged base.
            sub_resid = row.residential_taxable_value
            if cfg.resid_new_value_add_sub != "Yes":
                sub_resid -= row.new_home_taxable_value
            row.sub_taxable_value = (
                sub_resid
                + (row.centrally_assessed_taxable
                   if cfg.centrally_assessed_sub == "Yes" else 0.0)
                + row.state_assessed_taxable)
            row.sub_levy_collections = (row.sub_taxable_value / 1000.0
                                        * cfg.mill_levy_ds_target
                                        * cfg.tax_collect_mill_prc)
            row.sub_uniform_fee = row.sub_levy_collections * fee_rate
            sub_comm_tv = (row.comm_taxable_value
                           if cfg.comm_new_value_add_sub == "Yes" else 0.0)
            row.sub_comm_collections = (sub_comm_tv / 1000.0 * cfg.mill_levy_comm
                                        * cfg.tax_collect_mill_prc)
            row.sub_comm_uniform_fee = row.sub_comm_collections * cfg.uniform_fee_prc
            row.sub_total_revenue = (row.sub_levy_collections + row.sub_uniform_fee
                                     + row.sub_comm_collections
                                     + row.sub_comm_uniform_fee
                                     + row.pif_revenue + row.system_development_fee)
            row.sub_trustee_fee = (-cfg.trustee_fee_sub
                                   if charging and row.sub_levy_collections else 0.0)
            row.sub_net_revenue = row.sub_total_revenue + row.sub_trustee_fee

    # -- senior sizing ---------------------------------------------------------

    def _senior_dates(self) -> list[date]:
        cfg = self.cfg
        first = cfg.first_int
        anchor = max(first, date(first.year, cfg.prin_maturity,
                                 cfg.prin_maturity_day_senior))
        last = edate(anchor, 6 * 2 * cfg.final_mat_yrs)
        dates, d = [], first
        while d <= last:
            dates.append(d)
            d = edate(d, 6)
        return dates

    def _revenue_by_payment_date(self, rows: list[SummaryRow],
                                 attr: str) -> dict[tuple[int, int], float]:
        return {(r.tax_revenue_date.year, r.tax_revenue_date.month): getattr(r, attr)
                for r in rows}

    def _size_senior(self, rows: list[SummaryRow],
                     surplus_deposit: float) -> list[DebtServiceRow]:
        """
        Solve principal backwards from the final maturity.

        For each principal date the workbook takes the year's net revenue,
        divides by the coverage requirement, strips out the two coupon payments
        that the *already-sized* later maturities generate, and turns what is
        left into principal at $5,000 denominations:

            G = INT( ((Rev / DSC - 2·H_next) / (1 + rate) + release / (1 + rate))
                     / 5000 ) · 5000
        """
        cfg = self.cfg
        rate = cfg.senior_interest_rate
        dsc = cfg.dsc_senior_lien_bonds
        revenue = self._revenue_by_payment_date(rows, "senior_net_revenue")
        dates = self._senior_dates()
        first_prin = self._first_principal_date(rows)
        release_date = cfg.surplus_fund_release_date
        release_on = cfg.surplus_release_sizing == "Yes"

        ds = [DebtServiceRow(payment_date=d) for d in dates]
        h_next = 0.0                              # interest of all later maturities

        for i in range(len(ds) - 1, -1, -1):
            row = ds[i]
            d = row.payment_date
            is_principal_date = (d.month == cfg.prin_maturity and d >= first_prin)
            row.rate = rate if d.month == cfg.prin_maturity else 0.0

            if is_principal_date:
                rev = revenue.get((d.year, d.month), 0.0)
                earnings = surplus_deposit * cfg.interest_earn_rate
                release = surplus_deposit if (release_on and d == release_date) else 0.0
                raw = ((rev + earnings) / dsc - 2.0 * h_next) / (1.0 + rate)
                raw += release / (1.0 + rate)
                row.principal = max(0.0, round_down_to(raw, PRINCIPAL_INCREMENT))
                row.interest_earnings = earnings
                row.surplus_release = release
            else:
                row.principal = 0.0

            row.interest = h_next + row.principal * row.rate / 2.0
            h_next = row.interest

        self._complete_debt_service(ds, rate, cfg.par_call_first,
                                    cfg.premium_call_first, cfg.premium_call_price,
                                    cfg.capi_end_date if cfg.capi == "Yes" else None,
                                    cfg.dsc_senior_lien_bonds, revenue)
        return ds

    def _first_principal_date(self, rows: list[SummaryRow]) -> date:
        """
        FIRST_PRIN: one year before the last year in which homes close, carried
        onto the principal payment month.  Bonds do not amortise while the tax
        base is still being built.
        """
        cfg = self.cfg
        closing_years = [r.assessment_date.year for r in rows if r.residential_units > 0]
        if not closing_years:
            return edate(cfg.first_int, 12)
        target = max(closing_years) - 1
        candidate = date(target, cfg.prin_maturity, cfg.prin_maturity_day_senior)
        earliest = date(cfg.first_int.year, cfg.prin_maturity,
                        cfg.prin_maturity_day_senior)
        if earliest < cfg.first_int:
            earliest = edate(earliest, 12)
        return max(candidate, earliest)

    def _complete_debt_service(self, ds: list[DebtServiceRow], rate: float,
                               par_call: date, premium_call: date,
                               premium_price: float, capi_end: Optional[date],
                               dsc: float,
                               revenue: dict[tuple[int, int], float]) -> None:
        """Fill price/premium, annual roll-ups, CAPI draws and coverage."""
        cfg = self.cfg
        total_par = sum(r.principal for r in ds)
        running = 0.0
        prev_date = cfg.delivery

        for i, row in enumerate(ds):
            d = row.payment_date
            if row.rate and row.principal:
                p = price(cfg.delivery, d, row.rate, row.rate, 100.0, 2)
                if d > par_call:
                    p = min(p, price(cfg.delivery, par_call, row.rate, row.rate, 100.0, 2))
                if d > premium_call:
                    p = min(p, price(cfg.delivery, premium_call, row.rate,
                                     row.rate, premium_price, 2))
                row.yld = row.rate
                row.price = truncate(p, 3) / 100.0
                row.premium_oid = row.price * row.principal - row.principal
                row.bond_years = row.price * row.principal * days360(cfg.delivery, d) / 360.0
            row.total = row.principal + row.interest

            if capi_end is not None:
                span = days360(max(cfg.delivery, prev_date), d)
                covered = days360(max(cfg.delivery, prev_date), min(d, capi_end))
                row.capitalized_interest = max(0.0, covered / span if span else 0.0) * row.interest
            prev_date = d

            row.net_total = row.total - row.capitalized_interest - row.interest_earnings
            running += row.principal
            row.bond_value = total_par - running

            row.revenue = revenue.get((d.year, d.month), 0.0)
            row.coverage_target = dsc if row.rate else 0.0

        # Annual roll-ups: a payment in the interest-only month carries no annual
        # total; the principal month sums itself and the preceding coupon.
        for i, row in enumerate(ds):
            if row.payment_date.month == cfg.int_maturity:
                row.annual_gross = 0.0
                row.annual_net = 0.0
            else:
                prev = ds[i - 1] if i else None
                row.annual_gross = row.total + (prev.total if prev else 0.0)
                row.annual_net = (row.net_total + (prev.net_total if prev else 0.0)
                                  - row.surplus_release)
            row.actual_coverage = safe_div(row.revenue, row.annual_net)

    def _surplus_fund_deposit(self, ds: list[DebtServiceRow],
                              rows: list[SummaryRow]) -> float:
        """
        SURPLUS_FUND_DEPOSIT: the least of 10% of par, 125% of average annual
        debt service, and maximum annual debt service before the release date.
        """
        cfg = self.cfg
        if cfg.surplus_on_off != "Yes":
            return 0.0
        par = sum(r.principal for r in ds)
        if par <= 0:
            return 0.0
        annuals = [r.annual_gross for r in ds if r.annual_gross]
        avg = sum(annuals) / len(annuals) if annuals else 0.0
        # Utah convention includes the final maturity in the MADS test; the
        # Colorado template stops one period short of the release date.
        mads = max((r.annual_gross for r in ds
                    if r.payment_date <= cfg.surplus_fund_release_date), default=0.0)
        return min(par * 0.10, avg * 1.25, mads)

    def _surplus_fund_target(self, ds: list[DebtServiceRow], deposit: float) -> float:
        cfg = self.cfg
        if cfg.surplus_on_off != "Yes":
            return 0.0
        mads = max((r.annual_gross for r in ds
                    if r.payment_date < cfg.surplus_fund_release_date), default=0.0)
        return mads * cfg.surplus_target_multiple

    # -- Summary: debt service, surplus, sub waterfall --------------------------

    def _apply_senior_debt_service(self, rows: list[SummaryRow],
                                   ds: list[DebtServiceRow],
                                   deposit: float, target: float) -> None:
        cfg = self.cfg
        annual_net = {(r.payment_date.year, r.payment_date.month): r.annual_net
                      for r in ds if r.payment_date.month == cfg.prin_maturity}
        outstanding = {(r.payment_date.year, r.payment_date.month): r.bond_value
                       for r in ds}

        prev_balance = 0.0
        cumulative = 0.0
        for r, row in enumerate(rows):
            key = (row.tax_revenue_date.year, row.tax_revenue_date.month)
            row.senior_net_debt_service = annual_net.get(key, 0.0)
            row.total_net_debt_service = (row.senior_net_debt_service
                                          + row.refunding_net_debt_service)
            row.coverage = safe_div(row.senior_net_revenue, row.total_net_debt_service)
            row.senior_surplus = row.senior_net_revenue - row.senior_net_debt_service
            row.annual_surplus = (row.senior_net_revenue - row.total_net_debt_service
                                  - row.funds_on_hand)

            # Surplus fund fills to target, then everything above it is released.
            if row.tax_revenue_date >= cfg.surplus_fund_release_date:
                balance = 0.0
            else:
                balance = min(target, max(0.0, prev_balance + row.senior_surplus))
            row.surplus_fund_balance = balance
            if balance >= target or row.tax_revenue_date >= cfg.surplus_fund_release_date:
                row.surplus_release = row.senior_surplus - (balance - prev_balance)
            else:
                row.surplus_release = 0.0
            prev_balance = balance

            cumulative += row.annual_surplus
            row.cumulative_surplus = cumulative

            bv = outstanding.get(key)
            nxt = rows[r + 1] if r + 1 < len(rows) else None
            if bv is not None and nxt:
                row.debt_to_taxable_value = safe_div(bv, nxt.senior_taxable_value, 0.0) or None
                row.debt_to_market_value = safe_div(bv, nxt.new_home_market_value, 0.0) or None

    def _sub_waterfall(self, rows: list[SummaryRow], sub_par: float,
                       write: bool = False) -> tuple[float, float, list[float]]:
        """
        Run the subordinate cashflow waterfall for a candidate par amount.

        Returns (ending accrued balance, ending principal balance, principal
        paid by year).  A par amount is "fully supported" when both ending
        balances are zero at final maturity.
        """
        cfg = self.cfg
        rate = cfg.sub_interest_rate
        accrued = 0.0
        principal_balance = sub_par
        principal_paid: list[float] = []
        final = cfg.end_bal_accrued_sub_bonds

        for r, row in enumerate(rows):
            # Everything the senior lien releases, less the subordinate
            # trustee's fee, plus any pledged value the senior lien does not
            # reach (Colorado adds back the senior's O&M carveout here; Utah
            # does not, because district administration is a real cost that the
            # subordinate cashflow cannot spend twice).
            available = (row.surplus_release + row.sub_trustee_fee
                         + (row.sub_total_revenue - row.senior_total_revenue))
            if row.tax_revenue_date == cfg.surplus_fund_release_date:
                available += self._surplus_target_cache
            total_available = available / cfg.dsc_sub_lien_bonds if cfg.dsc_sub_lien_bonds else 0.0

            past_final = row.assessment_date > final
            interest_due = 0.0 if past_final else principal_balance * rate
            interest_paid = max(0.0, min(interest_due, available))
            accrued_added = 0.0 if past_final else interest_due + accrued * rate - interest_paid

            # Principal is paid only after current and accrued interest.
            headroom = total_available - interest_due
            paid_principal = 0.0
            if principal_balance > 0 and headroom > 0 and not past_final:
                paid_principal = min(principal_balance,
                                     round_down_to(headroom, SUB_PRINCIPAL_INCREMENT))
                if paid_principal < cfg.sub_sizing_threshold:
                    paid_principal = 0.0
            headroom -= paid_principal
            accrued_paid = 0.0
            if not past_final and headroom > 0:
                accrued_paid = max(0.0, min(headroom, accrued + accrued_added))
            accrued = accrued + accrued_added - accrued_paid
            principal_balance -= paid_principal
            principal_paid.append(paid_principal)

            if write:
                row.sub_available = available
                row.sub_total_available = total_available
                row.sub_interest_due = interest_due
                row.sub_interest_paid = interest_paid
                row.sub_accrued_added = accrued_added
                row.sub_accrued_paid = accrued_paid
                row.sub_accrued_balance = accrued
                row.sub_principal_issued = (sub_par if row.assessment_date.year
                                            == cfg.delivery.year else 0.0)
                row.sub_principal_paid = paid_principal
                row.sub_principal_balance = principal_balance
                row.sub_total_payments = interest_paid + accrued_paid + paid_principal
                row.sub_surplus_cashflow = total_available - row.sub_total_payments
                prev = rows[r - 1].sub_cumulative_surplus if r else 0.0
                row.sub_cumulative_surplus = prev  # released each year

        return accrued, principal_balance, principal_paid

    _surplus_target_cache: float = 0.0

    def _size_sub_par(self, rows: list[SummaryRow]) -> float:
        """
        Largest sub-lien par (in $1,000s) that the residual cashflow retires in
        full — principal *and* accrued interest — by the final maturity.
        """
        cfg = self.cfg
        if cfg.sub_par_override is not None:
            return cfg.sub_par_override

        def supported(par: float) -> bool:
            accrued, balance, _ = self._sub_waterfall(rows, par)
            return balance <= 1e-6 and accrued <= 1e-6

        # Bracket: residual cash with no sub bonds at all is the upper bound on
        # what any sub par can service.
        residual = sum(max(0.0, r.surplus_release + r.sub_trustee_fee
                           + (r.sub_total_revenue - r.senior_total_revenue))
                       for r in rows)
        if residual <= cfg.sub_sizing_threshold:
            return 0.0

        lo, hi = 0.0, round_down_to(residual, SUB_PRINCIPAL_INCREMENT)
        if supported(hi):
            return hi
        for _ in range(60):
            mid = round_down_to((lo + hi) / 2.0, SUB_PRINCIPAL_INCREMENT)
            if mid <= lo or mid >= hi:
                break
            if supported(mid):
                lo = mid
            else:
                hi = mid
        return lo if lo >= cfg.sub_sizing_threshold else 0.0

    # -- sub debt-service tabs --------------------------------------------------

    def _build_sub_schedules(self, rows: list[SummaryRow], sub_par: float
                             ) -> tuple[list[SubDebtServiceRow], list[SubDebtServiceRow]]:
        cfg = self.cfg
        _, _, principal_paid = self._sub_waterfall(rows, sub_par, write=True)
        by_year = {rows[i].tax_revenue_date.year: p
                   for i, p in enumerate(principal_paid)}
        paid_interest = {r.tax_revenue_date.year: r.sub_interest_paid for r in rows}
        accrued_add = {r.tax_revenue_date.year: r.sub_accrued_added for r in rows}
        accrued_pay = {r.tax_revenue_date.year: r.sub_accrued_paid for r in rows}
        accrued_bal = {r.tax_revenue_date.year: r.sub_accrued_balance for r in rows}

        first = cfg.first_int_sub
        anchor = max(first, date(first.year, cfg.prin_maturity, cfg.prin_maturity_day_sub))
        last = edate(anchor, 6 * 2 * cfg.final_mat_sub_yrs)

        sa: list[SubDebtServiceRow] = []
        d = first
        while d <= last:
            sa.append(SubDebtServiceRow(payment_date=d))
            d = edate(d, 6)

        rate = cfg.sub_interest_rate
        balance = sub_par
        for row in sa:
            d = row.payment_date
            row.rate = rate if d.month == cfg.prin_maturity else 0.0
            row.principal = by_year.get(d.year, 0.0) if d.month == cfg.prin_maturity else 0.0
            row.interest = balance * rate / 2.0
            row.bond_value = balance
            balance -= row.principal
            row.interest_paid = paid_interest.get(d.year, 0.0) if row.rate else 0.0
            row.accrued_added = accrued_add.get(d.year, 0.0) if row.rate else 0.0
            row.accrued_paid = accrued_pay.get(d.year, 0.0) if row.rate else 0.0
            row.accrued_balance = accrued_bal.get(d.year, 0.0) if row.rate else 0.0
            row.aggregate_debt_service = (row.principal + row.interest_paid
                                          + row.accrued_paid)
            row.total = row.principal + row.interest + row.accrued_paid
            row.net_total = row.total - row.capitalized_interest
        for i, row in enumerate(sa):
            if row.payment_date.month == cfg.int_maturity:
                row.annual_interest = 0.0
                row.annual_gross = 0.0
                row.annual_net = 0.0
            else:
                prev = sa[i - 1] if i else None
                row.annual_interest = row.interest + (prev.interest if prev else 0.0)
                row.annual_gross = row.total + (prev.total if prev else 0.0)
                row.annual_net = row.net_total + (prev.net_total if prev else 0.0)

        # The annual tab is the same schedule collapsed to the principal date.
        annual = [r for r in sa if r.payment_date.month == cfg.prin_maturity]
        annual_rows: list[SubDebtServiceRow] = []
        for r in annual:
            copy = SubDebtServiceRow(**{k: v for k, v in r.__dict__.items()})
            copy.interest = copy.annual_interest
            copy.total = copy.principal + copy.interest_paid + copy.accrued_paid
            annual_rows.append(copy)
        return sa, annual_rows

    # -- capitalized interest fund ---------------------------------------------

    def _build_capi(self, ds: list[DebtServiceRow]) -> list[CapiRow]:
        cfg = self.cfg
        if cfg.capi != "Yes":
            return []
        draws = {(r.payment_date.year, r.payment_date.month): r.capitalized_interest
                 for r in ds if r.capitalized_interest}

        dates = [cfg.delivery]
        d = cfg.capi_first_draw
        while d <= cfg.capi_end_date:
            dates.append(d)
            d = edate(d, 1)

        monthly = (1 + cfg.interest_earn_rate / 2) ** (1 / 6) - 1 if cfg.interest_earn_rate else 0.0
        rows = [CapiRow(draw_date=dt) for dt in dates]
        for row in rows[1:]:
            row.capitalized_interest = draws.get((row.draw_date.year,
                                                  row.draw_date.month), 0.0)
            row.combined_draws = row.construction_amount + row.capitalized_interest

        def roll(opening: float) -> float:
            """Run the fund forward from `opening` and return the ending balance."""
            rows[0].beginning_balance = opening
            rows[0].draws = rows[0].combined_draws
            rows[0].ending_balance = opening - rows[0].draws
            for i in range(1, len(rows)):
                row, prev = rows[i], rows[i - 1]
                row.beginning_balance = prev.ending_balance
                row.days = days360(prev.draw_date, row.draw_date)
                if cfg.interest_earn_rate and row.days:
                    row.periodic_rate = (((1 + cfg.interest_earn_rate / 2)
                                          ** (2 / (360 / row.days)) - 1) * 360 / row.days)
                    row.interest = row.beginning_balance * row.periodic_rate * row.days / 360
                else:
                    row.periodic_rate = row.interest = 0.0
                row.draws = row.combined_draws
                row.ending_balance = row.beginning_balance + row.interest - row.draws
            return rows[-1].ending_balance

        # Deposit the present value of the draw stream, then solve the residual
        # away: the ending balance is affine in the opening deposit, because
        # interest accrues on a 30/360 day count rather than the flat monthly
        # rate the discounting uses.  Two passes give the exact deposit.
        flows = [r.combined_draws for r in rows[1:]]
        guess = (npv(monthly, flows) * (1 + monthly) ** (1 - cfg.frac)
                 if monthly else sum(flows))
        end_a = roll(guess)
        if abs(end_a) > 1e-9:
            end_b = roll(guess + 1.0)
            slope = end_b - end_a
            if slope:
                roll(guess - end_a / slope)
            else:
                roll(guess)
        return rows

    # -- refunding --------------------------------------------------------------

    def _build_refunding(self, senior: list[DebtServiceRow]) -> list[DebtServiceRow]:
        """
        Refunding schedule.  With REFUND_FINANCING = "No" the tab is laid out and
        dated but carries no principal, exactly as the source workbook does.
        """
        cfg = self.cfg
        first = cfg.first_int_refunding
        anchor = max(first, date(first.year, cfg.prin_maturity,
                                 cfg.prin_maturity_day_senior))
        last = edate(anchor, 6 * 2 * cfg.final_mat_yrs_refunding)
        rows, d = [], first
        while d <= last:
            rows.append(DebtServiceRow(payment_date=d))
            d = edate(d, 6)
        if cfg.refund_financing != "Yes":
            return rows

        refunded = sum(r.principal for r in senior
                       if r.payment_date > cfg.par_call_first)
        rate = cfg.senior_refunding_interest_rate
        n = len([r for r in rows if r.payment_date.month == cfg.prin_maturity])
        level = refunded / n if n else 0.0
        for r in rows:
            if r.payment_date.month == cfg.prin_maturity:
                r.rate = rate
                r.principal = round_down_to(level, PRINCIPAL_INCREMENT)
        balance = sum(r.principal for r in rows)
        total_par = balance
        for r in rows:
            r.interest = balance * rate / 2.0
            r.total = r.principal + r.interest
            balance -= r.principal
            r.bond_value = balance
            r.net_total = r.total
        for i, r in enumerate(rows):
            if r.payment_date.month == cfg.int_maturity:
                r.annual_gross = r.annual_net = 0.0
            else:
                prev = rows[i - 1] if i else None
                r.annual_gross = r.total + (prev.total if prev else 0.0)
                r.annual_net = r.net_total + (prev.net_total if prev else 0.0)
        return rows

    # -- statistics -------------------------------------------------------------

    def _statistics(self, ds: list[DebtServiceRow], rate: float,
                    capi_deposit: float, uwd_rate: float, coi: float) -> BondStatistics:
        cfg = self.cfg
        stats = BondStatistics()
        stats.par = sum(r.principal for r in ds)
        stats.premium = sum(r.premium_oid for r in ds)
        stats.bond_years = sum(r.bond_years for r in ds)
        proceeds = stats.par + stats.premium
        stats.average_life = safe_div(stats.bond_years, proceeds)
        stats.total_debt_service = sum(r.annual_net for r in ds)
        stats.max_annual_debt_service = max((r.annual_net for r in ds), default=0.0)
        maturities = [r.payment_date for r in ds if r.principal]
        stats.first_maturity = min(maturities) if maturities else None
        stats.final_maturity = max(maturities) if maturities else None

        flows = [r.total for r in ds]
        offset = 0.0                                   # workbook cells AD11 / AI11
        stats.arbitrage_tic = solve_tic(flows, proceeds, offset)
        net = proceeds - (uwd_rate * stats.par) - coi
        stats.all_in_tic = solve_tic(flows, net, offset)
        return stats

    def _sub_statistics(self, annual: list[SubDebtServiceRow]) -> BondStatistics:
        stats = BondStatistics()
        stats.par = sum(r.principal for r in annual)
        stats.total_debt_service = sum(r.aggregate_debt_service for r in annual)
        stats.max_annual_debt_service = max((r.aggregate_debt_service for r in annual),
                                            default=0.0)
        maturities = [r.payment_date for r in annual if r.principal]
        stats.first_maturity = min(maturities) if maturities else None
        stats.final_maturity = max(maturities) if maturities else None
        if stats.par:
            years = sum(r.principal * yearfrac(self.cfg.delivery, r.payment_date)
                        for r in annual)
            stats.bond_years = years
            stats.average_life = years / stats.par
            flows = [r.aggregate_debt_service for r in annual]
            stats.arbitrage_tic = solve_tic(flows, stats.par, 0.0, periods_per_year=1)
            net = stats.par - self.cfg.uwd_sub * stats.par
            stats.all_in_tic = solve_tic(flows, net, 0.0, periods_per_year=1)
        return stats

    def _finalise_summary(self, rows: list[SummaryRow], res: Results) -> None:
        """Attach anything that could only be known once every tab was built."""
        cfg = self.cfg
        for row in rows:
            if row.sub_principal_issued:
                row.sub_principal_issued = res.sub_par
        if res.senior_par <= 0:
            res.warnings.append(
                "No senior par was sized — projected pledged revenue never "
                "covers debt service at the assumed coverage requirement.")

        # Interest-only years are not coverage-tested by the sizing routine (the
        # workbook behaves the same way), so flag any that fall short.
        short = [r.tax_revenue_date.year for r in rows
                 if r.total_net_debt_service > 0 and r.coverage < 1.0]
        if short:
            res.warnings.append(
                "Projected senior coverage falls below 1.00x in "
                + ", ".join(str(y) for y in short)
                + " — the shortfall is met from the surplus fund / debt service "
                  "reserve. Consider lengthening the capitalized interest period.")

        accrued = rows[-1].sub_accrued_balance if rows else 0.0
        if accrued > 1.0:
            res.warnings.append(
                f"Subordinate accrued interest of ${accrued:,.0f} remains "
                "unpaid at final maturity.")


def run_model(cfg: ModelConfig, dev: DevelopmentProjections) -> Results:
    return Model(cfg, dev).run()
