"""
series_c.py — Series C cash-flow bonds.

A third-tier **new-money** cash-flow bond sized against a *separate* assessment:
the same created taxable value, but biennially reassessed at
``SERIES_C_REASSESS_RATE`` (e.g. 6%) instead of the normal residential
retaxable ratio.  The net pledged revenue produced under that separate
assessment, less the senior and subordinate lien debt service, is what the
Series C bonds are sized against — using the **same accreting cash-flow
methodology** as the subordinate lien.

The separate assessment is used ONLY to size these bonds; the senior and
subordinate liens continue to be sized against the normal assessment.
"""

from __future__ import annotations

import dataclasses
from dataclasses import dataclass

from .config import ModelConfig
from .summary import SummaryModel
from .debt_service import BondTranche
from .subordinate import SubordinateLien, SubLienResult


class _SeriesCFund:
    """
    Availability provider for the Series C lien: the cash available each year is
    the Series C net revenue (revenue on the 6%-reassessed base) less the senior
    and subordinate debt service.  Presented to ``SubordinateLien`` in place of a
    surplus fund so the identical accreting-cash-flow amortization is reused.
    """

    def __init__(self, avail_by_year: dict[int, float]):
        self._avail = dict(avail_by_year)
        self._by_year: dict = {}          # sub sizer reads this only for display

    def available_to_sub(self, year: int) -> float:
        return self._avail.get(year, 0.0)


@dataclass
class SeriesCResult:
    par_amount: float
    schedule: SubLienResult               # accreting amortization (reuses sub math)
    available_by_year: dict[int, float]   # cash available to Series C
    rev_by_year: dict[int, float]         # Series C net revenue (6%-reassessed base)
    senior_ds_by_year: dict[int, float]
    sub_ds_by_year: dict[int, float]
    reassess_rate: float
    contribution: float = 0.0             # developer contribution applied to Series C

    # Pass-throughs mirroring SubLienResult so callers can treat it like a sub note.
    @property
    def coverage(self) -> float: return self.schedule.coverage
    @property
    def rate(self) -> float: return self.schedule.rate
    @property
    def first_year(self) -> int: return self.schedule.first_year
    @property
    def final_year(self) -> int: return self.schedule.final_year
    @property
    def rows(self) -> list: return self.schedule.rows
    @property
    def fully_repaid(self) -> bool: return self.schedule.fully_repaid
    @property
    def total_payments(self) -> float: return self.schedule.total_payments
    @property
    def total_interest_paid(self) -> float: return self.schedule.total_interest_paid
    @property
    def total_principal_paid(self) -> float: return self.schedule.total_principal_paid
    @property
    def ending_accrued_interest(self) -> float: return self.schedule.ending_accrued_interest


def series_c_revenue(cfg: ModelConfig, dev, first_year: int, final_year: int) -> dict[int, float]:
    """
    Net pledged revenue under the Series C separate assessment — the created AV
    biennially reassessed at ``cfg.series_c_reassess_rate`` — by year.

    Built by re-running the summary with the residential biennial-reassessment
    rate swapped for the Series C rate (the development program is unchanged, so
    the same ``dev`` is reused).
    """
    cfg_c = dataclasses.replace(
        cfg, reassess_rate=cfg.series_c_reassess_rate)
    sm_c = SummaryModel(cfg_c, dev).build()
    return {y: sm_c.net_senior_revenue(y) for y in range(first_year, final_year + 1)}


def size_series_c(
    cfg: ModelConfig,
    sm: SummaryModel,
    dev,
    senior: BondTranche,
    sub_result: SubLienResult | None,
    first_year: int,
    final_year: int,
) -> SeriesCResult:
    """
    Size the Series C cash-flow bonds: the largest par the residual Series C
    revenue (6%-reassessed net revenue − senior DS − sub DS) fully repays
    (principal + accreted interest) by ``final_year`` — same bisection and
    accreting amortization as the subordinate lien.
    """
    # Series C carries its OWN interest rate, coverage and final maturity — same
    # 12/15 payment dates as the sub lien.  Reuse the subordinate accreting sizer
    # via a config clone that swaps in the Series C coupon/coverage.
    final_year = cfg.series_c_final_year

    rev = series_c_revenue(cfg, dev, first_year, final_year)
    senior_nd = senior.annual_net_ds()
    sub_paid = ({r["year"]: r["total_paid"] for r in sub_result.rows}
                if sub_result is not None else {})

    senior_ds = {y: senior_nd.get(y, 0.0) for y in range(first_year, final_year + 1)}
    sub_ds = {y: sub_paid.get(y, 0.0) for y in range(first_year, final_year + 1)}
    avail = {y: max(0.0, rev.get(y, 0.0) - senior_ds[y] - sub_ds[y])
             for y in range(first_year, final_year + 1)}

    fund = _SeriesCFund(avail)
    cfg_c = dataclasses.replace(
        cfg, sub_interest_rate=cfg.series_c_interest_rate, dsc_sub=cfg.dsc_series_c)
    lien = SubordinateLien(cfg_c, sm)
    par = lien.size_par(senior, first_year, final_year,
                        surplus_fund=fund, dated=cfg.delivery)
    schedule = lien.size(par, senior, first_year, final_year,
                         surplus_fund=fund, dated=cfg.delivery)

    return SeriesCResult(
        par_amount=par, schedule=schedule, available_by_year=avail,
        rev_by_year=rev, senior_ds_by_year=senior_ds, sub_ds_by_year=sub_ds,
        reassess_rate=cfg.series_c_reassess_rate,
    )
