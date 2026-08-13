"""
subordinate.py — Subordinate-lien / cash-flow bonds + senior surplus fund.

This module reproduces the full subordinate-lien waterfall from the workbook's
"Summary" sheet (columns BU–CA for the senior surplus fund and CH–CY for the
subordinate cash flow) together with the "Sub Lien DS" amortization.

Structure modeled
-----------------
1. **Senior surplus / debt-service-reserve fund** (`SurplusFund`)
   * Accumulates the senior residual ``net_senior_revenue - senior_net_DS`` each
     year (Summary ``BU``), capped at the surplus-fund **target** (Summary
     ``BZ``; target = ``surplus_fund_target_factor`` × max senior annual DS).
   * Cash above the target spills to the subordinate lien (Summary ``BY``).
   * The funded debt-service-reserve (and any surplus balance) is **released**
     when the senior lien is retired, providing a surge of cash to the sub lien.

2. **Subordinate cash-flow bonds** (`SubordinateLien`)
   * Carry a **coverage factor** (``DSC_SUB``) applied to available cash
     (Summary ``CJ = (CH + CI) / DSC_SUB``).
   * Accrue interest at the subordinate rate on the unpaid balance; **unpaid
     interest accretes / compounds** (Summary ``CK``→``CO``).
   * Are repaid from available cash interest-first, then principal — so they sit
     and accrete while the senior lien has first call on revenue, and are paid
     down once surplus / reserve-release cash becomes available.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from .config import ModelConfig
from .summary import SummaryModel
from .debt_service import BondTranche


# ──────────────────────────────────────────────────────────────────────────────
# Senior surplus / debt-service-reserve fund
# ──────────────────────────────────────────────────────────────────────────────
@dataclass
class SurplusFundRow:
    year: int
    senior_residual: float       # BU = net senior revenue - senior net DS
    deposit_to_reserve: float    # contribution into the surplus fund this year
    reserve_balance: float       # BZ = funded surplus fund balance (<= target)
    excess_to_sub: float         # BY = cash above target available to sub lien
    reserve_release: float       # DSRF + balance released when senior retires


class SurplusFund:
    """
    Senior operating-surplus / debt-service-reserve fund (Summary BU–CA).

    The fund builds from the senior residual toward ``target``; anything beyond
    the target spills to the subordinate lien.  The funded reserve plus any
    remaining balance is released to the sub lien when the senior lien matures.
    """

    def __init__(self, config: ModelConfig, summary: SummaryModel):
        self.cfg = config
        self.sm = summary
        self.rows: list[SurplusFundRow] = []
        self._by_year: dict[int, SurplusFundRow] = {}

    def build(
        self,
        senior: BondTranche,
        refunding: BondTranche | None,
        first_year: int,
        last_year: int,
    ) -> "SurplusFund":
        cfg = self.cfg
        senior_net = senior.annual_net_ds()
        ref_net = refunding.annual_net_ds() if refunding else {}

        # Reserve target = factor x max annual net DS of the *active* lien
        # (Summary SURPLUS_FUND_TARGET = 0.5 x max senior DS).  The refunding
        # lien runs the same waterfall as the new-money lien: it funds its own
        # reserve target and spills excess to the sub lien.
        target_first = cfg.surplus_fund_target_factor * max(
            list(senior_net.values()) + [0.0])
        target_refunding = (
            cfg.surplus_fund_target_factor * max(list(ref_net.values()) + [0.0])
            if refunding else target_first
        )
        self.target = target_first
        self.target_refunding = target_refunding

        # After the refunding date the original senior bonds are defeased, so the
        # *active* senior debt service switches from the first lien to the
        # refunding lien — don't double-count both.  The reserve is released to
        # the sub lien at the active lien's final maturity (the refunding lien's
        # maturity when refunded).
        refund_year = refunding.delivery.year if refunding else None
        release_year = refunding.final_year if refunding else senior.final_year

        balance = 0.0
        self.rows = []
        self._by_year = {}
        for y in range(first_year, last_year + 1):
            if refund_year is not None and y > refund_year:
                sr_ds = ref_net.get(y, 0.0)
                target = target_refunding
            else:
                sr_ds = senior_net.get(y, 0.0)
                target = target_first
            residual = self.sm.net_senior_revenue(y) - sr_ds

            prev_balance = balance
            # Build the reserve toward target; cash beyond the target spills out.
            target_room = max(0.0, target - prev_balance)
            if residual >= 0:
                deposit = min(residual, target_room)
                excess = residual - deposit
            else:
                deposit = max(residual, -prev_balance)  # draw down (never below 0)
                excess = 0.0
            balance = prev_balance + deposit

            # Release the surplus-fund target down to the sub lien at the senior
            # surplus-fund release date (Summary CH release-date branch).
            release = 0.0
            if y == release_year:
                release = target
                balance = 0.0

            row = SurplusFundRow(
                year=y, senior_residual=residual, deposit_to_reserve=deposit,
                reserve_balance=balance, excess_to_sub=excess, reserve_release=release,
            )
            self.rows.append(row)
            self._by_year[y] = row
        return self

    def available_to_sub(self, year: int) -> float:
        """
        Cash passed down to the subordinate lien in ``year`` (Summary CH).

        Mirrors ``CH = IF(D=release, BY+TARGET+(BO-AX), IF(BY=0, 0, BY+(BO-AX)))``:
        the sub lien receives the incremental subordinate revenue (BO−AX) and any
        surplus above target **only once the senior reserve is full**, plus the
        released reserve target at the senior surplus-fund release date.
        """
        r = self._by_year.get(year)
        if r is None:
            return 0.0
        sub_incremental = max(
            0.0, self.sm.net_sub_revenue(year) - self.sm.net_senior_revenue(year))
        if r.reserve_release > 0:
            return r.excess_to_sub + r.reserve_release + sub_incremental
        if r.excess_to_sub > 0:
            return r.excess_to_sub + sub_incremental
        return 0.0

    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame([{
            "year": r.year,
            "senior_residual": r.senior_residual,
            "deposit_to_reserve": r.deposit_to_reserve,
            "reserve_balance": r.reserve_balance,
            "excess_to_sub": r.excess_to_sub,
            "reserve_release": r.reserve_release,
        } for r in self.rows])


# ──────────────────────────────────────────────────────────────────────────────
# Subordinate cash-flow bonds
# ──────────────────────────────────────────────────────────────────────────────
@dataclass
class SubLienResult:
    par_amount: float
    first_year: int
    final_year: int
    rate: float
    coverage: float
    rows: list[dict] = field(default_factory=list)

    @property
    def total_payments(self) -> float:
        return sum(r["total_paid"] for r in self.rows)

    @property
    def total_interest_paid(self) -> float:
        return sum(r["interest_paid"] for r in self.rows)

    @property
    def total_principal_paid(self) -> float:
        return sum(r["principal_paid"] for r in self.rows)

    @property
    def ending_accrued_interest(self) -> float:
        return self.rows[-1]["accrued_balance"] if self.rows else 0.0

    @property
    def fully_repaid(self) -> bool:
        return self.rows and self.rows[-1]["principal_balance"] <= 1.0


class SubordinateLien:
    """
    Subordinate cash-flow bond repaid from residual surplus after the senior lien.

    Waterfall each year (Summary CH–CO), with the subordinate coverage factor:

        available    = surplus_fund.available_to_sub(y)
                       + (net_sub_revenue_y - net_senior_revenue_y)        # CH
        coverage_adj = available / DSC_SUB                                  # CJ
        interest_due = principal_balance * sub_rate  (+ accreted interest)  # CK
        interest_paid= min(interest_due + accrued_balance, coverage_adj)    # CL/CN
        accrued      accretes at sub_rate on the unpaid balance             # CO
        principal_paid = remaining coverage_adj, capped at balance          # CR/CN
    """

    def __init__(self, config: ModelConfig, summary: SummaryModel):
        self.cfg = config
        self.sm = summary

    def size(
        self,
        par_amount: float,
        senior: BondTranche,
        first_year: int,
        final_year: int,
        *,
        refunding: BondTranche | None = None,
        surplus_fund: SurplusFund | None = None,
    ) -> SubLienResult:
        cfg = self.cfg
        rate = cfg.sub_interest_rate
        coverage = cfg.dsc_sub

        if surplus_fund is None:
            surplus_fund = SurplusFund(cfg, self.sm).build(
                senior, refunding, first_year, final_year)

        principal_balance = par_amount   # CS — outstanding principal
        accrued_balance = 0.0            # CO — accreted unpaid interest
        rows: list[dict] = []

        for y in range(first_year, final_year + 1):
            # ── Cash available to the subordinate lien (Summary CH) ──────────
            available = surplus_fund.available_to_sub(y)
            # Break down where that cash comes from (for the sub-lien detail):
            #   available = excess-above-target (BY) + sub-incremental (BO−AX)
            #               + released reserve (at senior maturity).
            sf = surplus_fund._by_year.get(y)
            sub_incremental = max(
                0.0, self.sm.net_sub_revenue(y) - self.sm.net_senior_revenue(y))
            senior_residual = sf.senior_residual if sf else 0.0
            surplus_deposit = sf.deposit_to_reserve if sf else 0.0
            surplus_balance = sf.reserve_balance if sf else 0.0
            excess_to_sub = sf.excess_to_sub if sf else 0.0
            reserve_release = sf.reserve_release if sf else 0.0
            # Subordinate coverage factor (Summary CJ)
            coverage_adj = available / coverage if coverage else available

            # ── Interest: current coupon + accretion on unpaid balance ───────
            current_interest = principal_balance * rate
            interest_due = current_interest + accrued_balance
            interest_paid = min(interest_due, coverage_adj)

            # Unpaid interest accretes (compounds) into the accrued balance.
            accrued_balance = interest_due - interest_paid

            # ── Principal: residual cash after interest, capped at balance ───
            residual_cash = max(0.0, coverage_adj - interest_paid)
            principal_paid = min(principal_balance, residual_cash)
            principal_balance -= principal_paid

            rows.append({
                "year": y,
                # Derivation of "available to sub"
                "senior_residual": senior_residual,
                "surplus_deposit": surplus_deposit,
                "surplus_balance": surplus_balance,
                "excess_to_sub": excess_to_sub,
                "sub_incremental": sub_incremental,
                "reserve_release": reserve_release,
                "available_to_sub": available,
                "coverage_adj_available": coverage_adj,
                "current_interest": current_interest,
                "interest_paid": interest_paid,
                "accrued_balance": accrued_balance,
                "principal_paid": principal_paid,
                "principal_balance": principal_balance,
                "total_paid": interest_paid + principal_paid,
            })
            if principal_balance <= 0 and accrued_balance <= 0:
                break

        return SubLienResult(par_amount=par_amount, first_year=first_year,
                             final_year=final_year, rate=rate, coverage=coverage,
                             rows=rows)

    def size_par(
        self,
        senior: BondTranche,
        first_year: int,
        final_year: int,
        *,
        refunding: BondTranche | None = None,
        surplus_fund: SurplusFund | None = None,
        max_par: float = 50_000_000.0,
    ) -> float:
        """
        Largest subordinate par the residual surplus fully repays (principal +
        accreted interest) by ``final_year`` — derived from the district's AV /
        revenue rather than input.  Solved by bisection (repayment is monotone
        decreasing in par).  Rounded to $1,000.
        """
        if surplus_fund is None:
            surplus_fund = SurplusFund(self.cfg, self.sm).build(
                senior, refunding, first_year, final_year)

        def repaid(par: float) -> bool:
            if par <= 0:
                return True
            res = self.size(par, senior, first_year, final_year,
                            refunding=refunding, surplus_fund=surplus_fund)
            last = res.rows[-1]
            return last["principal_balance"] <= 1.0 and last["accrued_balance"] <= 1.0

        lo, hi = 0.0, max_par
        if repaid(hi):
            return round(hi / 1000.0) * 1000.0
        for _ in range(40):
            mid = (lo + hi) / 2.0
            if repaid(mid):
                lo = mid
            else:
                hi = mid
        return round(lo / 1000.0) * 1000.0

    def to_dataframe(self, result: SubLienResult) -> pd.DataFrame:
        return pd.DataFrame(result.rows)
