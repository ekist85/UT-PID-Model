"""
debt_service.py — Senior- and Subordinate-lien debt-service engines.

The senior lien is a **revenue bond**: principal is sized year-by-year so that
net annual debt service is fully covered by the pledged net revenue at the
target coverage ratio.  This reproduces the "Senior Lien DS" sheets, where the
serial principal grows as the assessed-value base (and therefore mill-levy
revenue) grows.

Sizing rule (per December-1 principal payment, matching the Excel FUDGE/iterate):

    P_t = FLOOR( ( NetRevenue_t / coverage + DSRF_interest_earnings
                   - outstanding_balance_t * rate ) / 5000 ) * 5000

with the par amount = sum(P_t) solved by fixed-point iteration (the outstanding
balance for early maturities depends on the eventual par).  Capitalized-interest
years carry zero principal, and the final maturity absorbs the released surplus
(debt-service-reserve) fund as a balloon.

The subordinate lien is a **cash-flow bond**: it accrues interest on its
balance and is paid only from residual surplus after the senior lien — modeled
in ``subordinate.py``.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date
from typing import Optional

import pandas as pd

from .config import ModelConfig
from .summary import SummaryModel


@dataclass
class CallProvisions:
    """
    Optional-redemption (call) provisions for a bond series.

    Utah PID bonds are issued with **call protection** until the
    first optional redemption date, then become callable at a **declining
    redemption premium** that steps down 1.00% per year to **par** (100%) — the
    standard official-statement structure.  These provisions are what make the
    refunding possible: the refunding bonds redeem the outstanding bonds at the
    applicable call price.

      * Non-callable before ``premium_call_date`` (call protection).
      * On/after ``premium_call_date``, callable at par plus a redemption
        premium of ``(premium_call_price - 100)`` that declines 1.00% per year.
      * Callable at par (100) once the premium has stepped down to zero
        (``par_call_date``).
    """
    premium_call_date: Optional[date] = None
    par_call_date: Optional[date] = None
    premium_call_price: float = 100.0   # e.g. 103.0 ⇒ a 3.00% initial premium

    def is_callable_on(self, as_of: date) -> bool:
        """True if the bonds may be optionally redeemed on ``as_of``."""
        return self.premium_call_date is not None and as_of >= self.premium_call_date

    def _whole_years_since_premium_call(self, as_of: date) -> int:
        """Completed years from the premium-call date to ``as_of`` (the 1%/yr step)."""
        years = as_of.year - self.premium_call_date.year
        if (as_of.month, as_of.day) < (self.premium_call_date.month,
                                       self.premium_call_date.day):
            years -= 1
        return years

    def call_price_on(self, as_of: date) -> Optional[float]:
        """
        Redemption price as a percent of par on ``as_of``.

        Par plus a redemption premium of ``(premium_call_price - 100)`` that
        declines 1.00% for each completed year after the premium-call date, to a
        floor of par (100%).  Returns ``None`` while the bonds are call-protected.
        """
        if self.premium_call_date is None or as_of < self.premium_call_date:
            return None
        initial_premium = self.premium_call_price - 100.0
        premium = max(initial_premium - self._whole_years_since_premium_call(as_of), 0.0)
        return 100.0 + premium


@dataclass
class PaymentRow:
    """One semi-annual coupon / principal payment."""
    payment_date: date
    principal: float
    interest: float
    capitalized_interest: float = 0.0     # portion covered by the CAPI fund
    dsrf_earnings: float = 0.0            # interest earnings credit on the DSRF
    surplus_release: float = 0.0          # DSRF released at final maturity

    @property
    def gross_total(self) -> float:
        return self.principal + self.interest

    @property
    def net_total(self) -> float:
        return self.gross_total - self.capitalized_interest - self.dsrf_earnings - self.surplus_release


def coupon_at(coupon_scale, term_bonds, rate: float, year: int) -> float:
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
class BondTranche:
    """A sized serial revenue bond (senior new-money or refunding)."""
    name: str
    rate: float
    coverage: float
    delivery: date
    first_principal_year: int
    final_year: int
    prin_month: int
    prin_day: int
    capi_end_year: int | None
    dsrf_deposit: float
    dsrf_earn_rate: float
    par_amount: float = 0.0
    call_provisions: Optional[CallProvisions] = None
    reoffering_yield: Optional[float] = None     # flat reoffering yield; None ⇒ priced at par
    coupon_scale: Optional[dict] = None          # {maturity_year: coupon} (pricing day; else flat rate)
    yield_scale: Optional[dict] = None           # {maturity_year: yield} curve (overrides flat)
    term_bonds: Optional[list] = None            # [(first_year, last_year, term_yield), ...]
    price_scale: Optional[dict] = None           # {maturity_year: price} — entered on pricing day
    schedule: list[PaymentRow] = field(default_factory=list)

    def coupon_for(self, year: int) -> float:
        """
        Coupon for a maturity — the per-maturity scale (carry-forward) or the
        flat rate.

        A term bond carries one coupon, entered on its FINAL maturity row, so an
        installment inside a term looks that row up rather than carrying forward
        from whatever precedes it.
        """
        return coupon_at(self.coupon_scale, self.term_bonds, self.rate, year)

    # ── Pricing (price / yield / premium-OID, priced to worst call) ──────────
    def _term_for(self, year: int):
        """Return (first, last, term_yield) if ``year`` is in a term bond, else None."""
        for tb in (self.term_bonds or []):
            first, last, ty = tb
            if first <= year <= last:
                return (first, last, ty)
        return None

    def yield_for(self, year: int) -> float:
        """
        Reoffering yield for a maturity.  Precedence: term-bond yield > yield
        curve (carry-forward) > flat reoffering yield > coupon (par).
        """
        term = self._term_for(year)
        if term is not None:
            return term[2]
        if self.yield_scale:
            from .config import _schedule_lookup
            return _schedule_lookup(self.yield_scale, year, self.rate)
        if self.reoffering_yield is not None:
            return self.reoffering_yield
        return self.rate

    def is_term(self, year: int) -> bool:
        return self._term_for(year) is not None

    def term_maturity(self, year: int):
        term = self._term_for(year)
        return term[1] if term is not None else None

    def _call_scenarios(self) -> list[tuple[date, float]]:
        cp = self.call_provisions
        if not cp or cp.premium_call_date is None:
            return []
        out = [(cp.premium_call_date, cp.premium_call_price)]
        if cp.par_call_date is not None:
            out.append((cp.par_call_date, 100.0))
        return out

    def price_for(self, year: int) -> float:
        """
        Yield-to-worst clean price per 100 for the maturity in ``year``.

        A serial maturity prices to its own date at its yield; a term-bond
        installment prices to the term maturity at the single term yield.
        """
        from .pricing import price_to_worst
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
            first, last, ty = term
            maturity = date(last, self.prin_month, self.prin_day)
            return price_to_worst(self.delivery, maturity, self.coupon_for(last), ty,
                                  self._call_scenarios())
        maturity = date(year, self.prin_month, self.prin_day)
        return price_to_worst(self.delivery, maturity, self.coupon_for(year),
                              self.yield_for(year), self._call_scenarios())

    def premium_for(self, year: int, principal: float) -> float:
        """Premium / (original-issue discount) on a maturity = (price − par) × principal."""
        if not principal:
            return 0.0
        return (self.price_for(year) / 100.0 - 1.0) * principal

    @property
    def total_premium(self) -> float:
        """Net premium / (discount) across all maturities (a Source of Funds)."""
        return sum(self.premium_for(p.payment_date.year, p.principal)
                   for p in self.schedule if p.principal)

    # ── Optional redemption ──────────────────────────────────────────────────
    def is_callable_on(self, as_of: date) -> bool:
        return self.call_provisions.is_callable_on(as_of) if self.call_provisions else False

    def call_price_on(self, as_of: date) -> Optional[float]:
        return self.call_provisions.call_price_on(as_of) if self.call_provisions else None

    def callable_principal_after(self, as_of: date) -> float:
        """
        Principal maturing strictly after ``as_of`` — the amount that would be
        redeemed (defeased) in an optional redemption / refunding on that date.
        """
        return sum(p.principal for p in self.schedule if p.payment_date > as_of)

    # ── Annual roll-ups (keyed by calendar year of December payment) ─────────
    def annual_net_ds(self) -> dict[int, float]:
        out: dict[int, float] = {}
        for p in self.schedule:
            out[p.payment_date.year] = out.get(p.payment_date.year, 0.0) + p.net_total
        return out

    def annual_gross_ds(self) -> dict[int, float]:
        out: dict[int, float] = {}
        for p in self.schedule:
            out[p.payment_date.year] = out.get(p.payment_date.year, 0.0) + p.gross_total
        return out

    def annual_principal(self) -> dict[int, float]:
        out: dict[int, float] = {}
        for p in self.schedule:
            if p.principal:
                out[p.payment_date.year] = out.get(p.payment_date.year, 0.0) + p.principal
        return out

    def outstanding_after(self, year: int) -> float:
        """Principal still outstanding after the December payment of ``year``."""
        paid = sum(p.principal for p in self.schedule if p.payment_date.year <= year)
        return max(self.par_amount - paid, 0.0)

    @property
    def max_annual_ds(self) -> float:
        return max(self.annual_net_ds().values()) if self.schedule else 0.0

    @property
    def total_net_ds(self) -> float:
        return sum(p.net_total for p in self.schedule)


class SeniorLienSizer:
    """
    Sizes a senior-lien serial revenue bond against a stream of net pledged
    revenue, at a target debt-service-coverage ratio.
    """

    def __init__(self, config: ModelConfig, summary: SummaryModel):
        self.cfg = config
        self.sm = summary

    def size(
        self,
        name: str,
        rate: float,
        coverage: float,
        delivery: date,
        first_principal_year: int,
        final_year: int,
        *,
        capi_end_year: int | None,
        dsrf_deposit: float,
        release_surplus: bool = True,
        call_provisions: Optional[CallProvisions] = None,
        reoffering_yield: Optional[float] = None,
        coupon_scale: Optional[dict] = None,
        yield_scale: Optional[dict] = None,
        term_bonds: Optional[list] = None,
        par_schedule: Optional[dict] = None,
        price_scale: Optional[dict] = None,
    ) -> BondTranche:
        cfg = self.cfg
        dsrf_earn = dsrf_deposit * cfg.interest_earn_rate

        principal_years = list(range(first_principal_year, final_year + 1))
        # Share of a year's interest accruing AFTER the principal date.
        post_prin = 0.5 if cfg.int_maturity > cfg.prin_maturity else 0.0

        def _wrap(par: float, interest_of) -> dict[int, float]:
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
                annual_int = interest_of(y, balance)
                avail = target - annual_int
                if release_surplus and y == final_year:
                    avail += dsrf_deposit  # released DSRF pays down the final maturity
                # Principal retired on the principal date stops earning the
                # coupons that follow it later in the same year.
                eff_rate = (annual_int / balance) if balance else rate
                avail /= max(1e-9, 1.0 - eff_rate * post_prin)
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
            return principals

        if par_schedule:
            # Manual amortization override (Debt Structure "Par Amount" column):
            # use the entered principal-by-maturity schedule instead of the wrap.
            principals = {y: float(par_schedule.get(y, 0.0)) for y in principal_years}
            par = sum(principals.values())
        else:
            # ── Bisection on par: f(par) = sum(principals) is monotone decreasing,
            #    so the consistent par solves f(par) = par. ───────────────────────
            lo, hi = 0.0, 1.0
            while sum(size_for_par(hi).values()) > hi:
                hi *= 2.0
                if hi > 1e12:
                    break
            for _ in range(100):
                mid = (lo + hi) / 2.0
                if sum(size_for_par(mid).values()) > mid:
                    lo = mid
                else:
                    hi = mid
            par = round((lo + hi) / 2.0 / 5000.0) * 5000.0
            principals = size_for_par(par)
            par = sum(principals.values())

            # Ensure full amortization: any rounding residual lands on final maturity.
            assigned = sum(principals.values())
            if par - assigned != 0:
                principals[final_year] = principals.get(final_year, 0.0) + (par - assigned)

        tranche = BondTranche(
            name=name, rate=rate, coverage=coverage, delivery=delivery,
            first_principal_year=first_principal_year, final_year=final_year,
            prin_month=cfg.prin_maturity, prin_day=cfg.prin_maturity_day_senior,
            capi_end_year=capi_end_year, dsrf_deposit=dsrf_deposit,
            dsrf_earn_rate=cfg.interest_earn_rate, par_amount=par,
            call_provisions=call_provisions, reoffering_yield=reoffering_yield,
            coupon_scale=coupon_scale, yield_scale=yield_scale, term_bonds=term_bonds,
            price_scale=price_scale,
        )
        tranche.schedule = self._build_schedule(tranche, principals, release_surplus)
        if coupon_scale:
            # Pricing-day: re-derive interest from per-maturity coupons.
            self._apply_coupon_scale(tranche, capi_end_year)
        return tranche

    @staticmethod
    def _apply_coupon_scale(tranche: "BondTranche", capi_end_year: int | None) -> None:
        """
        Recompute each period's interest from per-maturity coupons (the schedule
        is sized with the flat rate; on pricing day coupons differ by maturity).

        Annual interest in year t = sum over maturities y >= t of P_y x coupon_y;
        split evenly across the two semi-annual coupons.  Capitalized-interest
        years still capitalize the (recomputed) interest.
        """
        from .pricing import days_30_360

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
            prev = p.payment_date

    # ── Semi-annual amortization schedule ────────────────────────────────────
    def _build_schedule(
        self, t: BondTranche, principals: dict[int, float], release_surplus: bool,
    ) -> list[PaymentRow]:
        cfg = self.cfg
        rate = t.rate
        rows: list[PaymentRow] = []
        balance = t.par_amount
        dsrf_earn = t.dsrf_deposit * t.dsrf_earn_rate

        from .pricing import days_30_360

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

        return rows


def schedule_dataframe(tranche: BondTranche) -> pd.DataFrame:
    """Flat semi-annual schedule for a single tranche."""
    return pd.DataFrame([{
        "payment_date": p.payment_date,
        "principal": p.principal,
        "interest": p.interest,
        "capitalized_interest": p.capitalized_interest,
        "dsrf_earnings": p.dsrf_earnings,
        "surplus_release": p.surplus_release,
        "gross_total": p.gross_total,
        "net_total": p.net_total,
    } for p in tranche.schedule])


def standard_dsrf(tranche: BondTranche) -> float:
    """
    Debt-service-reserve fund by the standard 3-prong test, rounded to $5,000:
    the least of 10% of par, maximum annual NET debt service, and 125% of average
    annual NET debt service.  Derived from the sized bond, so it follows the
    district's taxable value and tax revenue.

    Net debt service (principal + interest less capitalized interest, DSRF
    earnings and the final-year DSRF release) is the amount actually paid from
    pledged revenue, so the reserve is sized to that.  Both the maximum- and the
    average-annual prongs EXCLUDE the senior bonds' final maturity year: its net
    debt service is distorted by the released DSRF paying down the final maturity,
    so it is not representative of ongoing annual debt service.
    """
    annual = tranche.annual_net_ds()
    ongoing = {y: v for y, v in annual.items() if y != tranche.final_year}
    vals = [v for v in ongoing.values() if v > 1.0]
    if not vals or tranche.par_amount <= 0:
        return 0.0
    max_ds = max(vals)
    raw = min(0.10 * tranche.par_amount, max_ds, 1.25 * sum(vals) / len(vals))
    return round(raw / 5000.0) * 5000.0


def size_senior_with_dynamic_dsrf(sizer: "SeniorLienSizer", **kwargs) -> BondTranche:
    """
    Size the senior bond and its DSRF together: the DSRF depends on the debt
    service, which depends slightly on the DSRF (earnings credit + final-year
    release).  Iterate to a fixed point (a few passes).
    """
    dsrf = kwargs.pop("dsrf_deposit", 0.0) or 0.0
    tranche = None
    for _ in range(6):
        tranche = sizer.size(dsrf_deposit=dsrf, **kwargs)
        new_dsrf = standard_dsrf(tranche)
        if abs(new_dsrf - dsrf) < 1000.0:
            dsrf = new_dsrf
            break
        dsrf = new_dsrf
    return sizer.size(dsrf_deposit=dsrf, **kwargs)


def senior_coverage_schedule(cfg, summary, senior, refunding=None):
    """
    Annual debt-service-coverage schedule for the senior lien.

    For each collection year, coverage = net pledged senior revenue / senior
    net annual debt service, compared against the target coverage factor
    (``dsc_senior``, or ``dsc_refunding`` once the bonds are refunded).

    Returns a list of row dicts: year, net_senior_revenue, senior_net_ds,
    coverage, target, meets_target.
    """
    senior_net = senior.annual_net_ds()
    ref_net = refunding.annual_net_ds() if refunding else {}
    refund_year = refunding.delivery.year if refunding else None

    rows = []
    for r in summary.rows:
        y = r.collection_year
        if refunding is not None and y > refund_year:
            ds = ref_net.get(y, 0.0)
            target = cfg.dsc_refunding
        else:
            ds = senior_net.get(y, 0.0)
            target = cfg.dsc_senior
        if ds <= 0:
            continue
        coverage = r.net_senior_revenue / ds
        rows.append({
            "year": y,
            "net_senior_revenue": round(r.net_senior_revenue, 2),
            "senior_net_ds": round(ds, 2),
            "coverage": round(coverage, 4),
            "target": target,
            "meets_target": coverage >= target - 1e-9,
        })
    return rows


def senior_coverage_dataframe(cfg, summary, senior, refunding=None) -> pd.DataFrame:
    """Annual senior debt-service-coverage schedule as a DataFrame."""
    return pd.DataFrame(senior_coverage_schedule(cfg, summary, senior, refunding))
