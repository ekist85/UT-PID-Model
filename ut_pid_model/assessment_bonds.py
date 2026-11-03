"""
Special assessment bond sizing (Inputs!ASSESSMENT_BONDS = Yes).

A Utah public infrastructure district may levy assessments on the benefited
lots under the Assessment Area Act (UCA Title 11, Chapter 42) and issue
assessment bonds secured by those liens — a second, separate security from the
ad valorem mill levy that the senior and subordinate liens are sized against.
The two programs do not compete: an assessment is a lien on the lot, not a
charge against the levy, so assessment par does not consume room under the
§ 17D-4-303 fifteen-mill cap and assessment installments are not netted from
the revenue available to the mill levy bonds.  Nothing here runs, and nothing appears in the memo or
the workbook, unless the Inputs cell says Yes.  The amounts are used exactly
as entered on the Inputs page:

  • Developer Bond — DEVELOPER_BOND per developed lot ("Developer Assessment
    Bond Initial Amount", e.g. $40,000); with that cell blank or zero, each
    series is instead sized to the MAXIMUM the finished-lot value of its block
    supports at the Loan-to-Value ratio (par = lot value ÷ LTV).  The first bond is sold
    ASSESSMENT_FUND_TIMING years before the first lot delivery it funds and
    funds ASSESSMENT_FUND_TERM years of deliveries; each later bond funds the
    next block of deliveries and is likewise sold ASSESSMENT_FUND_TIMING years
    before that block's first delivery (timing 1, term 3: 2026 funds
    2027-2029, 2029 funds 2030-2032, 2032 funds 2033-2035, …), at
    DEVELOPER_BOND × the lots in its block.  Sales fall on
    the mill levy bond dates.  Lots delivered before the first bond's first funded
    year are not financed.
  • Payoff at conversion — on the same schedule that lots convert to homes
    (the Home Closings grid), the developer pays the lot's balance down to
    RESIDENT_BOND ("Resident Assessment Bond Remaining Amount", e.g. $10,000)
    — about DEVELOPER_BOND − RESIDENT_BOND on a fresh lot — plus the interest
    accrued on the amount paid off; those bonds are redeemed at par.  The
    homeowner assumes exactly the RESIDENT_BOND balance on the bond's original
    amortization schedule to its final maturity — no new bond is sold.  The developer pays the installments on lots it still holds.

Level annual installments at the Special Assessment Interest Rate
(Inputs!BONDRATE_SPECIAL) over the program's Final Maturity (Inputs!FINAL_MAT_YRS,
the same term the senior lien runs), beginning the year after each series is dated;
par = the assessment principal exactly, with interest paid on top; interest
for the first CAPI_YEARS_SA years is capitalized from proceeds (lot owners'
installments start after that period, separate from the GO CAPI_YEARS); costs
of issuance use the SAME assumptions as the mill levy bonds (Municipal Advisor %, Bond
Discount % and COI % of par); an optional debt-service reserve
(ASSESSMENT_DSRF_PCT of par, 0 by default) is funded from proceeds.  The
finished-lot value ÷ par is reported against the required value-to-lien
(LTV) as a test.

Assessment bonds are secured by the liens, not the District's taxing power, so
their par does NOT count against the § 17D-4-303 fifteen-mill cap and their
installments sit outside the mill levy entirely.  What the homeowner pays is
the sum of the two: the mill levy on the taxable value of the home plus the
annual installment on the assessment still attached to the lot.
"""
from __future__ import annotations

import calendar
import datetime as _dt
import math
from dataclasses import dataclass, field
from datetime import date
from dateutil.relativedelta import relativedelta

import pandas as pd

from .config import ModelConfig
from .summary import SummaryModel


def _md(year: int, month: int, day: int = 1) -> date:
    """A month/day in a year, clamped to the length of the month."""
    return date(year, month, min(day, calendar.monthrange(year, month)[1]))


def _add_months(d: date, months: int) -> date:
    return _md(*_ym_plus(d.year, d.month, months), d.day)


def _ym_plus(year: int, month: int, months: int) -> tuple[int, int]:
    total = (year * 12 + (month - 1)) + months
    return total // 12, total % 12 + 1


def _go_first_interest(dated: date, prin_month: int, int_month: int, day: int = 1) -> date:
    """First coupon date strictly after the dated date."""
    lo = _md(dated.year, min(prin_month, int_month), day)
    hi = _md(dated.year, max(prin_month, int_month), day)
    if lo <= dated:
        return hi if hi > dated else _md(dated.year + 1, min(prin_month, int_month), day)
    return lo


def _go_first_principal(dated: date, capi_years: float, prin_month: int,
                        int_month: int, first_int: date, day: int = 1) -> date:
    """
    First principal date — the first principal date strictly AFTER the
    capitalized-interest period (dated + ``capi_years``).

    Nothing amortizes while interest is being capitalized, which is the same
    rule the mill levy bonds follow: a series dated 9/30/2026 with three CAPI
    years capitalizes every coupon through 9/30/2029 and pays its first
    principal in 2030.
    """
    end_capi = _add_months(dated, int(round(12 * capi_years)))
    if end_capi >= _md(end_capi.year, prin_month, day):
        candidate = _md(end_capi.year + 1, prin_month, day)
    else:
        candidate = _md(end_capi.year, prin_month, day)
    return max(candidate, first_int)


# ── Inputs helpers ─────────────────────────────────────────────────────────
def parse_value_to_lien(raw, default: float = 3.0) -> float:
    """
    Read the Loan-to-Value / value-to-lien input as a VALUE ÷ LIEN multiple.

    Accepts "3:1", "3,1", "3/1", "3 to 1", "3", 3, 3.0, a lien-to-value fraction
    ("1:3", 0.333 → 3.0) and the time value Excel produces when "3:1" is typed
    into a General cell (03:01 → 3.0).  Blank / unreadable → ``default``.
    """
    if raw is None or raw == "":
        return default
    if isinstance(raw, (_dt.time, _dt.datetime)):   # Excel turned "3:1" into 03:01
        return _ratio(float(raw.hour), float(raw.minute or 1), default)
    if isinstance(raw, (int, float)):
        return _ratio(float(raw), 1.0, default)
    s = str(raw).strip().lower().replace("to", " ")
    for ch in (":", ",", "/", "x", ";"):
        s = s.replace(ch, " ")
    toks = [t for t in s.split() if t]
    try:
        nums = [float(t) for t in toks]
    except ValueError:
        return default
    if not nums:
        return default
    if len(nums) == 1:
        return _ratio(nums[0], 1.0, default)
    return _ratio(nums[0], nums[1], default)


def _uwd_rate(cfg) -> float:
    """
    Underwriter's discount on an assessment series, as a share of par.

    Assessment bonds are sold on the lots' security rather than the district's
    levy, so they take the SUBORDINATE lien's discount (UWD_SUB), not the
    senior's — the senior rate understates what a lien-secured developer bond
    costs to place.
    """
    return float(getattr(cfg, "uwd_sub", 0.0) or 0.0)


def _coi_rate(cfg) -> float:
    """
    Costs of issuance on an assessment series, as a share of par.

    Utah enters COI as a DOLLAR amount for the mill levy financing, which cannot
    be charged whole to each assessment series without multiplying it by the
    number of series.  "The same assumptions as the mill levy bonds" therefore
    means the same RATE: the dollar COI as a share of the senior par it was
    quoted against.  With no senior par to measure against, nothing is charged.
    """
    coi = float(getattr(cfg, "coi", 0.0) or 0.0)
    basis = float(getattr(cfg, "_senior_par_for_coi", 0.0) or 0.0)
    return (coi / basis) if (coi > 0 and basis > 0) else 0.0


def _ratio(a: float, b: float, default: float) -> float:
    if a <= 0 or b <= 0:
        return default
    r = a / b
    if r < 1.0:                               # entered as lien ÷ value → invert
        r = 1.0 / r
    return r


def _with_prin_day(d: date, cfg: ModelConfig) -> date:
    day = min(int(getattr(cfg, "prin_maturity_day_senior", 1) or 1), calendar.monthrange(d.year, d.month)[1])
    return d.replace(day=day)


def go_first_principal_date(dated: date, capi_years: float, cfg: ModelConfig) -> date:
    """GO rule: the first principal date strictly after dated + capi_years (no
    principal is amortized during the CAPI period)."""
    fi = _go_first_interest(dated, cfg.prin_maturity, cfg.int_maturity, cfg.prin_maturity_day_senior)
    return _go_first_principal(dated, capi_years, cfg.prin_maturity, cfg.int_maturity, fi, cfg.prin_maturity_day_senior)


def capi_end_date(dated: date, capi_years: float, cfg: ModelConfig) -> date | None:
    """End of the CAPI period (dated + capi_years), the same window the mill levy bonds
    use: coupons on or before it are capitalized (a bond dated 8/15/26 with 3
    CAPI years capitalizes every coupon through 8/15/29).  None without CAPI."""
    if not capi_years:
        return None
    return dated + relativedelta(months=int(round(12 * capi_years)))


def assessment_first_principal_date(dated: date, capi_years: float, cfg: ModelConfig) -> date:
    """First principal payment on a special assessment bond -- the same rule as
    the mill levy bonds: the first principal date after the CAPI period (8/15/30 for a
    bond dated 8/15/26 with 3 CAPI years; the year after dating without CAPI)."""
    return go_first_principal_date(dated, capi_years, cfg)


def level_payment(balance: float, rate: float, years: int) -> float:
    """Level annual payment that retires ``balance`` over ``years`` at ``rate``."""
    if years <= 0 or balance <= 0:
        return 0.0
    if rate <= 0:
        return balance / years
    return balance * rate / (1.0 - (1.0 + rate) ** (-years))


def level_balance(payment: float, rate: float, years: int) -> float:
    """Principal that a level annual ``payment`` (P+I) retires over ``years`` at
    ``rate`` — the inverse of level_payment."""
    if years <= 0 or payment <= 0:
        return 0.0
    if rate <= 0:
        return payment * years
    return payment * (1.0 - (1.0 + rate) ** (-years)) / rate


def three_prong_reserve(par: float, rate: float, term: int, capi_years: int = 0) -> dict:
    """
    Debt service reserve fund under the three-prong test: the LESSER of
      (1) 10% of par,
      (2) maximum annual debt service, and
      (3) 125% of average annual debt service,
    measured on the series held to maturity (interest only during the CAPI
    years, level installments after).  Returns the three prongs, the reserve
    and which prong governs.
    """
    if par <= 0 or term <= 0:
        return {"ten_pct": 0.0, "mads": 0.0, "avg125": 0.0, "reserve": 0.0, "basis": ""}
    capi = max(min(int(capi_years), term - 1), 0)
    ds, bal = [], par
    for k in range(1, term + 1):
        i = bal * rate
        if k <= capi:
            p = 0.0
        else:
            p = min(max(level_payment(bal, rate, term - k + 1) - i, 0.0), bal)
        if k == term:
            p = bal
        ds.append(p + i)
        bal -= p
    ten = 0.10 * par
    mads = max(ds)
    avg125 = 1.25 * sum(ds) / len(ds)
    reserve = min(ten, mads, avg125)
    basis = ("10% of par" if reserve == ten else
             "maximum annual debt service" if reserve == mads else
             "125% of average annual debt service")
    return {"ten_pct": ten, "mads": mads, "avg125": avg125, "reserve": reserve, "basis": basis}


# Payoffs at conversion are redeemed on the semiannual interest date following
# the closing, so on average one half-year of interest accrues on the principal
# paid off since the prior installment.
PAYOFF_ACCRUAL_YEARS = 0.5

DEVELOPER = "Developer"


# ── Data classes ───────────────────────────────────────────────────────────
@dataclass
class AssessmentSeries:
    """One Developer Bond series (one per issuance window)."""
    program: str
    series_id: int
    dated_date: date
    units: float                     # lots assessed
    unit_value: float                # finished-lot value of those lots ($, escalated)
    lien_per_unit: float             # assessment per lot ($), as entered
    required_value_to_lien: float    # LTV input (e.g. 3.0 for 3:1)
    par_amount: float
    interest_rate: float
    term_years: int
    muniadv_fee: float               # unused in Utah (no municipal advisor fee cell); kept at 0.0
    bond_discount: float             # UWD_SUB x par — the underwriter's discount
    coi_pct_fee: float               # COI as a share of senior par, applied to this series
    reserve_fund: float              # DSRF funded from proceeds: three-prong test (lesser of the three below)
    dsrf_ten_pct: float = 0.0        # prong 1: 10% of par
    dsrf_mads: float = 0.0           # prong 2: maximum annual debt service
    dsrf_avg125: float = 0.0         # prong 3: 125% of average annual debt service
    dsrf_basis: str = ""             # which prong governs
    dsrf_haircut: float = 1.0        # Inputs DSRF_SA_PCT: share of the three-prong amount actually deposited
    dsrf_balance: float = 0.0        # running reserve balance (cashflow)
    dsrf_released_payoff: float = 0.0  # released pro rata as lots pay off, credited to the developer's payoff
    dsrf_releases: dict = field(default_factory=dict)   # {conversion year: reserve released with that year's payoffs}
    dsrf_applied_year: int | None = None                # year the remaining reserve was applied (final maturity)
    dsrf_applied_maturity: float = 0.0  # remaining reserve applied to the final year's installment
    capi_years: int = 0              # capitalized-interest years (Inputs!CAPI_YEARS_SA)
    capitalized_interest: float = 0.0  # funded from proceeds: exactly the coupons through the end of the CAPI period
    capi_fund: float = 0.0           # = capitalized_interest once sized (kept for reporting)
    first_principal_date: date | None = None   # first principal: the year after the CAPI period ends
    capi_end_date: date | None = None          # last capitalized coupon date (dated + capi_years)
    capi_interest_paid: float = 0.0  # interest paid from the fund (incl. any release credit)
    window_start: int = 0            # first / last delivery year this series funds
    window_end: int = 0
    # Per-product cohorts within the series (assessments can differ by product):
    # {product: {"units", "per_unit", "res_per_unit", "dev_units", "dev_bal_pu",
    #            "res_units", "res_bal", "xfer_balance", "xfer_units", "xfer_unit_years"}}
    cohorts: dict = field(default_factory=dict)
    # Running state (filled by the program cashflow) — series aggregates of the cohorts:
    dev_units: float = 0.0           # lots still developer-held (unconverted)
    dev_balance_per_unit: float = 0.0
    resident_units: float = 0.0      # lots converted — balance carried by homeowners
    resident_balance: float = 0.0    # aggregate balance held by homeowners
    outstanding: float = 0.0
    rows: list = field(default_factory=list)   # annual {year, principal, interest}
    redemptions: dict = field(default_factory=dict)   # {year: principal paid off at conversion}
    transfers: dict = field(default_factory=dict)     # {year: balance transferred to homeowners}
    payoff_principal: float = 0.0    # principal paid off by the developer at conversion
    payoff_interest: float = 0.0     # accrued interest paid by the developer with those payoffs
    payoff_interest_capi: float = 0.0  # accrued interest on payoffs inside the CAPI period (from the fund, not the developer)
    # Transfer statistics (for the homeowner installment / effective rate)
    xfer_balance: float = 0.0
    xfer_units: float = 0.0
    xfer_unit_years: float = 0.0     # Σ units × remaining years at transfer

    # Aliases used by the report / memo
    @property
    def lots(self) -> float:
        return self.units

    @property
    def lot_value(self) -> float:
        return self.unit_value

    @property
    def lien_per_lot(self) -> float:
        return self.lien_per_unit

    @property
    def coi(self) -> float:
        """Costs of issuance — identical line items to the senior lien series."""
        return self.muniadv_fee + self.bond_discount + self.coi_pct_fee

    @property
    def project_proceeds(self) -> float:
        return max(self.par_amount - self.coi - self.reserve_fund - self.capitalized_interest, 0.0)

    @property
    def first_payment_date(self) -> date:
        """First principal date (the year after the CAPI period; else the year after dating)."""
        if self.first_principal_date is not None:
            return self.first_principal_date
        return date(self.dated_date.year + 1, self.dated_date.month, self.dated_date.day)

    @property
    def final_maturity(self) -> date:
        return date(self.dated_date.year + self.term_years, self.dated_date.month,
                    self.dated_date.day)

    @property
    def total_interest(self) -> float:
        return sum(r["interest"] for r in self.rows) + self.payoff_interest + self.payoff_interest_capi

    @property
    def total_principal(self) -> float:
        return sum(r["principal"] for r in self.rows) + self.payoff_principal

    @property
    def value_to_lien(self) -> float:
        """Finished-lot value ÷ par (the assessment on the lots)."""
        return (self.unit_value / self.par_amount) if self.par_amount > 0 else 0.0

    @property
    def meets_value_to_lien(self) -> bool:
        return self.value_to_lien + 1e-9 >= self.required_value_to_lien

    @property
    def overleveraged(self) -> bool:
        """Finished-lot value ÷ par is below the required value-to-lien (e.g. 3:1)."""
        return not self.meets_value_to_lien

    @property
    def max_par_at_required(self) -> float:
        """Largest par the finished-lot value supports at the required value-to-lien."""
        return (self.unit_value / self.required_value_to_lien) if self.required_value_to_lien > 0 else 0.0

    @property
    def overleverage_excess(self) -> float:
        """Par above what the lot value supports at the required ratio ($)."""
        return max(self.par_amount - self.max_par_at_required, 0.0)


class AssessmentProgram:
    """
    The Developer Bond program: a series per issuance window at ``per_unit`` ×
    lots, its year-by-year cashflow (installments split developer- /
    homeowner-paid, payoffs at conversion, balances transferred to homeowners)
    and its totals.
    """

    DENOM = 5_000.0
    payoff_at_conversion = True

    def __init__(self, name: str, cfg: ModelConfig, per_unit: float, go_dated_dates=None):
        self.name = name
        self.cfg = cfg
        self.per_unit = per_unit
        # senior lien series dated dates, by year — assessment bonds sold in a year a GO
        # series is dated take the identical date; otherwise the GO convention.
        self.go_dates_by_year = {d.year: d for d in (go_dated_dates or [])}
        self.series: list[AssessmentSeries] = []
        self.annual: pd.DataFrame = pd.DataFrame()
        self.max_to_ltv: bool = False
        self.unmatched_closings: float = 0.0
        self.unmatched_by_year: dict = {}      # conversions with no Developer Bond lot (early lots / excess)
        self.units_not_issued: float = 0.0     # lots never financed (par rounds to zero)
        self.units_unfunded_early: float = 0.0 # lots delivered before the first bond could be sold
        self.first_fundable_year: int | None = None

    @property
    def per_lot(self) -> float:          # alias used by the report / memo
        return self.per_unit

    def dated_for_year(self, year: int) -> date:
        """Sale date aligned with the mill levy bonds: the senior lien series dated in that year
        if there is one, else the GO dated-date convention (PRIN_MATURITY / PRIN_DAY)."""
        d = self.go_dates_by_year.get(year)
        if d is not None:
            return d
        return _with_prin_day(date(year, self.cfg.prin_maturity, 1), self.cfg)

    # ── Sizing + cashflow ─────────────────────────────────────────────────
    def run(self, units_by_year: dict, value_by_year: dict, conversions_by_year: dict | None = None,
            resident_per_unit: float = 0.0, window_years: int = 1, lead_years: float = 0.0,
            max_to_ltv: bool = False, units_by_year_prod: dict | None = None,
            value_by_year_prod: dict | None = None, conversions_by_year_prod: dict | None = None,
            prod_params: dict | None = None) -> "AssessmentProgram":
        """
        ``window_years`` (Inputs!ASSESSMENT_FUND_TERM) — each series funds this many
        consecutive years of deliveries.  ``lead_years`` (Inputs!ASSESSMENT_FUND_TIMING)
        — the first series is dated this many years BEFORE the first delivery it
        funds; later series are sold (window − lead) years apart.
        ``resident_per_unit`` — the balance left on each converting lot for the
        homeowner (RESIDENT_BOND); the developer pays off everything above it.
        """
        cfg = self.cfg
        vtl = float(getattr(cfg, "value_to_lien", 3.0) or 3.0)
        rate = float(getattr(cfg, "assessment_rate", 0.07) or 0.0)
        term = int(getattr(cfg, "assessment_term_yrs", 25) or 25)
        capi_yrs = max(int(round(float(getattr(cfg, "assessment_capi_years", 0.0) or 0.0))), 0)
        conversions_by_year = conversions_by_year or {}
        # Per-product detail: fall back to a single pseudo-product carrying the
        # program-wide amounts when the caller passes none.
        if not units_by_year_prod:
            units_by_year_prod = {y: {"All": n} for y, n in units_by_year.items()}
            value_by_year_prod = {y: {"All": v} for y, v in (value_by_year or {}).items()}
            conversions_by_year_prod = {y: {"All": n} for y, n in conversions_by_year.items()}
            prod_params = {"All": {"developer": (None if max_to_ltv else self.per_unit),
                                   "resident": resident_per_unit, "ltv": vtl}}
        prod_params = prod_params or {}
        self.prod_params = prod_params
        self.series = []
        self.max_to_ltv = bool(max_to_ltv)
        _any_amount = any((v.get("developer") or 0) > 0 for v in prod_params.values())
        if (not _any_amount and not max_to_ltv) or not units_by_year:
            self.annual = pd.DataFrame()
            return self

        years = sorted(set(units_by_year) | set(conversions_by_year))
        first_delivery_year = min(units_by_year)
        last_issue_year = max(units_by_year)
        W = max(int(window_years or 1), 1)
        lead_yrs = int(round(max(float(lead_years or 0.0), 0.0)))
        # The first bond is dated `lead` years before the first delivery it funds
        # (no earlier than the first delivery year or the current year, whichever
        # is earlier), so lots delivered before its first funded year are not
        # financed by this program.  It funds W consecutive delivery years; each
        # later bond funds the next W-year block and is likewise dated `lead`
        # years before that block's first delivery (term 3, timing 1 → sales
        # every 3 years, each one year ahead of its lots).
        floor_year = min(first_delivery_year, date.today().year)
        first_issue = max(first_delivery_year - lead_yrs, floor_year)
        first_window = first_issue + lead_yrs
        self.first_fundable_year = first_window
        self.units_unfunded_early = sum(n for k, n in units_by_year.items() if k < first_window)
        self.units_unfunded_early_prod = {}
        for k, by_p in units_by_year_prod.items():
            if k < first_window:
                for p, n in by_p.items():
                    self.units_unfunded_early_prod[p] = self.units_unfunded_early_prod.get(p, 0.0) + n

        pend_units = pend_val = 0.0
        pend_start = pend_issue = None
        pend_u_p, pend_v_p = {}, {}          # per-product lots / lot value in the pending block
        sid = 0
        k = 0
        y = first_window
        while y <= last_issue_year:
            w_end = min(y + W - 1, last_issue_year)
            u = sum(units_by_year.get(k_, 0.0) for k_ in range(y, w_end + 1))
            if u > 0 and pend_start is None:
                pend_start = y
                pend_issue = max(y - lead_yrs, first_issue)
            k += 1
            pend_units += u
            pend_val += sum(value_by_year.get(k_, 0.0) for k_ in range(y, w_end + 1))
            for k_ in range(y, w_end + 1):
                for p, n in units_by_year_prod.get(k_, {}).items():
                    pend_u_p[p] = pend_u_p.get(p, 0.0) + n
                for p, v in (value_by_year_prod or {}).get(k_, {}).items():
                    pend_v_p[p] = pend_v_p.get(p, 0.0) + v
            y = w_end + 1
            if pend_units <= 0:
                continue
            # Par by product: lots × the product's assessment, or — with the
            # developer amount blank — the maximum the product's finished-lot
            # value supports at its LTV (value ÷ LTV).  The series par is the
            # sum, rounded down to the denomination.
            par_p, ltv_p = {}, {}
            for p, n in pend_u_p.items():
                if n <= 0:
                    continue
                pp = prod_params.get(p, {})
                d_amt = float(pp.get("developer") or 0.0)
                l_req = float(pp.get("ltv") or vtl or 3.0)
                ltv_p[p] = l_req
                if d_amt > 0:
                    par_p[p] = n * d_amt
                elif pp.get("max_at_ltv", False):
                    par_p[p] = (pend_v_p.get(p, 0.0) / l_req) if l_req > 0 else 0.0
                else:
                    par_p[p] = 0.0                      # no assessment on this product
            par_p = {p: v for p, v in par_p.items() if v > 0}
            raw_par = sum(par_p.values())
            par = math.floor(raw_par / self.DENOM) * self.DENOM
            _scale = (par / raw_par) if raw_par > 0 else 0.0
            # Series value-to-lien requirement: par-weighted average of the products' LTVs.
            vtl_series = (sum(par_p[p] * ltv_p[p] for p in par_p) / raw_par) if raw_par > 0 else vtl
            # No sizing threshold for assessment bonds: every delivery block with
            # lots gets its own series, however small (the GO SIZING_THRESHOLD
            # does not apply here).
            if par <= 0:
                continue
            sid += 1
            dated = self.dated_for_year(min(pend_issue, pend_start))
            _capi = min(capi_yrs, max(term - 2, 0))
            _dsrf = three_prong_reserve(par, rate, term, _capi)
            _hair = float(getattr(cfg, "assessment_dsrf_pct", 1.0)) if getattr(cfg, "assessment_dsrf", True) else 0.0
            cohorts = {}
            for p, n in pend_u_p.items():
                if n <= 0 or par_p.get(p, 0.0) <= 0:
                    continue
                pp = prod_params.get(p, {})
                cohorts[p] = {"units": n, "per_unit": par_p[p] * _scale / n,
                              "res_per_unit": float(pp.get("resident") or 0.0),
                              "installment": float(pp.get("installment") or 0.0),
                              "dev_units": n, "dev_bal_pu": par_p[p] * _scale / n,
                              "res_units": 0.0, "res_bal": 0.0,
                              "xfer_balance": 0.0, "xfer_units": 0.0, "xfer_unit_years": 0.0}
            # Series lots and finished-lot value cover the ASSESSED products only;
            # products with no assessment are neither funded nor paid off.
            a_units = sum(c["units"] for c in cohorts.values()) or pend_units
            a_val = sum(pend_v_p.get(p, 0.0) for p in cohorts) if cohorts else pend_val
            self.series.append(AssessmentSeries(
                program=self.name, series_id=sid, dated_date=dated, units=a_units,
                unit_value=a_val, lien_per_unit=par / a_units, required_value_to_lien=vtl_series,
                cohorts=cohorts,
                par_amount=par, interest_rate=rate, term_years=term,
                muniadv_fee=0.0, bond_discount=par * _uwd_rate(cfg),
                coi_pct_fee=par * _coi_rate(cfg), reserve_fund=_dsrf["reserve"] * _hair,
                dsrf_ten_pct=_dsrf["ten_pct"], dsrf_mads=_dsrf["mads"], dsrf_avg125=_dsrf["avg125"],
                dsrf_basis=(_dsrf["basis"] if _hair > 0 else "none (DSRF_SA = No)"), dsrf_haircut=_hair,
                capi_years=min(capi_yrs, max(term - 2, 0)),
                capitalized_interest=par * rate * min(capi_yrs, max(term - 2, 0)),
                capi_fund=par * rate * min(capi_yrs, max(term - 2, 0)),
                first_principal_date=assessment_first_principal_date(dated, min(capi_yrs, max(term - 2, 0)), cfg),
                capi_end_date=capi_end_date(dated, min(capi_yrs, max(term - 2, 0)), cfg),
                dev_units=pend_units, dev_balance_per_unit=par / pend_units, outstanding=par,
                window_start=pend_start, window_end=w_end,
            ))
            pend_units = pend_val = 0.0
            pend_start = pend_issue = None
            pend_u_p, pend_v_p = {}, {}
        self.units_not_issued = pend_units

        # Year order: scheduled installment → new series dated → conversions.
        last_year = max(years[-1], max((s.final_maturity.year for s in self.series), default=years[-1]))
        start_year = min([years[0]] + [s.dated_date.year for s in self.series])
        # Two passes: the first measures the coupons actually capitalized on each
        # series (they depend on payoffs during the CAPI window); the capitalized
        # amount is then set to exactly that figure and the cashflow is rebuilt so
        # the issuance-year net proceeds are right.  Balances, payoffs and
        # installments are identical in both passes.
        for _pass in (1, 2):
            for s in self.series:
                for c in s.cohorts.values():
                    c["dev_units"], c["dev_bal_pu"] = c["units"], c["per_unit"]
                    c["res_units"] = c["res_bal"] = 0.0
                    c["xfer_balance"] = c["xfer_units"] = c["xfer_unit_years"] = 0.0
                s.dev_units, s.dev_balance_per_unit = s.units, s.par_amount / s.units
                s.resident_units = s.resident_balance = 0.0
                s.outstanding = s.par_amount
                s.rows, s.redemptions, s.transfers = [], {}, {}
                s.payoff_principal = s.payoff_interest = s.payoff_interest_capi = 0.0
                s.dsrf_balance = s.reserve_fund
                s.dsrf_released_payoff = s.dsrf_applied_maturity = 0.0
                s.dsrf_releases, s.dsrf_applied_year = {}, None
                s.xfer_balance = s.xfer_units = s.xfer_unit_years = 0.0
                s.capi_interest_paid = 0.0
            self.unmatched_by_year = {}
            rows = self._cashflow(units_by_year, conversions_by_year, resident_per_unit, years, start_year, last_year,
                                  conversions_by_year_prod)
            if _pass == 1:
                if not any(s.capi_years for s in self.series):
                    break
                for s in self.series:
                    s.capitalized_interest = s.capi_interest_paid
                    s.capi_fund = s.capitalized_interest
        for s in self.series:
            for c in s.cohorts.values():
                if c.get("installment", 0.0) > 0 and c["xfer_units"] > 0:
                    c["res_per_unit"] = c["xfer_balance"] / c["xfer_units"]   # realised average balance left
        self.annual = pd.DataFrame(rows)
        return self

    @staticmethod
    def _sync(s: "AssessmentSeries") -> None:
        """Refresh the series aggregates from its product cohorts."""
        if not s.cohorts:
            return
        du = sum(c["dev_units"] for c in s.cohorts.values())
        dh = sum(c["dev_units"] * c["dev_bal_pu"] for c in s.cohorts.values())
        s.dev_units = du
        s.dev_balance_per_unit = (dh / du) if du > 0 else 0.0
        s.resident_units = sum(c["res_units"] for c in s.cohorts.values())
        s.resident_balance = sum(c["res_bal"] for c in s.cohorts.values())
        s.xfer_balance = sum(c["xfer_balance"] for c in s.cohorts.values())
        s.xfer_units = sum(c["xfer_units"] for c in s.cohorts.values())
        s.xfer_unit_years = sum(c["xfer_unit_years"] for c in s.cohorts.values())

    def _cashflow(self, units_by_year, conversions_by_year, resident_per_unit, years, start_year, last_year,
                  conversions_by_year_prod=None):
        cfg = self.cfg
        rows = []
        unmatched = 0.0
        unfunded_pool = self.units_unfunded_early
        unfunded_pool_p = dict(getattr(self, "units_unfunded_early_prod", {}) or {})
        conversions_by_year_prod = conversions_by_year_prod or {y: {"All": n} for y, n in conversions_by_year.items()}
        for y in range(start_year, last_year + 1):
            prin = intr = dev_paid = res_paid = capi_used = 0.0
            dev_prin = dev_int = res_prin = res_int = res_paying = 0.0
            payoff_prin = payoff_int = payoff_int_capi = xfer = 0.0
            res_used = dsrf_rel = 0.0
            issued_par = issued_proceeds = issued_units = 0.0
            for s in self.series:
                if s.dated_date.year >= y or s.outstanding <= 0:
                    continue
                kk = y - s.dated_date.year
                remaining = s.term_years - kk + 1
                i_y = s.outstanding * s.interest_rate
                # Homes paying installments this year: converted in a prior year
                # (resident_units is still last year-end here) and past the CAPI
                # period, during which their coupons come from the fund.
                _capi_thru = s.capi_end_date.year if (s.capi_years and s.capi_end_date) else None
                if _capi_thru is None or y > _capi_thru:
                    res_paying += s.resident_units
                fp_year = s.first_principal_date.year if s.first_principal_date else s.dated_date.year + 1
                if y < fp_year:
                    # Before the first principal date (the year after the CAPI period): interest only.
                    p_y = 0.0
                else:
                    pay = level_payment(s.outstanding, s.interest_rate, remaining)
                    p_y = min(max(pay - i_y, 0.0), s.outstanding)
                if remaining <= 1:
                    p_y = s.outstanding
                # Coupons through the end of the CAPI period (dated + CAPI years) are
                # capitalized -- three CAPI years from 8/15/26 cover every coupon
                # through 8/15/29 -- and principal begins the following year
                # (8/15/30).  The capitalized amount is sized to EXACTLY these
                # coupons on the actual balances (after any payoffs), so there is
                # never a surplus or a shortfall in the fund.
                from_fund = 0.0
                if s.capi_years and s.capi_end_date:
                    if _int_date(y, cfg) <= s.capi_end_date:
                        from_fund += i_y * 0.5
                    if _prin_date(y, cfg) <= s.capi_end_date:
                        from_fund += i_y * 0.5
                s.capi_interest_paid += from_fund
                owed = p_y + i_y - from_fund
                capi_used += from_fund
                # Final year: whatever remains in the debt service reserve fund is
                # applied to the last installment (pro rata to principal and interest).
                from_res = 0.0
                if remaining <= 1 and s.dsrf_balance > 0 and owed > 0:
                    from_res = min(s.dsrf_balance, owed)
                    s.dsrf_balance -= from_res
                    s.dsrf_applied_maturity += from_res
                    s.dsrf_applied_year = y
                    res_used += from_res
                _keep = (1.0 - from_res / owed) if owed > 0 else 1.0
                owed -= from_res
                # Who pays the rest of this installment: split by the balances held before it.
                dev_held = s.dev_units * s.dev_balance_per_unit
                share_dev = (dev_held / s.outstanding) if s.outstanding > 0 else 0.0
                dev_paid += owed * share_dev
                res_paid += owed * (1.0 - share_dev)
                dev_prin += p_y * _keep * share_dev
                res_prin += p_y * _keep * (1.0 - share_dev)
                dev_int += (i_y - from_fund) * _keep * share_dev
                res_int += (i_y - from_fund) * _keep * (1.0 - share_dev)
                f = (s.outstanding - p_y) / s.outstanding if s.outstanding > 0 else 0.0
                s.outstanding -= p_y
                for c in s.cohorts.values():          # every balance amortizes pro rata
                    c["dev_bal_pu"] *= f
                    c["res_bal"] *= f
                s.dev_balance_per_unit *= f
                s.resident_balance *= f
                self._sync(s)
                s.rows.append({"year": y, "principal": p_y, "interest": i_y})
                prin += p_y
                intr += i_y
            for s in self.series:
                if s.dated_date.year == y:
                    issued_par += s.par_amount
                    issued_proceeds += s.project_proceeds
                    issued_units += s.units
            # Conversions: lots delivered before the first bond convert first (no
            # bond on them); then FIFO against the oldest series still holding
            # lots.  The developer pays the lot's balance DOWN TO `resident_per_unit`
            # (plus accrued interest on the amount paid off); the homeowner assumes
            # exactly the resident amount (or the whole balance if it is already lower).
            n_total_unmatched = 0.0
            for p_, n in (conversions_by_year_prod.get(y, {}) or {}).items():
              if n <= 0:
                  continue
              if unfunded_pool_p.get(p_, 0.0) > 0:
                  take0 = min(n, unfunded_pool_p[p_])
                  unfunded_pool_p[p_] -= take0
                  unfunded_pool -= take0
                  n -= take0
                  self.unmatched_by_year[y] = self.unmatched_by_year.get(y, 0.0) + take0
              for s in self.series:
                if n <= 0:
                    break
                c = s.cohorts.get(p_)
                if s.dated_date.year > y or c is None or c["dev_units"] <= 0:
                    continue
                take = min(n, c["dev_units"])
                b = c["dev_bal_pu"]
                if c.get("installment", 0.0) > 0:
                    # Installment-driven: leave on the lot the principal that this
                    # level P+I retires over the homeowner's remaining principal
                    # years (first installment the year after conversion, or the
                    # first principal year if the bond is still in its CAPI period).
                    _fp = s.first_payment_date.year
                    _start = max(_fp, y + 1)
                    _n = s.final_maturity.year - _start + 1
                    res_target = level_balance(c["installment"], s.interest_rate, _n)
                else:
                    res_target = c["res_per_unit"]
                keep_pl = min(res_target, b)
                pay_pl = b - keep_pl
                amt = take * pay_pl
                acc = amt * s.interest_rate * PAYOFF_ACCRUAL_YEARS
                # The reserve requirement falls with the bonds redeemed: release the
                # pro rata share of the DSRF and credit it against the developer's payoff.
                rel = (s.dsrf_balance * amt / s.outstanding) if (s.outstanding > 0 and s.dsrf_balance > 0) else 0.0
                s.dsrf_balance -= rel
                s.dsrf_released_payoff += rel
                if rel:
                    s.dsrf_releases[y] = s.dsrf_releases.get(y, 0.0) + rel
                dsrf_rel += rel
                s.outstanding -= amt
                s.redemptions[y] = s.redemptions.get(y, 0.0) + amt
                s.transfers[y] = s.transfers.get(y, 0.0) + take * keep_pl
                s.payoff_principal += amt
                # The payoff is dated on the off-cycle coupon date after the
                # conversion year.  Inside the CAPI period that coupon (on the
                # full balance, including the amount paid off) is funded from
                # capitalized interest, so the developer pays principal only.
                if s.capi_years and s.capi_end_date and _int_date(y + 1, cfg) <= s.capi_end_date:
                    s.capi_interest_paid += acc
                    s.payoff_interest_capi += acc
                    payoff_int_capi += acc
                    acc = 0.0
                s.payoff_interest += acc
                c["res_bal"] += take * keep_pl
                c["res_units"] += take
                c["xfer_balance"] += take * keep_pl
                c["xfer_units"] += take
                c["xfer_unit_years"] += take * max(s.term_years - (y - s.dated_date.year), 1)
                c["dev_units"] -= take
                self._sync(s)
                payoff_prin += amt
                payoff_int += acc
                xfer += take * keep_pl
                n -= take
              n_total_unmatched += max(n, 0.0)
            n = n_total_unmatched
            unmatched += n
            if n > 0:
                self.unmatched_by_year[y] = self.unmatched_by_year.get(y, 0.0) + n
            for s in self.series:                # clear float noise once a series is retired
                if s.outstanding < 0.005:
                    s.outstanding = 0.0
                if s.resident_balance < 0.005:
                    s.resident_balance = 0.0
                if s.dev_units <= 0:
                    s.dev_units = 0.0
                for c in s.cohorts.values():
                    if c["dev_units"] <= 1e-9:
                        c["dev_units"] = 0.0
                    if c["res_bal"] < 0.005:
                        c["res_bal"] = 0.0
            outstanding = sum(s.outstanding for s in self.series)
            dev_held = sum(s.dev_units * s.dev_balance_per_unit for s in self.series)
            res_held = sum(s.resident_balance for s in self.series)
            row = {
                "year": y,
                "lots_delivered": units_by_year.get(y, 0.0),
                "lots_funded": issued_units,
                "par_issued": issued_par,
                "net_proceeds": issued_proceeds,
                "homes_converted": conversions_by_year.get(y, 0.0),
                "developer_payoff_principal": payoff_prin,
                "developer_payoff_interest": payoff_int,
                "payoff_interest_from_capi": payoff_int_capi,
                "dsrf_released_to_developer": dsrf_rel,
                "paid_from_reserve": res_used,
                "dsrf_balance_end": sum(s.dsrf_balance for s in self.series),
                "developer_payoff_total": payoff_prin + payoff_int,
                "balance_transferred_to_homeowners": xfer,
                "scheduled_principal": prin,
                "interest": intr,
                "total_installments": prin + intr,
                "paid_from_capitalized_interest": capi_used,
                "developer_paid_installments": dev_paid,
                "developer_paid_principal": dev_prin,
                "developer_paid_interest": dev_int,
                "homeowner_paid_installments": res_paid,
                "homeowner_paid_principal": res_prin,
                "homeowner_paid_interest": res_int,
                "outstanding_end": outstanding,
                "developer_held_balance": dev_held,
                "homeowner_held_balance": res_held,
                "developer_lots_outstanding": sum(s.dev_units for s in self.series),
                "homeowner_lots": sum(s.resident_units for s in self.series),
                "homeowner_lots_paying": res_paying,
            }
            rows.append({k_: (0.0 if isinstance(v, float) and abs(v) < 0.005 else v) for k_, v in row.items()})
            if (outstanding <= 0.5 and y >= years[-1]
                    and all(s.dated_date.year <= y for s in self.series)):
                break
        self.unmatched_closings = unmatched
        return rows

    # ── Totals ────────────────────────────────────────────────────────────
    @property
    def enabled(self) -> bool:
        return bool(self.series)

    def _col_sum(self, col: str) -> float:
        if self.annual.empty or col not in self.annual:
            return 0.0
        return float(self.annual[col].sum())

    @property
    def total_par(self) -> float:
        return sum(s.par_amount for s in self.series)

    @property
    def total_proceeds(self) -> float:
        return sum(s.project_proceeds for s in self.series)

    @property
    def total_coi(self) -> float:
        return sum(s.coi for s in self.series)

    @property
    def total_muniadv_fee(self) -> float:
        return sum(s.muniadv_fee for s in self.series)

    @property
    def total_bond_discount(self) -> float:
        return sum(s.bond_discount for s in self.series)

    @property
    def total_coi_pct_fee(self) -> float:
        return sum(s.coi_pct_fee for s in self.series)

    @property
    def total_reserve(self) -> float:
        return sum(s.reserve_fund for s in self.series)

    @property
    def total_dsrf_released(self) -> float:
        """Reserve released pro rata as lots pay off, credited to the developer."""
        return self._col_sum("dsrf_released_to_developer")

    @property
    def total_paid_from_reserve(self) -> float:
        """Reserve applied to the final installments at maturity."""
        return self._col_sum("paid_from_reserve")

    @property
    def total_capitalized_interest(self) -> float:
        return sum(s.capitalized_interest for s in self.series)

    @property
    def total_paid_from_capi(self) -> float:
        return self._col_sum("paid_from_capitalized_interest") + self._col_sum("payoff_interest_from_capi")

    @property
    def total_units(self) -> float:
        return sum(s.units for s in self.series)

    @property
    def total_lots(self) -> float:        # alias
        return self.total_units

    @property
    def total_payoff_principal(self) -> float:
        return self._col_sum("developer_payoff_principal")

    @property
    def total_payoff_interest(self) -> float:
        return self._col_sum("developer_payoff_interest")

    @property
    def total_payoff(self) -> float:
        return self.total_payoff_principal + self.total_payoff_interest

    @property
    def total_transferred(self) -> float:
        return self._col_sum("balance_transferred_to_homeowners")

    @property
    def total_interest(self) -> float:
        return (self._col_sum("interest") + self._col_sum("developer_payoff_interest")
                + self._col_sum("payoff_interest_from_capi"))

    @property
    def total_installments(self) -> float:
        return self._col_sum("total_installments")

    @property
    def total_developer_paid_installments(self) -> float:
        return self._col_sum("developer_paid_installments")

    @property
    def total_homeowner_paid_installments(self) -> float:
        return self._col_sum("homeowner_paid_installments")

    @property
    def net_developer_financing(self) -> float:
        """Proceeds less what the developer pays back (payoffs with accrued
        interest + installments on lots it holds)."""
        return (self.total_proceeds - self.total_payoff - self.total_developer_paid_installments
                + self.total_dsrf_released)

    @property
    def homes_transferred(self) -> float:
        return sum(s.xfer_units for s in self.series)

    @property
    def avg_transferred_balance(self) -> float:
        n = self.homes_transferred
        return (sum(s.xfer_balance for s in self.series) / n) if n else 0.0

    @property
    def avg_remaining_term(self) -> float:
        n = self.homes_transferred
        return (sum(s.xfer_unit_years for s in self.series) / n) if n else 0.0

    def product_transfer_stats(self) -> dict:
        """{product: {"units", "balance", "unit_years", "lots", "per_unit", "res_per_unit",
        "installment"}} aggregated across the series' cohorts; ``installment`` is the
        homes-weighted average of each cohort's level P+I on its average balance
        over its average remaining term."""
        out = {}
        for s in self.series:
            for p, c in s.cohorts.items():
                o = out.setdefault(p, {"units": 0.0, "balance": 0.0, "unit_years": 0.0, "lots": 0.0,
                                       "par": 0.0, "res_per_unit": c["res_per_unit"], "inst_x_units": 0.0})
                o["units"] += c["xfer_units"]; o["balance"] += c["xfer_balance"]
                o["unit_years"] += c["xfer_unit_years"]; o["lots"] += c["units"]
                o["par"] += c["units"] * c["per_unit"]
                if c["xfer_units"] > 0:
                    o["inst_x_units"] += c["xfer_units"] * level_payment(
                        c["xfer_balance"] / c["xfer_units"], s.interest_rate,
                        max(int(round(c["xfer_unit_years"] / c["xfer_units"])), 1))
        for o in out.values():
            o["per_unit"] = (o["par"] / o["lots"]) if o["lots"] else 0.0
            o["installment"] = (o["inst_x_units"] / o["units"]) if o["units"] else 0.0
        return out

    @property
    def required_value_to_lien(self) -> float:
        """Par-weighted required value-to-lien across the series (each series' requirement
        is the par-weighted LTV of the products it funds, from the Development Inputs table)."""
        par = self.total_par
        return (sum(s.required_value_to_lien * s.par_amount for s in self.series) / par) if par > 0 else 0.0

    @property
    def weighted_value_to_lien(self) -> float:
        par = self.total_par
        return (sum(s.unit_value for s in self.series) / par) if par > 0 else 0.0

    @property
    def all_meet_value_to_lien(self) -> bool:
        return all(s.meets_value_to_lien for s in self.series)

    @property
    def overleveraged_series(self) -> list:
        return [s for s in self.series if s.overleveraged]

    @property
    def max_assessment_per_lot_supported(self) -> float:
        """Lots-weighted assessment per lot the finished-lot value supports at the required ratio."""
        lots = sum(s.units for s in self.series)
        return (sum(s.max_par_at_required for s in self.series) / lots) if lots else 0.0

    @property
    def overleverage_excess(self) -> float:
        return sum(s.overleverage_excess for s in self.series)

    def series_dataframe(self) -> pd.DataFrame:
        rows = []
        for s in self.series:
            rows.append({
                "series": s.series_id,
                "dated_date": s.dated_date,
                "final_maturity": s.final_maturity,
                "delivery_years": (f"{s.window_start}-{s.window_end}" if s.window_end > s.window_start
                                   else f"{s.window_start}"),
                "lots_funded": s.units,
                "assessment_per_lot": s.lien_per_unit,
                "par_amount": s.par_amount,
                "lot_value": s.unit_value,
                "value_to_lien": s.value_to_lien,
                "value_to_lien_test": ("Max at " + f"{s.required_value_to_lien:g}:1" if self.max_to_ltv
                                       else ("OK" if s.meets_value_to_lien
                                             else f"OVER-LEVERAGED ({s.value_to_lien:.2f}:1 vs {s.required_value_to_lien:.3g}:1)")),
                "interest_rate": s.interest_rate,
                "muniadv_fee": s.muniadv_fee,
                "bond_discount": s.bond_discount,
                "coi_pct_fee": s.coi_pct_fee,
                "total_coi": s.coi,
                "reserve_fund": s.reserve_fund,
                "dsrf_basis": s.dsrf_basis,
                "dsrf_released_payoff": s.dsrf_released_payoff,
                "dsrf_applied_maturity": s.dsrf_applied_maturity,
                "capitalized_interest": s.capitalized_interest,
                "first_principal_date": s.first_payment_date,
                "net_proceeds": s.project_proceeds,
                "developer_payoff_principal": s.payoff_principal,
                "balance_transferred_to_homeowners": s.xfer_balance,
                "total_interest": s.total_interest,
            })
        return pd.DataFrame(rows)

    def summary(self) -> dict:
        return {
            "program": self.name,
            "per_lot": self.per_unit,
            "series": len(self.series),
            "lots_assessed": self.total_units,
            "total_par": self.total_par,
            "total_coi": self.total_coi,
            "total_reserve": self.total_reserve,
            "dsrf_released_to_developer": self.total_dsrf_released,
            "dsrf_applied_maturity": self.total_paid_from_reserve,
            "capitalized_interest": self.total_capitalized_interest,
            "paid_from_capitalized_interest": self.total_paid_from_capi,
            "capi_years": (self.series[0].capi_years if self.series else 0),
            "total_proceeds": self.total_proceeds,
            "developer_payoff_principal": self.total_payoff_principal,
            "developer_payoff_interest": self.total_payoff_interest,
            "developer_payoff_total": self.total_payoff,
            "balance_transferred_to_homeowners": self.total_transferred,
            "homes_transferred": self.homes_transferred,
            "avg_transferred_balance": self.avg_transferred_balance,
            "avg_remaining_term": self.avg_remaining_term,
            "developer_paid_installments": self.total_developer_paid_installments,
            "homeowner_paid_installments": self.total_homeowner_paid_installments,
            "net_developer_financing": self.net_developer_financing,
            "total_interest": self.total_interest,
            "weighted_value_to_lien": self.weighted_value_to_lien,
            "all_meet_value_to_lien": self.all_meet_value_to_lien,
            "unmatched_closings": self.unmatched_closings,
            "lots_not_issued": self.units_not_issued,
            "lots_unfunded_early": self.units_unfunded_early,
            "first_fundable_year": self.first_fundable_year,
        }


class AssessmentBondAnalysis:
    """
    Sizes the special assessment program (only constructed when
    Inputs!ASSESSMENT_BONDS = Yes).

    Usage:
        asmt = AssessmentBondAnalysis(cfg, summary_model, go_dated_dates=[...]).run()
        asmt.developer (AssessmentProgram), asmt.summary(), asmt.effective_tax_rate_rows()
    """

    def __init__(self, config: ModelConfig, summary_model: SummaryModel, go_dated_dates=None):
        self.cfg = config
        self.summary_model = summary_model
        # Assessment amounts EXACTLY as entered on the Inputs page.  A blank / zero
        # DEVELOPER_BOND with ASSESSMENT_BONDS = Yes means "size to the maximum the
        # finished-lot value supports at the Loan-to-Value ratio"; a blank / zero
        # RESIDENT_BOND means nothing stays with the homeowner (the developer pays
        # off each lot's whole balance at conversion).
        self.developer_per_lot = float(getattr(config, "developer_bond", 0.0) or 0.0)
        self.resident_per_lot = float(getattr(config, "resident_bond", 0.0) or 0.0)
        self.max_to_ltv = self.developer_per_lot <= 0
        self.go_dated_dates = list(go_dated_dates or [])
        # Per-product amounts (Development Inputs 'Special Assessment' table),
        # falling back to the Inputs-tab cells for products without an entry.
        self.prod_params = self._product_params(config, summary_model)
        if self.prod_params:
            self.max_to_ltv = any(v.get("max_at_ltv") for v in self.prod_params.values())
        self.developer = AssessmentProgram(DEVELOPER, config, self.developer_per_lot, self.go_dated_dates)
        self._built = False

    @staticmethod
    def _product_params(cfg, summary_model) -> dict:
        """Resolve {product: {"developer", "resident", "ltv", "installment"}} from the
        Development Inputs table (cfg.assessment_products), falling back to the
        Inputs cells.  A blank resident amount with a Resident Annual Installment
        Payment entered is backed out as the balance that level payment retires
        over the homeowner's term."""
        table = dict(getattr(cfg, "assessment_products", {}) or {})
        try:
            products = [p.name for p in getattr(summary_model.dev, "products", []) or []]
        except Exception:
            products = []
        if not table and not products:
            return {}
        fb_dev = float(getattr(cfg, "developer_bond", 0.0) or 0.0) or None
        fb_res = float(getattr(cfg, "resident_bond", 0.0) or 0.0) or None
        fb_ltv = float(getattr(cfg, "value_to_lien", 3.0) or 3.0)
        rate = float(getattr(cfg, "assessment_rate", 0.07) or 0.0)
        term = int(getattr(cfg, "assessment_term_yrs", 25) or 25)
        capi = int(round(float(getattr(cfg, "assessment_capi_years", 0.0) or 0.0)))
        lead = int(round(float(getattr(cfg, "assessment_fund_timing", 1.0) or 0.0)))
        n_pay = max(term - max(capi, lead + 1), 1)     # homeowner's principal years on a first-cohort lot
        out = {}
        n_prin = max(term - capi, 1)                    # principal years from the first principal date to maturity
        for p in (products or list(table)):
            e = table.get(p, {}) or {}
            dev = e.get("developer") if e.get("developer") not in (None, 0, 0.0) else fb_dev
            res = e.get("resident") if e.get("resident") not in (None, 0, 0.0) else None
            inst = float(e.get("installment") or 0.0)
            if inst > 0:
                # The installment governs: the balance left on each lot is SOLVED at
                # conversion so the homeowner's level P+I equals this amount every
                # year to final maturity (AssessmentProgram does this per cohort).
                # The figure here is the first-cohort estimate for reporting only.
                res = level_balance(inst, rate, n_pay)
            if res is None:
                res = fb_res
            resident_only = False
            if not dev and (res or inst > 0):
                # Developer amount blank = no developer assessment: the bond is sized
                # only to what the homeowner carries — the entered resident amount, or
                # the principal the installment retires over the full principal term.
                # Level amortisation then leaves exactly the solved balance on every
                # lot at conversion, so nothing is paid off by the developer.
                dev = level_balance(inst, rate, n_prin) if inst > 0 else res
                resident_only = True
            ltv = e.get("ltv") or fb_ltv
            # With the per-product table present a blank is $0: developer, resident
            # and installment all blank = no assessment on that product.  Sizing to
            # the maximum at the LTV applies only to legacy workbooks that still use
            # the Inputs-tab cells (no table) with the developer cell blank.
            max_at_ltv = (not table) and not dev
            none = bool(table) and not dev and not res and inst <= 0
            out[p] = {"developer": dev, "resident": res or 0.0, "ltv": ltv, "installment": inst or None,
                      "installment_driven": inst > 0, "resident_only": resident_only,
                      "max_at_ltv": max_at_ltv, "none": none, "from_table": bool(e)}
        return out

    @property
    def multi_product(self) -> bool:
        """True when the per-product amounts actually differ between products."""
        vals = {(round(float(v.get("developer") or 0)), round(float(v.get("resident") or 0)))
                for v in self.prod_params.values()}
        return len(vals) > 1

    @property
    def product_amounts(self) -> dict:
        return self.prod_params

    @property
    def sizing_mode(self) -> str:
        pp = self.prod_params or {}
        blank = [p for p, v in pp.items() if v.get("max_at_ltv")]
        res_only = [p for p, v in pp.items() if v.get("resident_only")]
        none = [p for p, v in pp.items() if v.get("none")]
        parts = []
        if none:
            parts.append(", ".join(none) + (" carry" if len(none) > 1 else " carries") + " no assessment (all amounts blank)")
        if res_only:
            _ro = [p + (" (solved from the annual installment)" if pp[p].get("installment_driven") else "") for p in res_only]
            parts.append(", ".join(_ro) + (" carry" if len(res_only) > 1 else " carries")
                         + " a resident-only assessment (developer amount blank; the bond is sized to the resident amount)")
        if blank:
            parts.append(", ".join(blank) + " sized to the maximum at its LTV (all amounts blank)")
        if pp and parts and (len(blank) + len(res_only) + len(none)) < len(pp):
            return "Developer Assessment amounts by product (Development Inputs tab); " + "; ".join(parts)
        if pp and none and len(none) == len(pp):
            return "No assessment: every product's developer, resident and installment cells are blank"
        if pp and res_only and not blank:
            return "Resident-only assessments (developer amount blank): " + parts[0]
        if self.max_to_ltv:
            return f"Maximum at {float(getattr(self.cfg, 'value_to_lien', 3.0) or 3.0):g}:1 value-to-lien (Developer Assessment cell blank)"
        if self.multi_product:
            return "Developer Assessment amounts as entered by product (Development Inputs tab)"
        return "Developer Assessment amount as entered"

    @property
    def total_per_lot(self) -> float:
        return self.developer_per_lot

    @property
    def developer_payoff_per_lot(self) -> float:
        """Nominal developer payoff per lot (developer − resident amount) — the amount on
        an unamortized lot; lots-weighted across products when amounts differ."""
        st = self.developer.product_transfer_stats() if self.developer.enabled else {}
        if st and self.prod_params:
            tot = sum(o["lots"] for o in st.values())
            if tot > 0:
                return sum(max(o["per_unit"] - o["res_per_unit"], 0.0) * o["lots"] for o in st.values()) / tot
        return max(self.developer_per_lot - self.resident_per_lot, 0.0)

    @property
    def programs(self) -> list[AssessmentProgram]:
        return [self.developer] if self.developer.enabled else []

    # ── Public ────────────────────────────────────────────────────────────
    def run(self) -> "AssessmentBondAnalysis":
        cfg = self.cfg
        dev = self.summary_model.dev
        (lots_by_year, lotval_by_year, homes_by_year,
         lots_p, lotval_p, homes_p) = _development_feed(dev, cfg)
        if not self.prod_params:
            self.prod_params = {p: {"developer": (None if self.max_to_ltv else self.developer_per_lot),
                                    "resident": self.resident_per_lot,
                                    "ltv": float(getattr(cfg, "value_to_lien", 3.0) or 3.0)}
                                for p in {p for by in lots_p.values() for p in by}}
        self.developer.run(lots_by_year, lotval_by_year, homes_by_year,
                           resident_per_unit=self.resident_per_lot, max_to_ltv=self.max_to_ltv,
                           window_years=int(getattr(cfg, "assessment_fund_term", 1) or 1),
                           lead_years=float(getattr(cfg, "assessment_fund_timing", 0.0) or 0.0),
                           units_by_year_prod=lots_p, value_by_year_prod=lotval_p,
                           conversions_by_year_prod=homes_p, prod_params=self.prod_params)
        if self.developer.total_units > 0:
            # Program-wide (lots-weighted) assessment and resident amounts for the
            # headline figures; the per-product detail is in effective_tax_rate_rows().
            self.developer_per_lot = self.developer.total_par / self.developer.total_units
            self.developer.per_unit = self.developer_per_lot
            st = self.developer.product_transfer_stats()
            tot = sum(o["lots"] for o in st.values())
            if tot > 0:
                self.resident_per_lot = sum(o["res_per_unit"] * o["lots"] for o in st.values()) / tot
        self._built = True
        return self

    # ── Aggregates ────────────────────────────────────────────────────────
    @property
    def enabled(self) -> bool:
        return bool(self.programs)

    @property
    def series(self) -> list[AssessmentSeries]:
        return list(self.developer.series)

    @property
    def total_par(self) -> float:
        return self.developer.total_par

    @property
    def total_proceeds(self) -> float:
        return self.developer.total_proceeds

    @property
    def total_coi(self) -> float:
        return self.developer.total_coi

    @property
    def total_reserve(self) -> float:
        return self.developer.total_reserve

    @property
    def total_interest(self) -> float:
        return self.developer.total_interest

    @property
    def total_payoff(self) -> float:
        return self.developer.total_payoff

    @property
    def total_lots(self) -> float:
        return self.developer.total_units

    @property
    def net_developer_financing(self) -> float:
        return self.developer.net_developer_financing

    @property
    def weighted_value_to_lien(self) -> float:
        return self.developer.weighted_value_to_lien

    @property
    def all_meet_value_to_lien(self) -> bool:
        return self.developer.all_meet_value_to_lien

    def resident_annual_installment(self) -> float:
        """Level annual installment a homeowner pays on the balance transferred
        at conversion: the average transferred balance over the average
        remaining term (falls back to RESIDENT_BOND over the full term before
        any conversion has occurred)."""
        cfg = self.cfg
        rate = float(getattr(cfg, "assessment_rate", 0.0) or 0.0)
        d = self.developer
        if d.homes_transferred > 0:
            st = d.product_transfer_stats()
            n = sum(o["units"] for o in st.values())
            if n > 0:
                return sum(o["installment"] * o["units"] for o in st.values()) / n
            return level_payment(d.avg_transferred_balance, rate, max(int(round(d.avg_remaining_term)), 1))
        return level_payment(self.resident_per_lot, rate, int(getattr(cfg, "assessment_term_yrs", 25) or 25))

    @property
    def resident_balance_basis(self) -> float:
        d = self.developer
        return d.avg_transferred_balance if d.homes_transferred > 0 else self.resident_per_lot

    def effective_tax_rate_rows(self) -> list[dict]:
        """
        Effective tax rate of the homeowner's assessment: the annual
        installment on the balance transferred at conversion ÷ home value
        (starting ASP), per product and weighted by homes — shown as a % of
        home value and as $ per $100 of home value.
        """
        dev = self.summary_model.dev
        prices = {p.name: p.asp_base for p in (getattr(dev, "products", []) or [])}
        inst = self.resident_annual_installment()
        res = self.resident_balance_basis
        cfg = self.cfg
        rate = float(getattr(cfg, "assessment_rate", 0.0) or 0.0)
        term = int(getattr(cfg, "assessment_term_yrs", 25) or 25)
        stats = self.developer.product_transfer_stats() if self.developer.enabled else {}
        units = {}
        values = {}          # Σ closing value by product — fallback home value when no base ASP
        _infl = float(getattr(cfg, "inflation_rate", 0.0) or 0.0)
        _infl_start = int(getattr(cfg, "inflation_start_year", 0) or 0)
        for p in (getattr(dev, "products", []) or []):
            for y, n in (p.home_closings or {}).items():
                if not n:
                    continue
                units[p.name] = units.get(p.name, 0.0) + float(n)
                values[p.name] = values.get(p.name, 0.0) + float(n) * p.asp(
                    int(y), _infl, _infl_start)
        rows = []
        w_units = w_asp = 0.0
        for prod in [p.name for p in (getattr(dev, "products", []) or [])] or list(set(prices) | set(units)):
            n = float(units.get(prod, 0.0) or 0.0)
            asp = float(prices.get(prod, 0.0) or 0.0)
            if asp <= 0 and n > 0 and values.get(prod, 0.0) > 0:
                asp = values[prod] / n              # average closing value per home (escalated)
            if asp <= 0 or n <= 0:
                continue
            # Per-product balance carried and installment (from that product's
            # cohorts); falls back to the program-wide figures.
            st = stats.get(prod)
            pp = (self.prod_params or {}).get(prod, {})
            if st and st["units"] > 0:
                res_p = st["balance"] / st["units"]
                inst_p = st["installment"]
            elif pp:
                res_p = float(pp.get("resident") or 0.0)
                inst_p = level_payment(res_p, rate, term)
            else:
                res_p, inst_p = res, inst
            rows.append({
                "product": prod, "homes": n, "home_price": asp,
                "developer_assessment": (st["per_unit"] if st else float(pp.get("developer") or 0.0)),
                "resident_assessment": res_p, "annual_installment": inst_p,
                "effective_tax_rate": inst_p / asp,
                "rate_per_100_of_value": inst_p / asp * 100.0,
            })
            w_units += n; w_asp += asp * n
        _ar = [r_ for r_ in rows if (r_["developer_assessment"] > 0 or r_["resident_assessment"] > 0)] or rows
        w_units = sum(r_["homes"] for r_ in _ar); w_asp = sum(r_["home_price"] * r_["homes"] for r_ in _ar)
        rows_all, rows = rows, _ar
        if w_units > 0 and rows:
            asp = w_asp / w_units
            w_inst = sum(r_["annual_installment"] * r_["homes"] for r_ in rows) / w_units
            w_res = sum(r_["resident_assessment"] * r_["homes"] for r_ in rows) / w_units
            w_dev = sum(r_["developer_assessment"] * r_["homes"] for r_ in rows) / w_units
            rows = rows_all
            rows.append({
                "product": "Weighted Average (homes carrying an assessment)", "homes": w_units, "home_price": asp,
                "developer_assessment": w_dev,
                "resident_assessment": w_res, "annual_installment": w_inst,
                "effective_tax_rate": w_inst / asp,
                "rate_per_100_of_value": w_inst / asp * 100.0,
            })
        return rows

    def effective_tax_rate_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame(self.effective_tax_rate_rows())

    @property
    def effective_tax_rate(self) -> float:
        rows = self.effective_tax_rate_rows()
        return float(rows[-1]["effective_tax_rate"]) if rows else 0.0

    def overlap_rows(self) -> list[dict]:
        """
        Year-by-year overlap: for each series its par issued, principal paid off
        at conversion, scheduled principal and year-end outstanding; program
        totals (issued, paid off, net new money, outstanding); the split of the
        outstanding between developer-held and homeowner-held balances; and the
        developer's cash view (net proceeds vs payoffs and installments).
        """
        dev = self.developer
        if not dev.series or dev.annual.empty:
            return []
        ann = dev.annual.set_index("year")
        bal = {s.series_id: 0.0 for s in dev.series}
        out = []
        cum = 0.0
        for y in ann.index:
            row = {"year": int(y), "lots_delivered": float(ann.at[y, "lots_delivered"]),
                   "homes_converted": float(ann.at[y, "homes_converted"]), "series": {}}
            tot_iss = tot_ret = tot_sched = 0.0
            for s in dev.series:
                iss = s.par_amount if s.dated_date.year == y else 0.0
                sched = sum(r["principal"] for r in s.rows if r["year"] == y)
                ret = s.redemptions.get(y, 0.0)
                bal[s.series_id] += iss - sched - ret
                if bal[s.series_id] < 0.005:
                    bal[s.series_id] = 0.0
                row["series"][s.series_id] = {"issued": iss, "retired": ret, "scheduled": sched,
                                              "outstanding": bal[s.series_id]}
                tot_iss += iss; tot_ret += ret; tot_sched += sched
            row.update({
                "dev_issued": tot_iss, "dev_retired": tot_ret, "dev_scheduled": tot_sched,
                "dev_net_new": tot_iss - tot_ret - tot_sched,
                "developer_held": float(ann.at[y, "developer_held_balance"]),
                "homeowner_held": float(ann.at[y, "homeowner_held_balance"]),
                "transferred": float(ann.at[y, "balance_transferred_to_homeowners"]),
                "total_outstanding": float(ann.at[y, "outstanding_end"]),
                "cash_proceeds": float(ann.at[y, "net_proceeds"]),
                "cash_payoff": float(ann.at[y, "developer_payoff_total"]),
                "cash_installments": float(ann.at[y, "developer_paid_installments"]),
            })
            row["cash_net"] = row["cash_proceeds"] - row["cash_payoff"] - row["cash_installments"]
            cum += row["cash_net"]
            row["cash_cumulative"] = cum
            out.append(row)
        return out

    def summary(self) -> dict:
        d = self.developer.summary()
        return {
            "developer_per_lot": self.developer_per_lot,
            "resident_per_lot": self.resident_per_lot,
            "developer_payoff_per_lot": self.developer_payoff_per_lot,
            "max_to_ltv": self.max_to_ltv,
            "multi_product": self.multi_product,
            "product_amounts": self.prod_params,
            "sizing_mode": self.sizing_mode,
            "total_per_lot": self.total_per_lot,
            "lots_assessed": self.total_lots,
            "homes_transferred": d["homes_transferred"],
            "total_par": self.total_par,
            "total_coi": self.total_coi,
            "total_reserve": self.total_reserve,
            "dsrf_released_to_developer": d["dsrf_released_to_developer"],
            "dsrf_applied_maturity": d["dsrf_applied_maturity"],
            "capitalized_interest": self.developer.total_capitalized_interest,
            "capi_years": d["capi_years"],
            "total_proceeds": self.total_proceeds,
            "total_interest": self.total_interest,
            "developer_payoff_total": self.total_payoff,
            "net_developer_financing": self.net_developer_financing,
            "weighted_value_to_lien": self.weighted_value_to_lien,
            "overleveraged_series": [s.series_id for s in self.developer.overleveraged_series],
            "overleveraged_ratio": (sum(s.unit_value for s in self.developer.overleveraged_series)
                                    / max(sum(s.par_amount for s in self.developer.overleveraged_series), 1e-9)
                                    if self.developer.overleveraged_series else 0.0),
            "overleveraged_ratios": {s.series_id: s.value_to_lien for s in self.developer.overleveraged_series},
            "max_assessment_per_lot_supported": self.developer.max_assessment_per_lot_supported,
            "overleverage_excess": self.developer.overleverage_excess,
            "required_value_to_lien": self.developer.required_value_to_lien,
            "all_meet_value_to_lien": self.all_meet_value_to_lien,
            "effective_tax_rate": self.effective_tax_rate,
            "resident_annual_installment": self.resident_annual_installment(),
            "resident_balance_basis": self.resident_balance_basis,
            "developer": d,
        }


# ── Semiannual debt-service schedules (workbook 'Assessment Debt Service Per Lot') ──
def _development_feed(dev, cfg):
    """
    Lots delivered, their finished-lot value and homes closed — by year, and by
    year and product — read straight off ``DeveloperProjections``.

    The lot value is the same figure the ad valorem model puts on a delivered
    lot: the product's ASP inflated to the delivery year, times the platted-lot
    percentage.  It is a MARKET value, which is what a value-to-lien test wants;
    the 45% residential exemption reduces what the lot is TAXED on, not what it
    is worth to a lienholder.
    """
    lots_by_year: dict[int, float] = {}
    lotval_by_year: dict[int, float] = {}
    homes_by_year: dict[int, float] = {}
    lots_p: dict[int, dict] = {}
    lotval_p: dict[int, dict] = {}
    homes_p: dict[int, dict] = {}
    infl = float(getattr(cfg, "inflation_rate", 0.0) or 0.0)
    infl_start = int(getattr(cfg, "inflation_start_year", 0) or 0)
    platted = float(getattr(cfg, "platted_lot_value", 0.10) or 0.10)
    for prod in (getattr(dev, "products", []) or []):
        for y, n in (prod.lot_deliveries or {}).items():
            if not n:
                continue
            y = int(y)
            val = float(n) * prod.lot_value(y, infl, platted, infl_start)
            lots_by_year[y] = lots_by_year.get(y, 0.0) + float(n)
            lotval_by_year[y] = lotval_by_year.get(y, 0.0) + val
            by_u = lots_p.setdefault(y, {})
            by_v = lotval_p.setdefault(y, {})
            by_u[prod.name] = by_u.get(prod.name, 0.0) + float(n)
            by_v[prod.name] = by_v.get(prod.name, 0.0) + val
        for y, n in (prod.home_closings or {}).items():
            if not n:
                continue
            y = int(y)
            homes_by_year[y] = homes_by_year.get(y, 0.0) + float(n)
            by_h = homes_p.setdefault(y, {})
            by_h[prod.name] = by_h.get(prod.name, 0.0) + float(n)
    return lots_by_year, lotval_by_year, homes_by_year, lots_p, lotval_p, homes_p


def _int_date(year: int, cfg: ModelConfig) -> date:
    return _with_prin_day(date(year, cfg.int_maturity, 1), cfg)


def _prin_date(year: int, cfg: ModelConfig) -> date:
    return _with_prin_day(date(year, cfg.prin_maturity, 1), cfg)


def per_unit_schedule(amount: float, rate: float, term: int, dated: date, cfg: ModelConfig,
                      capi_years: int = 0, dsrf_ratio: float = 0.0) -> list[dict]:
    """``dsrf_ratio`` = the series' reserve ÷ par: this lot's share of the DSRF is
    ratio × balance, applied to its final installment (released at payoff in the
    conversion schedule)."""
    fp = assessment_first_principal_date(dated, capi_years, cfg) if capi_years else None
    ce = capi_end_date(dated, capi_years, cfg)
    """
    Level-installment schedule for ONE assessment held to maturity, laid out
    semiannually: interest on the off-cycle date and on the principal date
    (rate/2 × balance each), principal annually on the principal date so the
    two payments in each year sum to the level annual installment.  Row 0 is
    the dated date / opening balance.
    """
    rows = [{"no": "", "date": dated, "rate": None, "principal": None, "interest": None,
             "capi": None, "redemption": None, "installment": None, "balance": amount, "payer": ""}]
    bal = amount
    n = 0
    for k in range(1, term + 1):
        y = dated.year + k
        d_int, d_prin = _int_date(y, cfg), _prin_date(y, cfg)
        capi_int = bool(ce) and d_int <= ce          # coupon capitalized (through the end of the CAPI period)
        capi_prin_int = bool(ce) and d_prin <= ce    # principal-date coupon capitalized
        capi_prin = bool(fp) and d_prin < fp         # no principal until the year after the CAPI period
        level = level_payment(bal, rate, term - k + 1) if not capi_prin else 0.0
        i_half = bal * rate / 2.0
        n += 1
        rows.append({"no": n, "date": d_int, "rate": rate, "principal": 0.0,
                     "interest": i_half, "capi": i_half if capi_int else 0.0,
                     "redemption": 0.0, "installment": i_half, "balance": bal,
                     "payer": "Capitalized interest" if capi_int else "Lot owner"})
        prin = 0.0 if capi_prin else min(max(level - 2.0 * i_half, 0.0), bal)
        if k == term:
            prin = bal
        dsrf = min(dsrf_ratio * bal, prin + i_half) if (k == term and dsrf_ratio > 0) else 0.0
        n += 1
        bal -= prin
        rows.append({"no": n, "date": d_prin, "rate": rate, "principal": prin,
                     "interest": i_half, "capi": i_half if capi_prin_int else 0.0,
                     "redemption": 0.0, "dsrf": dsrf, "installment": prin + i_half, "balance": bal,
                     "payer": ("Capitalized interest" if capi_prin
                               else ("Lot owner (principal); interest capitalized" if capi_prin_int else "Lot owner"))
                              + ("; DSRF applied" if dsrf else "")})
    return rows


def per_lot_conversion_schedule(amount: float, resident_amount: float, rate: float, term: int,
                                dated: date, convert_year_offset: int, cfg: ModelConfig,
                                capi_years: int = 0, dsrf_ratio: float = 0.0) -> list[dict]:
    """
    Illustrative schedule for ONE lot that converts ``convert_year_offset``
    years after the bond is dated: level installments on the full assessment
    until conversion; on the off-cycle date after the conversion year the
    developer pays off ``payoff`` of principal (with the half-year of interest
    accrued on the whole balance); the remaining balance then re-levels over
    the remaining term and is paid by the homeowner.  Each row carries the
    payer (Developer / Homeowner).
    """
    rows = [{"no": "", "date": dated, "rate": None, "principal": None, "interest": None,
             "capi": None, "redemption": None, "installment": None, "balance": amount, "payer": ""}]
    fp = assessment_first_principal_date(dated, capi_years, cfg) if capi_years else None
    ce = capi_end_date(dated, capi_years, cfg)
    bal = amount
    n = 0
    converted = False
    for k in range(1, term + 1):
        y = dated.year + k
        if bal <= 0.005:
            break
        in_capi = bool(ce) and _int_date(y, cfg) <= ce          # off-cycle coupon capitalized (through the end of the CAPI period)?
        capi_prin_int = bool(ce) and _prin_date(y, cfg) <= ce   # principal-date coupon capitalized?
        capi_prin = bool(fp) and _prin_date(y, cfg) < fp        # no principal until the year after the CAPI period
        red = 0.0
        if not converted and k == convert_year_offset + 1 and convert_year_offset >= 0:
            red = max(bal - resident_amount, 0.0)
        i_half = bal * rate / 2.0
        n += 1
        bal -= red
        payer = "Capitalized interest" if in_capi else ("Homeowner" if converted else "Developer")
        if red:
            converted = True
            payer = ("Developer (payoff of principal); interest capitalized" if in_capi else "Developer (payoff)")
        # Inside the CAPI period the whole coupon (including the half year accrued
        # on the amount paid off) comes from the fund: the developer pays principal only.
        # The lot's pro rata share of the reserve (ratio × amount paid off) is
        # released with the payoff and credited to the developer.
        rows.append({"no": n, "date": _int_date(y, cfg), "rate": rate, "principal": 0.0,
                     "interest": i_half, "capi": i_half if in_capi else 0.0,
                     "redemption": red, "dsrf": red * dsrf_ratio, "installment": i_half + red, "balance": bal,
                     "payer": payer + ("; DSRF released" if red and dsrf_ratio > 0 else "")})
        if bal <= 0.005:
            break
        level = level_payment(bal, rate, term - k + 1) if not capi_prin else 0.0
        i_half = bal * rate / 2.0
        prin = 0.0 if capi_prin else min(max(level - 2.0 * i_half, 0.0), bal)
        if k == term:
            prin = bal
        dsrf = min(dsrf_ratio * bal, prin + i_half) if (k == term and dsrf_ratio > 0) else 0.0
        n += 1
        bal -= prin
        rows.append({"no": n, "date": _prin_date(y, cfg), "rate": rate, "principal": prin,
                     "interest": i_half, "capi": i_half if capi_prin_int else 0.0,
                     "redemption": 0.0, "dsrf": dsrf, "installment": prin + i_half, "balance": bal,
                     "payer": ("Capitalized interest" if capi_prin
                               else (("Homeowner" if converted else "Developer")
                                     + ("; interest capitalized" if capi_prin_int else "")))
                              + ("; DSRF applied" if dsrf else "")})
    return rows


def series_schedule(series: AssessmentSeries, cfg: ModelConfig) -> list[dict]:
    """
    Semiannual debt service for one series, consistent with the annual program
    cashflow: the year's scheduled principal falls on the principal date,
    interest is split across the two payment dates, and the principal paid off
    at conversion is redeemed on the off-cycle interest date following the
    conversion year (with the half-year of interest accrued on it).
    """
    rate = series.interest_rate
    rows = [{"no": "", "date": series.dated_date, "rate": None, "principal": None, "interest": None,
             "capi": None, "redemption": None, "installment": None, "balance": series.par_amount}]
    bal = series.par_amount
    ce = series.capi_end_date if series.capi_years else None
    by_year = {r["year"]: r for r in series.rows}
    last_year = max([series.final_maturity.year] + list(series.redemptions.keys()) + list(by_year.keys()))
    n = 0
    for y in range(series.dated_date.year + 1, last_year + 2):
        if bal <= 0.005:
            break
        red = min(series.redemptions.get(y - 1, 0.0), bal)
        rel = series.dsrf_releases.get(y - 1, 0.0) if red else 0.0   # reserve released with this redemption
        # Remaining reserve applied in the final year: split across that year's two
        # coupons (first the off-cycle coupon, the rest on the principal date).
        applied_total = series.dsrf_applied_maturity if (series.dsrf_applied_year == y) else 0.0
        i_half = bal * rate / 2.0
        a1 = min(applied_total, i_half)
        n += 1
        bal -= red
        d_ = _int_date(y, cfg)
        rows.append({"no": n, "date": d_, "rate": rate, "principal": 0.0,
                     "interest": i_half, "capi": i_half if (ce and d_ <= ce) else 0.0,
                     "redemption": red, "dsrf": rel + a1, "installment": i_half + red, "balance": bal})
        if bal <= 0.005:
            break
        prin = min(by_year.get(y, {}).get("principal", 0.0), bal)
        if y >= series.final_maturity.year:
            prin = bal
        i_half = bal * rate / 2.0
        applied = applied_total - a1
        n += 1
        bal -= prin
        d_ = _prin_date(y, cfg)
        rows.append({"no": n, "date": d_, "rate": rate, "principal": prin,
                     "interest": i_half, "capi": i_half if (ce and d_ <= ce) else 0.0,
                     "redemption": 0.0, "dsrf": applied, "installment": prin + i_half, "balance": bal})
    return rows
