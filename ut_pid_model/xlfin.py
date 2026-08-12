"""
xlfin.py — Excel-compatible date and bond-math primitives.

The Colorado workbook this model mirrors leans heavily on a handful of Excel
functions whose conventions are not the same as Python's.  Getting these exactly
right is what makes the Python output tie to the workbook to the penny, so they
live here on their own rather than being inlined at the call sites.

Implemented:
    EDATE, YEARFRAC (basis 0), DAYS360 (US/NASD), PRICE (basis 0, freq 2),
    plus an IRR-style solver used for TIC / All-in TIC.
"""

from __future__ import annotations

from datetime import date
from typing import Iterable, Sequence


# ── Date helpers ──────────────────────────────────────────────────────────────

def edate(d: date, months: int) -> date:
    """Excel EDATE: d shifted by `months`, clamped to the end of the month."""
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    day = min(d.day, _days_in_month(y, m))
    return date(y, m, day)


def _days_in_month(year: int, month: int) -> int:
    if month == 2:
        leap = (year % 4 == 0 and year % 100 != 0) or year % 400 == 0
        return 29 if leap else 28
    return 30 if month in (4, 6, 9, 11) else 31


def days360(start: date, end: date, european: bool = False) -> int:
    """Excel DAYS360 with the US (NASD) convention by default."""
    d1, m1, y1 = start.day, start.month, start.year
    d2, m2, y2 = end.day, end.month, end.year

    if european:
        d1 = min(d1, 30)
        d2 = min(d2, 30)
    else:
        if d1 == 31 or (m1 == 2 and d1 == _days_in_month(y1, 2)):
            d1 = 30
        if d2 == 31:
            d2 = 30 if d1 == 30 else 31
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)


def yearfrac(start: date, end: date) -> float:
    """Excel YEARFRAC with basis 0 (30/360 US) — the workbook's default."""
    if start == end:
        return 0.0
    sign = 1.0
    if start > end:
        start, end = end, start
        sign = -1.0
    return sign * days360(start, end) / 360.0


# ── Bond math ─────────────────────────────────────────────────────────────────

def price(settlement: date, maturity: date, rate: float, yld: float,
          redemption: float = 100.0, frequency: int = 2) -> float:
    """
    Excel PRICE with basis 0 (30/360).  Returns price per 100 of face.

    Used by the debt-service tabs to derive premium/OID on each maturity.
    """
    if yld <= -1.0:
        raise ValueError("yield must exceed -100%")

    months = 12 // frequency
    # Walk coupon dates backwards from maturity to the first one after settlement.
    coupons: list[date] = []
    d = maturity
    while d > settlement:
        coupons.append(d)
        d = edate(d, -months)
    if not coupons:
        return redemption
    coupons.reverse()
    prev_coupon = edate(coupons[0], -months)

    e = 360.0 / frequency                       # days in the coupon period
    dsc = days360(settlement, coupons[0])       # settlement → next coupon
    a = days360(prev_coupon, settlement)        # accrued days
    n = len(coupons)

    coupon = 100.0 * rate / frequency
    y = yld / frequency

    value = redemption / (1.0 + y) ** (n - 1 + dsc / e)
    for k in range(1, n + 1):
        value += coupon / (1.0 + y) ** (k - 1 + dsc / e)
    value -= coupon * a / e
    return value


def truncate(value: float, digits: int) -> float:
    """Excel TRUNC — toward zero, no rounding."""
    factor = 10.0 ** digits
    return int(value * factor) / factor


def npv(rate: float, cashflows: Sequence[float]) -> float:
    """Excel NPV: first cashflow is discounted one full period."""
    return sum(cf / (1.0 + rate) ** (i + 1) for i, cf in enumerate(cashflows))


def solve_tic(cashflows: Sequence[float], proceeds: float, offset_periods: float,
              periods_per_year: int = 2,
              lo: float = -0.5, hi: float = 1.0, tol: float = 1e-12) -> float:
    """
    Solve for the nominal annual rate `r` such that

        NPV(r / f, cashflows) * (1 + r / f) ** offset_periods == proceeds

    which is the construction the workbook uses for the arbitrage TIC and the
    all-in TIC (the two differ only in what `proceeds` nets out).

    `offset_periods` is (1 - fractional first period), matching the workbook's
    `AD11` / `AI11` cells.
    """
    if proceeds <= 0 or not any(cashflows):
        return 0.0

    def f(r: float) -> float:
        per = r / periods_per_year
        if per <= -1.0:
            return float("inf")
        return npv(per, cashflows) * (1.0 + per) ** offset_periods - proceeds

    f_lo, f_hi = f(lo), f(hi)
    if f_lo * f_hi > 0:                      # no sign change → give up gracefully
        return 0.0
    for _ in range(300):
        mid = (lo + hi) / 2.0
        f_mid = f(mid)
        if abs(f_mid) < tol or (hi - lo) < tol:
            return mid
        if f_lo * f_mid <= 0:
            hi, f_hi = mid, f_mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2.0


def round_down_to(value: float, increment: float) -> float:
    """Workbook's `INT(x / 5000) * 5000` denomination rounding."""
    if increment <= 0:
        return value
    return int(value / increment) * increment


def safe_div(numerator: float, denominator: float, default: float = 0.0) -> float:
    return numerator / denominator if denominator else default


def total(values: Iterable[float]) -> float:
    return sum(v for v in values if v)
