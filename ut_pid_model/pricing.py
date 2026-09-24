"""
pricing.py — Bond pricing (price / yield / premium-OID), priced to worst call.

Implements the clean-price PV used by the Excel ``PRICE()`` column on the Senior
Lien DS sheets: a bond is priced to **yield-to-worst** — the minimum price across
pricing to maturity (at par), to the premium-call date (at the premium call
price, e.g. 103), and to the par-call date (at 100).  When the reoffering yield
equals the coupon the price is exactly par (100) and there is no premium/OID.

Settlement is the bond's dated date and all maturities fall on coupon dates, so
the clean price reduces to the standard level-coupon present value (no accrued
interest adjustment).
"""

from __future__ import annotations

from datetime import date


def days_30_360(d1: date, d2: date) -> int:
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
    return price - c * accrued


def bond_price(n_periods: float, coupon: float, ytm: float,
               redemption: float = 100.0, freq: int = 2) -> float:
    """
    Clean price per 100 of face, settling on a coupon date.

      price = redemption / (1+i)^n  +  (coupon/freq*100) * [1 - (1+i)^-n] / i
    where i = ytm/freq.  With ytm == coupon and redemption == 100 this is 100.
    """
    if n_periods <= 0:
        return redemption
    i = ytm / freq
    c = coupon / freq * 100.0
    if abs(i) < 1e-12:
        return redemption + c * n_periods
    disc = (1.0 + i) ** (-n_periods)
    pv_coupons = c * (1.0 - disc) / i
    return pv_coupons + redemption * disc


def price_to_worst(settlement: date, maturity: date, coupon: float, ytm: float,
                   calls: list[tuple[date, float]], freq: int = 2) -> float:
    """
    Yield-to-worst clean price: the minimum price across pricing to maturity
    (at 100) and to each call date (at its call price) that precedes maturity.
    """
    scenarios = [(maturity, 100.0)]
    for call_date, call_price in calls:
        if settlement < call_date < maturity:
            scenarios.append((call_date, call_price))
    # A bond reoffered AT its coupon is a par bond and is quoted at 100 — which
    # is what the underwriter's run shows for the 9.000%/9.000% series.  Excel's
    # PRICE() returns 99.90 for that off a coupon date, because it discounts the
    # odd first period compound while interest accrues across it simple; quoting
    # par is the market convention, not a fudge.
    if abs(ytm - coupon) < 1e-12:
        return 100.0
    prices = [clean_price(settlement, red_date, coupon, ytm, red_price, freq)
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
    return price + redemption * df(redemption_date)
