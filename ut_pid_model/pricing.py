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


def semiannual_periods(settlement: date, redemption: date, freq: int = 2) -> float:
    """Number of coupon periods between two coupon-cycle dates."""
    months = (redemption.year - settlement.year) * 12 + (redemption.month - settlement.month)
    return months / (12.0 / freq)


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
    prices = []
    for red_date, red_price in scenarios:
        n = semiannual_periods(settlement, red_date, freq)
        if n > 0:
            prices.append(bond_price(n, coupon, ytm, red_price, freq))
    return min(prices) if prices else 100.0
