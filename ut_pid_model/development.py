"""
development.py — residential and commercial absorption schedules.

Mirrors the "Residential Development" and "Comm Development" tabs.  Each tab is
a set of side-by-side blocks over the same year axis; this module produces the
same blocks as plain lists so the workbook writer can lay them out unchanged.

Residential blocks (Colorado column groups in brackets):
    lot_delivery          [D:Q]    finished lots delivered, by product
    lot_delivery_scenario [S:AF]   the above × LOT_DELIVERY_SCENARIO
    lot_value             [AH:AU]  lots × vacant land value per lot
    lot_value_lagged      [AW:BK]  lot_value shifted onto the tax roll
    lots_consumed         [BM:CA]  lot value removed as homes close
    home_closings         [CD:CQ]  closings, by product
    closings_scenario     [CS:DF]  the above × ABSORPTION_SCENARIO
    pricing               [DH:DU]  ASP inflated to the closing year
    av_creation           [DW:EL]  closings × inflated ASP
    av_creation_lagged    [EN:FC]  the above shifted onto the tax roll
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from .config import ModelConfig

#: Rows 14-40 of the development tabs — 27 projection years.
PROJECTION_YEARS = 27


@dataclass
class Product:
    """One builder/product column on the Residential Development tab."""
    name: str
    product_type: str = "SFD"           # SFD | TH | DU
    total_units: int = 0
    asp: float = 0.0
    #: {calendar year: units}
    lot_delivery: dict[int, int] = field(default_factory=dict)
    home_closings: dict[int, int] = field(default_factory=dict)

    def average_absorption(self) -> float:
        """Workbook row 47: mean of the non-zero closing years, rounded."""
        vals = [v for v in self.home_closings.values() if v > 0]
        return round(sum(vals) / len(vals)) if vals else 0.0

    def average_lot_delivery(self) -> float:
        vals = [v for v in self.lot_delivery.values() if v > 0]
        return round(sum(vals) / len(vals)) if vals else 0.0


@dataclass
class CommercialProduct:
    """One column on the Comm Development tab (square footage rather than lots)."""
    name: str
    total_sf: float = 0.0
    value_per_sf: float = 0.0
    sf_delivered: dict[int, float] = field(default_factory=dict)
    sf_sold: dict[int, float] = field(default_factory=dict)


@dataclass
class DevelopmentProjections:
    """The full absorption picture for one district."""

    start_year: int
    products: list[Product] = field(default_factory=list)
    commercial: list[CommercialProduct] = field(default_factory=list)
    #: Value already on the roll at the start of the projection.
    existing_home_market_value: float = 0.0
    existing_lot_value: float = 0.0

    # ── Year axis ─────────────────────────────────────────────────────────────

    def years(self) -> list[int]:
        return [self.start_year + i for i in range(PROJECTION_YEARS)]

    def assessment_dates(self, cfg: ModelConfig) -> list[date]:
        """Column C of the Summary tab: the 1 January valuation, carried on the
        district's principal-payment month so every tab shares one axis."""
        return [date(y, cfg.prin_maturity, 1) for y in self.years()]

    # ── Residential blocks ────────────────────────────────────────────────────

    def lot_delivery(self, cfg: ModelConfig) -> list[list[float]]:
        return [[float(p.lot_delivery.get(y, 0) or 0) for p in self.products]
                for y in self.years()]

    def lot_delivery_scenario(self, cfg: ModelConfig) -> list[list[float]]:
        return self._scenario(self.lot_delivery(cfg),
                              [p.average_lot_delivery() for p in self.products],
                              cfg.lot_delivery_scenario, cfg)

    def home_closings(self, cfg: ModelConfig) -> list[list[float]]:
        return [[float(p.home_closings.get(y, 0) or 0) for p in self.products]
                for y in self.years()]

    def closings_scenario(self, cfg: ModelConfig) -> list[list[float]]:
        return self._scenario(self.home_closings(cfg),
                              [p.average_absorption() for p in self.products],
                              cfg.absorption_scenario, cfg)

    def _scenario(self, base: list[list[float]], averages: list[float],
                  factor: float, cfg: ModelConfig) -> list[list[float]]:
        """
        Workbook columns S:AF / CS:DF.  When HYPOTHETICAL_SCENARIO is "No" the
        base schedule passes straight through; when it is "Yes" the schedule is
        rebuilt at `average × factor` units per year, capped at total units.
        """
        if cfg.hypothetical_scenario != "Yes":
            return [row[:] for row in base]

        out: list[list[float]] = []
        running = [0.0] * len(self.products)
        for r, row in enumerate(base):
            new_row: list[float] = []
            for i, p in enumerate(self.products):
                target = round(averages[i] * factor)
                remaining = p.total_units - running[i]
                if remaining >= p.total_units:          # nothing drawn yet
                    val = 0.0 if row[i] <= 0 else target
                else:
                    val = min(remaining, target)
                val = max(0.0, val)
                running[i] += val
                new_row.append(val)
            out.append(new_row)
        return out

    def vacant_land_value_per_lot(self, cfg: ModelConfig) -> list[float]:
        """Workbook row 11: finished-lot value as a share of the home ASP."""
        return [p.asp * cfg.platted_lot_value for p in self.products]

    def lot_value(self, cfg: ModelConfig) -> list[list[float]]:
        per_lot = self.vacant_land_value_per_lot(cfg)
        return [[units * per_lot[i] for i, units in enumerate(row)]
                for row in self.lot_delivery_scenario(cfg)]

    def lot_value_lagged(self, cfg: ModelConfig) -> list[list[float]]:
        """Shift lot value onto the tax roll (Utah: one year; Colorado: two)."""
        return _lag(self.lot_value(cfg), cfg.value_lag_years)

    def lots_consumed(self, cfg: ModelConfig) -> list[list[float]]:
        """
        Lot value backed out of the roll once a home closes on it.  The workbook
        negates the lagged lot value one row further on; economically the lot
        stops being taxed as a lot when the improvement lands.
        """
        lagged = self.lot_value_lagged(cfg)
        return [[-v for v in row] for row in _lag(lagged, 1)]

    def pricing(self, cfg: ModelConfig) -> list[list[float]]:
        """ASP inflated from the delivery year to each closing year (row DH:DU)."""
        base_year = cfg.resid_delivery_year.year
        out: list[list[float]] = []
        for y in self.years():
            steps = max(0, y - base_year)
            out.append([max(p.asp, p.asp * (1 + cfg.inflation_rate) ** steps)
                        if p.asp else 0.0 for p in self.products])
        return out

    def av_creation(self, cfg: ModelConfig) -> list[list[float]]:
        """Market value created by home closings (block DW:EL)."""
        prices = self.pricing(cfg)
        closings = self.closings_scenario(cfg)
        return [[closings[r][i] * prices[r][i] for i in range(len(self.products))]
                for r in range(len(self.years()))]

    def av_creation_lagged(self, cfg: ModelConfig) -> list[list[float]]:
        """Block EN:FC — the Utah addition that puts new homes on the next roll."""
        return _lag(self.av_creation(cfg), cfg.value_lag_years)

    # ── Commercial blocks ─────────────────────────────────────────────────────

    def comm_sf_delivered(self, cfg: ModelConfig) -> list[list[float]]:
        return [[float(c.sf_delivered.get(y, 0) or 0) for c in self.commercial]
                for y in self.years()]

    def comm_sf_sold(self, cfg: ModelConfig) -> list[list[float]]:
        return [[float(c.sf_sold.get(y, 0) or 0) for c in self.commercial]
                for y in self.years()]

    def comm_value_created(self, cfg: ModelConfig) -> list[list[float]]:
        sold = self.comm_sf_sold(cfg)
        return [[sold[r][i] * c.value_per_sf for i, c in enumerate(self.commercial)]
                for r in range(len(self.years()))]

    # ── Roll-ups used by the Summary tab ──────────────────────────────────────

    def total_units(self) -> int:
        return sum(p.total_units for p in self.products)

    def total_lot_units_by_year(self, cfg: ModelConfig) -> list[float]:
        return [sum(row) for row in self.lot_delivery_scenario(cfg)]

    def total_closings_by_year(self, cfg: ModelConfig) -> list[float]:
        return [sum(row) for row in self.closings_scenario(cfg)]

    def total_lot_value_lagged_by_year(self, cfg: ModelConfig) -> list[float]:
        consumed = self.lots_consumed(cfg)
        lagged = self.lot_value_lagged(cfg)
        return [sum(lagged[r]) + sum(consumed[r]) for r in range(len(self.years()))]

    def total_av_creation_lagged_by_year(self, cfg: ModelConfig) -> list[float]:
        return [sum(row) for row in self.av_creation_lagged(cfg)]

    def total_comm_value_by_year(self, cfg: ModelConfig) -> list[float]:
        return [sum(row) for row in self.comm_value_created(cfg)]

    def total_comm_sf_by_year(self, cfg: ModelConfig) -> list[float]:
        return [sum(row) for row in self.comm_sf_sold(cfg)]

    def market_value_at_buildout(self, cfg: ModelConfig) -> float:
        prices = self.pricing(cfg)
        closings = self.closings_scenario(cfg)
        return sum(closings[r][i] * prices[r][i]
                   for r in range(len(self.years()))
                   for i in range(len(self.products)))


def _lag(block: list[list[float]], years: int) -> list[list[float]]:
    """Shift a block down by `years` rows, zero-filling the top."""
    if years <= 0:
        return [row[:] for row in block]
    width = len(block[0]) if block else 0
    pad = [[0.0] * width for _ in range(years)]
    return (pad + [row[:] for row in block])[:len(block)]


# ── The Viridian Farm PID No. 1 schedule (this repo's reference deal) ─────────

def viridian_farm_projections() -> DevelopmentProjections:
    """Absorption exactly as priced on 17 September 2024."""
    lots = {
        "Rear-Load Townhome":  (139, 365_620.0, "TH",
                                {2024: 38, 2025: 60, 2026: 41}),
        "Front-Load Townhome": (327, 394_910.0, "DU",
                                {2024: 28, 2025: 48, 2026: 64, 2027: 96, 2028: 91}),
        "Alley-Load Cottages": (30,  434_350.0, "SFD", {2024: 24, 2025: 6}),
        "Front-Load Cottages": (88,  492_150.0, "SFD",
                                {2024: 20, 2025: 36, 2026: 32}),
        "8,000 Lots":          (23,  563_750.0, "SFD", {2023: 2, 2024: 21}),
        "12,000 Lots":         (67,  635_500.0, "SFD",
                                {2023: 2, 2024: 28, 2025: 30, 2026: 7}),
        "18,000 Lots":         (41,  709_813.0, "SFD",
                                {2024: 12, 2025: 22, 2026: 7}),
        "21,000 Lots":         (1,   761_063.0, "SFD", {2024: 1}),
    }
    products = [
        Product(name=name, product_type=ptype, total_units=units, asp=asp,
                lot_delivery=dict(sched),
                # Homes close the year after the finished lots are delivered.
                home_closings={y + 1: u for y, u in sched.items()})
        for name, (units, asp, ptype, sched) in lots.items()
    ]
    return DevelopmentProjections(start_year=2020, products=products)
