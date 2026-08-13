"""
development.py — Residential & Commercial Development sheets.

Models the lot-delivery and home-closing absorption schedule provided by the
developer, and derives the market-value drivers that feed the assessed-value
engine (the "Summary" sheet):

  * ``home_closings[year]``   — number of homes that close (sell) that year
  * ``lot_market_value[year]``— market value of platted/developed lots placed
  * ``asp(year)``             — average selling price, inflated annually
  * ``cumulative_home_market_value[year]`` (Summary column M)
  * ``new_home_market_value[year]``        (homes closed x ASP)

Two input modes:

  * **Aggregate** (default) — a single absorption stream via the dict fields
    below (reproduces the Viridian Farm program: 716 residential units).
  * **Per-product** — a list of ``ProductLine`` objects (one per builder ×
    product, 1 to 8 lines, e.g. one builder/one product, or four builders with
    two products each).  Each product carries its own lot deliveries, home
    closings, lot value, and ASP; the aggregate drivers are rolled up from them.
"""

from __future__ import annotations

from dataclasses import dataclass, field


def _slow_schedule(schedule: dict[int, float], pace_factor: float) -> dict[int, float]:
    """Scale annual absorption by pace_factor, rolling deferred units forward."""
    if not schedule:
        return {}
    total = sum(schedule.values())
    years = sorted(schedule)
    avg_reduced = (total / len(years)) * pace_factor
    seq = [(y, schedule[y] * pace_factor) for y in years]
    placed = sum(v for _, v in seq)
    y = years[-1]
    while placed < total - 1e-6:
        y += 1
        amt = min(avg_reduced, total - placed)
        seq.append((y, amt))
        placed += amt
    out: dict[int, float] = {}
    for yr, v in seq:
        out[yr] = out.get(yr, 0.0) + v
    return out


@dataclass
class ProductLine:
    """
    One builder × product absorption line (e.g. "Builder 1 — 40' Product").

    A development may have a single product line or up to eight (e.g. four
    builders with two products each).
    """
    name: str
    # "Total Units" on the inputs page — the total planned residential units for
    # this product (the build-out count / cap).  This is an informational total
    # only; it is NOT pre-existing built value and does NOT seed the starting
    # home market value.  Home value accrues over time from ``home_closings``.
    existing_units: int = 0
    lot_deliveries: dict[int, int] = field(default_factory=dict)   # units delivered by year
    home_closings: dict[int, int] = field(default_factory=dict)    # units closed by year
    asp_base: float = 515_000                                      # base average selling price
    asp_base_year: int = 2024

    def asp(self, year: int, inflation: float, inflation_start_year: int) -> float:
        # Price is flat at the base until inflation starts, then compounds one year
        # of inflation per year — the inflation start year is the FIRST inflated
        # year (base × (1+i)^1).  Never deflated below the base.
        n = max(0, year - inflation_start_year + 1)
        return self.asp_base * (1.0 + inflation) ** n

    def lot_value(self, year: int, inflation: float, platted_pct: float,
                  inflation_start_year: int) -> float:
        """Market value of a delivered lot = ASP × platted-lot value %."""
        return self.asp(year, inflation, inflation_start_year) * platted_pct

    @property
    def total_units(self) -> int:
        """Total planned units: the 'Total Units' input, or the closings total."""
        return self.existing_units or int(round(sum(self.home_closings.values())))

    def stressed(self, pace_factor: float) -> "ProductLine":
        if pace_factor >= 1.0:
            return ProductLine(self.name, self.existing_units,
                               dict(self.lot_deliveries), dict(self.home_closings),
                               self.asp_base, self.asp_base_year)
        return ProductLine(
            self.name, self.existing_units,
            _slow_schedule(self.lot_deliveries, pace_factor),
            _slow_schedule(self.home_closings, pace_factor),
            self.asp_base, self.asp_base_year)


@dataclass
class DeveloperProjections:
    """
    Developer-supplied absorption assumptions, in aggregate or per-product mode.

    Aggregate defaults match Viridian Farm PID No. 1 (716 units, weighted-average
    base ASP $449,000 (2024) appreciating at the inflation rate).  Supply
    ``products`` to drive the model from per-builder/per-product detail instead.
    """

    base_asp: float = 449_000
    asp_base_year: int = 2024

    # Aggregate drivers (used when ``products`` is empty) — the Viridian Farm
    # PID No. 1 program: 716 units closing 2024-2029, lots delivered a year ahead.
    home_closings: dict[int, int] = field(default_factory=lambda: {
        2024: 4, 2025: 172, 2026: 202, 2027: 151, 2028: 96, 2029: 91,
    })
    lot_deliveries: dict[int, int] = field(default_factory=lambda: {
        2023: 4, 2024: 172, 2025: 202, 2026: 151, 2027: 96, 2028: 91,
    })
    lot_market_value: dict[int, float] = field(default_factory=lambda: {
        2023: 4 * 44_900.0, 2024: 172 * 44_900.0, 2025: 202 * 44_900.0,
        2026: 151 * 44_900.0, 2027: 96 * 44_900.0, 2028: 91 * 44_900.0,
    })

    # Commercial market value placed by year (Comm Development sheet driver).
    commercial_market_value: dict[int, float] = field(default_factory=dict)

    total_lots: int = 716
    pace_factor: float = 1.0

    # Per-product input (overrides the aggregate drivers when non-empty).
    products: list[ProductLine] = field(default_factory=list)

    # ── Stress cases ──────────────────────────────────────────────────────
    def stressed(self, pace_factor: float) -> "DeveloperProjections":
        """Return a new projection at a reduced absorption pace (stress case)."""
        if self.products:
            return DeveloperProjections(
                base_asp=self.base_asp, asp_base_year=self.asp_base_year,
                commercial_market_value=_slow_schedule(self.commercial_market_value, pace_factor)
                if (self.commercial_market_value and pace_factor < 1.0) else dict(self.commercial_market_value),
                total_lots=self.total_lots, pace_factor=pace_factor,
                products=[p.stressed(pace_factor) for p in self.products],
            )
        if pace_factor >= 1.0:
            return DeveloperProjections(
                base_asp=self.base_asp, asp_base_year=self.asp_base_year,
                home_closings=dict(self.home_closings),
                lot_market_value=dict(self.lot_market_value),
                lot_deliveries=dict(self.lot_deliveries),
                commercial_market_value=dict(self.commercial_market_value),
                total_lots=self.total_lots, pace_factor=1.0,
            )
        return DeveloperProjections(
            base_asp=self.base_asp, asp_base_year=self.asp_base_year,
            home_closings=_slow_schedule(self.home_closings, pace_factor),
            lot_market_value=_slow_schedule(self.lot_market_value, pace_factor),
            lot_deliveries=dict(self.lot_deliveries),
            commercial_market_value=_slow_schedule(self.commercial_market_value, pace_factor)
            if self.commercial_market_value else {},
            total_lots=self.total_lots, pace_factor=pace_factor,
        )

    # ── Queries ───────────────────────────────────────────────────────────
    def asp(self, year: int) -> float:
        """Aggregate average selling price (no-products mode)."""
        n = max(0, year - self._inflation_start_year + 1)
        return self.base_asp * (1.0 + self._infl) ** n

    def new_home_market_value(self, year: int) -> float:
        """Market value of homes closing in ``year`` (sum across products)."""
        if self.products:
            return sum(p.home_closings.get(year, 0) * p.asp(year, self._infl, self._inflation_start_year)
                       for p in self.products)
        return self.home_closings.get(year, 0) * self.asp(year)

    # ── Lot-inventory inventory (drawn down as homes complete) ──────────────────
    def _cumulative_lots(self, year: int) -> tuple[float, float, float]:
        """
        Cumulative (lots delivered units, lots delivered market value, homes
        closed units) through ``year``.
        """
        cum_units = cum_value = cum_closed = 0.0
        for y in sorted(set(self.lot_deliveries) | set(self.home_closings)):
            if y > year:
                break
            cum_units += self.lot_deliveries.get(y, 0)
            cum_value += self.lot_market_value.get(y, 0.0)
            cum_closed += self.home_closings.get(y, 0)
        return cum_units, cum_value, cum_closed

    def vacant_lot_units(self, year: int) -> float:
        """Platted/finished lots still vacant as of ``year`` = delivered − closed."""
        cum_units, _, cum_closed = self._cumulative_lots(year)
        return max(cum_units - cum_closed, 0.0)

    def vacant_lot_market_value(self, year: int) -> float:
        """
        Market value of lots still vacant as of ``year``: the remaining lot
        count (cumulative lots delivered − cumulative homes closed) valued at the
        average delivered lot value.  Lots that have been built out into homes
        are removed — their value is captured in the cumulative home market value
        instead — so vacant lots and completed homes are not double-counted.
        """
        cum_units, cum_value, cum_closed = self._cumulative_lots(year)
        if cum_units <= 0:
            return 0.0
        remaining = max(cum_units - cum_closed, 0.0)
        return remaining * (cum_value / cum_units)

    def existing_value_adjustments(self, cfg) -> dict:
        """
        Existing-value true-up credit and its amortization, keyed by ROLL year.

        For each roll year entered in the Inputs historical table:
          * Vacant land credit  = (entered vacant AV / vacant rate)  −  the
            development-page "lot value to homes" for that year.
          * Residential credit  = cumulative home market value  −  (entered
            residential AV / residential rate)  −  home value added.
        Both credits are recognised (positive) in the entry roll year so the deal
        gets credit for the initial land/home value, then amortised back out
        (equal negatives) so it is not double-counted as lots become homes:
          * vacant   → the year after the last lot delivery through build-out;
          * residential → one year after the vacant window (build-out + 1).
        """
        hist = getattr(cfg, "historical_av", None) or {}
        if not hist:
            return {}
        out: dict[int, float] = {}
        vac_sum = res_sum = 0.0
        for R in sorted(hist):
            h = hist[R]
            vac = h.get("vacant_land")
            res = h.get("residential")
            if vac:
                vrate = cfg.lot_inventory_taxable_rate(R + 1)
                trued = vac / vrate if vrate else 0.0
                base = (self.vacant_lot_market_value(R - 1)
                        + self.lot_market_value.get(R, 0.0)
                        - self.vacant_lot_market_value(R))
                adj = trued - base
                out[R] = out.get(R, 0.0) + adj
                vac_sum += adj
            if res:
                rrate = cfg.residential_assessment_rate(R + 1)
                beg = res / rrate if rrate else 0.0
                added = self.new_home_market_value(R)
                cum = self.cumulative_home_market_value.get(R, 0.0)
                adj = cum - beg - added
                out[R] = out.get(R, 0.0) + adj
                res_sum += adj
        lot_years = [y for y, u in self.lot_deliveries.items() if u]
        home_years = [y for y, u in self.home_closings.items() if u]
        if lot_years and home_years:
            lld, lhc = max(lot_years), max(home_years)
            # Vacant rolls off from the year after the last lot delivery through
            # one year past build-out (last home closing); residential one year
            # after vacant throughout.
            vac_win = list(range(lld + 1, lhc + 2))          # e.g. 2027..2030
            res_win = [y + 1 for y in vac_win]               # e.g. 2028..2031
            if vac_win and abs(vac_sum) > 1e-6:
                per = vac_sum / len(vac_win)
                for y in vac_win:
                    out[y] = out.get(y, 0.0) - per
            if res_win and abs(res_sum) > 1e-6:
                per = res_sum / len(res_win)
                for y in res_win:
                    out[y] = out.get(y, 0.0) - per
        return out

    @property
    def is_per_product(self) -> bool:
        return bool(self.products)

    # ── Build ───────────────────────────────────────────────────────────────
    def build(self, cfg, last_year: int = 2067) -> "DeveloperProjections":
        """
        Roll up per-product detail (if any) into the aggregate drivers, then
        compute the cumulative home market value series (Summary column M) with
        biennial reassessment.
        """
        self._infl = cfg.inflation_rate
        self._inflation_start_year = cfg.inflation_start_year
        self.last_year = last_year

        # Derive aggregate drivers from products when in per-product mode.
        if self.products:
            years = set()
            for p in self.products:
                years |= set(p.lot_deliveries) | set(p.home_closings)
            yrs = sorted(years)
            self.home_closings = {
                y: sum(p.home_closings.get(y, 0) for p in self.products) for y in yrs}
            self.lot_deliveries = {
                y: sum(p.lot_deliveries.get(y, 0) for p in self.products) for y in yrs}
            # Lot market value = lots delivered × (ASP × platted-lot value %).
            self.lot_market_value = {
                y: sum(p.lot_deliveries.get(y, 0)
                       * p.lot_value(y, self._infl, cfg.platted_lot_value,
                                     cfg.inflation_start_year)
                       for p in self.products) for y in yrs}
            self.total_lots = sum(p.total_units for p in self.products)

        # The home market value starts at zero (greenfield) and accrues as homes
        # close.  "Total Units" is the planned build-out count, not pre-built
        # value, so it does NOT seed the starting cumulative home market value.
        M: dict[int, float] = {cfg.first_year: 0.0}
        for y in range(cfg.first_year + 1, last_year + 1):
            prev = M[y - 1]
            reassess = prev * cfg.reassess_rate if self._is_reassess_year(y, cfg) else 0.0
            M[y] = prev + self.new_home_market_value(y) + reassess

        # Existing-home market-value floor from the certified roll.  When the
        # district carries a hard-coded existing residential ASSESSED value
        # (> $100k), gross it up to MARKET value using the primary residential
        # taxable ratio for the certification year (taxable ÷ ratio).  If that implied market
        # value exceeds the modeled cumulative home value on the certification
        # roll, seed the cumulative home value with it (existing homes are already
        # worth more than the greenfield build has reached) and re-accrue forward.
        from .config import residential_taxable_ratio_for
        # Adjustment (delta) applied to the cumulative home market value in the
        # certification year, exposed for a highlighted column in the report.
        self.cert_mv_adjustment_year: int | None = None
        self.cert_mv_adjustment: float = 0.0
        if (cfg.certification_date is not None
                and cfg.existing_residential_value > 100_000):
            cy = cfg.certification_date.year
            ratio = residential_taxable_ratio_for(cy)
            implied_mv = cfg.existing_residential_value / ratio if ratio else 0.0
            lov_year = cfg.av_source_year(cy + 1)   # level of value backing the cert roll
            calc_mv = M.get(lov_year, 0.0)
            if implied_mv > calc_mv:
                self.cert_mv_adjustment_year = lov_year
                self.cert_mv_adjustment = implied_mv - calc_mv
                M[lov_year] = implied_mv
                for y in range(lov_year + 1, last_year + 1):
                    reassess = (M[y - 1] * cfg.reassess_rate
                                if self._is_reassess_year(y, cfg) else 0.0)
                    M[y] = M[y - 1] + self.new_home_market_value(y) + reassess
        self.cumulative_home_market_value = M

        # Cumulative commercial market value.
        C: dict[int, float] = {cfg.first_year: 0.0}
        for y in range(cfg.first_year + 1, last_year + 1):
            prev = C[y - 1]
            reassess = (prev * cfg.reassess_comm_rate
                        if self._is_reassess_year(y, cfg) else 0.0)
            C[y] = prev + self.commercial_market_value.get(y, 0.0) + reassess
        self.cumulative_commercial_market_value = C
        return self

    @staticmethod
    def _is_reassess_year(year: int, cfg) -> bool:
        """
        Should value already on the roll be grown this year?

        Utah county assessors update values **annually** based on a systematic
        review of current market data (§ 59-2-303.1), with a detailed review of
        each parcel at least every five years — so the answer is normally "every
        year".  Set ``reassess_frequency = "Biennial"`` to model Colorado's
        two-year reappraisal cycle instead, in which case the step is booked on
        the level-of-value year that feeds the odd (or even) roll year.

        No growth is applied on a certified roll: the county-certified value
        already IS the roll for that year, so the model must not layer a
        reassessment on top of it.
        """
        if year < cfg.first_year + 2:
            return False
        roll_year = year + 1
        cert_date = getattr(cfg, "certification_date", None)
        if cert_date is not None and roll_year == cert_date.year:
            return False
        if getattr(cfg, "reassess_frequency", "Annual").strip().lower() != "biennial":
            return True
        return ((roll_year % 2 == 0) if getattr(cfg, "reassess_on_even_years", False)
                else (roll_year % 2 == 1))

    # ── Per-product display helpers (for the Residential tab) ───────────────
    def product_lines(self) -> list[ProductLine]:
        """Return the product lines, or a single synthetic line in aggregate mode."""
        if self.products:
            return self.products
        return [ProductLine(
            name="All Products",
            lot_deliveries=dict(self.lot_deliveries),
            home_closings=dict(self.home_closings),
            asp_base=self.base_asp, asp_base_year=self.asp_base_year,
        )]


# ── The Viridian Farm PID No. 1 program (this repo's reference deal) ──────────

#: Product, units, base ASP, and the lot-delivery schedule priced 17 Sept 2024.
#: Homes close the year after their finished lots are delivered.
VIRIDIAN_FARM_PRODUCTS = [
    ("Rear-Load Townhome",  139, 365_620.0, {2023: 0, 2024: 38, 2025: 60, 2026: 41}),
    ("Front-Load Townhome", 327, 394_910.0,
     {2024: 28, 2025: 48, 2026: 64, 2027: 96, 2028: 91}),
    ("Alley-Load Cottages",  30, 434_350.0, {2024: 24, 2025: 6}),
    ("Front-Load Cottages",  88, 492_150.0, {2024: 20, 2025: 36, 2026: 32}),
    ("8,000 Lots",           23, 563_750.0, {2023: 2, 2024: 21}),
    ("12,000 Lots",          67, 635_500.0, {2023: 2, 2024: 28, 2025: 30, 2026: 7}),
    ("18,000 Lots",          41, 709_813.0, {2024: 12, 2025: 22, 2026: 7}),
    ("21,000 Lots",           1, 761_063.0, {2024: 1}),
]


def viridian_farm_projections(asp_base_year: int = 2024) -> "DeveloperProjections":
    """Per-product absorption for Viridian Farm PID No. 1, exactly as priced."""
    products = [
        ProductLine(
            name=name, existing_units=units,
            lot_deliveries={y: u for y, u in sched.items() if u},
            home_closings={y + 1: u for y, u in sched.items() if u},
            asp_base=asp, asp_base_year=asp_base_year,
        )
        for name, units, asp, sched in VIRIDIAN_FARM_PRODUCTS
    ]
    return DeveloperProjections(products=products, total_lots=716)
