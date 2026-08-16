"""
summary.py — Summary sheet (assessed-value & pledged-revenue engine).

Converts the development program into annual taxable value and the pledged
revenue available to pay debt service, separately for the **senior** and
**subordinate** liens.

Key column mapping (Summary sheet):
  C  — AV-set / assessment year
  D  — tax-revenue (collection) year  (= C + 1)
  M  — cumulative home market value
  N  — residential AV in collection year = M[set-1] x residential taxable ratio
  J  — developed-lot taxable value = lot market value[set-1] x developed-lot %
  Q / AG — total taxable value
  AI — district debt-service mill-levy collections
  AJ — personal property uniform fee allocated to the district (§ 59-2-405)
  AX — net revenue available for SENIOR lien debt service
  BO — net revenue available for SUBORDINATE lien debt service

All series here are keyed by **collection year** — the year the matching debt
service is due.  In Utah the roll is set 1 January of the prior year and the
taxes are due 30 November of that prior year, so the money is in hand before the
1 March payment it supports.
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .config import ModelConfig
from .development import DeveloperProjections


@dataclass
class SummaryRow:
    """One year of the assessed-value / revenue projection (by collection year)."""
    collection_year: int
    assessment_year: int
    total_av: float
    residential_av: float          # N — homes
    lot_av: float                  # J — developed lots
    commercial_av: float           # Z — commercial (own mill levy)
    centrally_assessed_av: float              # AD — centrally assessed producing property
    state_av: float                # AE — state assessed
    mill_revenue: float            # AI
    uniform_fee_revenue: float  # AJ
    net_senior_revenue: float      # AX
    net_sub_revenue: float         # BO


class SummaryModel:
    """
    Reproduces the Summary sheet's assessed-value and revenue waterfall.

    Usage::

        cfg = ModelConfig()
        dev = DeveloperProjections().build(cfg)
        sm  = SummaryModel(cfg, dev).build()
        sm.net_senior_revenue(2030)   # revenue available for senior DS in 2030
    """

    def __init__(self, config: ModelConfig, developer: DeveloperProjections):
        self.cfg = config
        self.dev = developer
        self.rows: list[SummaryRow] = []
        self._by_year: dict[int, SummaryRow] = {}
        self._built = False

    # ── Build ───────────────────────────────────────────────────────────────
    def build(self) -> "SummaryModel":
        cfg = self.cfg
        dev = self.dev
        M = dev.cumulative_home_market_value

        self.rows = []
        self._by_year = {}

        # Lot-inventory and residential taxable value flow from the value builds
        # (the "Builder Lot Inventory Value" and "Residential Value" tabs) so the Summary,
        # those detail tabs, and bond sizing are all driven from one source — the
        # certified true-ups and their amortization are already baked in there.
        vac_build = {r["collection"]: r for r in dev.lot_inventory_value_build(cfg)}
        res_build = {r["collection"]: r for r in dev.residential_value_build(cfg)}

        # Collection years run from first_year+2 (the first roll actually
        # collected) through the last year the value builds cover — past the
        # senior final maturity there is no build to read, and a tail of zero
        # taxable value with fee-only negative revenue is not a projection.
        last_collect = min(dev.last_year,
                           max(vac_build) if vac_build else dev.last_year,
                           max(res_build) if res_build else dev.last_year)
        for collect in range(cfg.first_year + 2, last_collect + 1):
            s = collect - 1  # assessment / AV-set year
            src = cfg.av_source_year(collect)

            lot_av = vac_build.get(collect, {}).get("assessed", 0.0)
            residential_av = res_build.get(collect, {}).get("assessed", 0.0)
            # State-assessed and exempt come from the certified inputs for a roll year
            # entered in the historical table (e.g. 2024 state = 7,990); otherwise the
            # standing STATE_ASSESSED / EXEMPT_VALUE inputs apply.
            _hist = (cfg.historical_av or {}).get(collect - 1) or {}
            state_av_used = (_hist["state_assessed"] if _hist.get("state_assessed") is not None
                             else cfg.state_assessed)
            exempt_used = (_hist["exempt"] if _hist.get("exempt") is not None
                           else cfg.exempt_value)

            # ── Oil & gas producing property (AD) — assessed at its own ratio ─
            centrally_assessed_av = cfg.centrally_assessed_av

            # ── Commercial property (Z) — assessed at the commercial ratio ───
            csrc = collect - cfg.comm_assessment_lag_years
            if cfg.comm_new_value_add == "Yes":
                comm_mkt = dev.cumulative_commercial_market_value.get(csrc, 0.0)
                commercial_av = comm_mkt * cfg.commercial_assessment_ratio
            else:
                commercial_av = 0.0

            ds_taxable_av = (residential_av + lot_av + centrally_assessed_av + state_av_used
                             - exempt_used)
            total_av = ds_taxable_av + commercial_av

            # Mill-levy revenue (AI + AM) and specific-ownership tax (AJ + AN)
            mill_revenue = (
                ds_taxable_av / 1000.0 * cfg.effective_ds_mill_levy * cfg.tax_collect_mill_prc
                + commercial_av / 1000.0 * cfg.commercial_mill_levy * cfg.tax_collect_mill_prc
            )
            if total_av < cfg.uniform_fee_av_threshold:
                uniform_fee = mill_revenue * (cfg.uniform_fee_prc / 2)
            else:
                uniform_fee = mill_revenue * cfg.uniform_fee_prc

            # Net revenue available for SENIOR lien debt service (AX):
            #   mill + uniform fee - county collection cost - senior trustee fee
            #   - annual district administration (inflated, and not charged
            #     before the district is up and running)
            collection_fee = mill_revenue * cfg.county_collection_fee
            admin_cost, trustee_fee, trustee_fee_sub = cfg.district_costs(collect)
            if cfg.admin_cost_av_limit and total_av > cfg.admin_cost_av_limit:
                admin_cost = 0.0
            net_senior_revenue = (
                mill_revenue + uniform_fee - collection_fee - trustee_fee - admin_cost
            )

            # Net revenue available for SUBORDINATE lien debt service (BO):
            #   mill + Uniform Fee - subordinate trustee fee   (no treasurer fee netted here)
            net_sub_revenue = mill_revenue + uniform_fee - trustee_fee_sub

            row = SummaryRow(
                collection_year=collect,
                assessment_year=s,
                total_av=total_av,
                residential_av=residential_av,
                lot_av=lot_av,
                commercial_av=commercial_av,
                centrally_assessed_av=centrally_assessed_av,
                state_av=state_av_used,
                mill_revenue=mill_revenue,
                uniform_fee_revenue=uniform_fee,
                net_senior_revenue=net_senior_revenue,
                net_sub_revenue=net_sub_revenue,
            )
            self.rows.append(row)
            self._by_year[collect] = row

        self._built = True
        return self

    # ── Lookups ──────────────────────────────────────────────────────────────
    def _ensure(self):
        if not self._built:
            self.build()

    def row(self, collection_year: int) -> SummaryRow | None:
        self._ensure()
        return self._by_year.get(collection_year)

    # Presence flags — used to hide empty columns in the Excel / notebook output.
    @property
    def has_commercial(self) -> bool:
        return any(abs(r.commercial_av) > 0.5 for r in self.rows)

    @property
    def has_centrally_assessed(self) -> bool:
        return any(abs(r.centrally_assessed_av) > 0.5 for r in self.rows)

    @property
    def has_state_assessed(self) -> bool:
        return any(abs(r.state_av) > 0.5 for r in self.rows)

    def total_av(self, collection_year: int) -> float:
        r = self.row(collection_year)
        return r.total_av if r else 0.0

    def net_senior_revenue(self, collection_year: int) -> float:
        r = self.row(collection_year)
        return r.net_senior_revenue if r else 0.0

    def net_sub_revenue(self, collection_year: int) -> float:
        r = self.row(collection_year)
        return r.net_sub_revenue if r else 0.0

    def to_dataframe(self) -> pd.DataFrame:
        self._ensure()
        return pd.DataFrame([{
            "collection_year": r.collection_year,
            "assessment_year": r.assessment_year,
            "total_av": r.total_av,
            "residential_av": r.residential_av,
            "lot_av": r.lot_av,
            "commercial_av": r.commercial_av,
            "centrally_assessed_av": r.centrally_assessed_av,
            "state_av": r.state_av,
            "mill_revenue": r.mill_revenue,
            "uniform_fee_revenue": r.uniform_fee_revenue,
            "net_senior_revenue": r.net_senior_revenue,
            "net_sub_revenue": r.net_sub_revenue,
        } for r in self.rows])
