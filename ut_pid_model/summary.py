"""
summary.py — Summary sheet (taxable-value & pledged-revenue engine).

Converts the development program into annual taxable value and the pledged
revenue available to pay debt service, separately for the **senior** and
**subordinate** liens.

Utah taxes primary residential property on 55% of fair market value (the 45%
primary residential exemption, § 59-2-103), and builder lot inventory at the
same ratio where the assessor determines the property will be a primary
residence once occupied (Utah Admin. Code R884-24P-52).  Both ratios come from
``ModelConfig``, so a district that carries inventory at full market value need
only change the input.

Key column mapping (Summary sheet):
  C  — AV-set / assessment year
  D  — tax-revenue (collection) year  (= C + 1)
  M  — cumulative home market value
  N  — residential AV in collection year = M[set-1] x residential taxable ratio
  J  — developed-lot AV = lot market value[set-1] x developed-lot %
  Q / AG — total taxable value
  AI — district debt-service mill-levy collections
  AJ — specific-ownership tax collected
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

        # Existing-value true-up credit + amortization, keyed by roll year.
        existing_adj = dev.existing_value_adjustments(cfg)

        # Collection years run from first_year+2 (first AV actually collected)
        for collect in range(cfg.first_year + 2, dev.last_year + 1):
            s = collect - 1  # assessment / AV-set year

            # ── Utah taxable-value timing ────────────────────────────────────
            # Value created in calendar year V is assessed on the 1 January roll
            # for V+1, billed that November, and pays debt service the following
            # 1 March — a two-year lag from creation to the payment it supports.
            # ``av_source_year`` applies that lag; the ratio lookups apply the
            # primary residential exemption and the builder-inventory treatment.
            src = cfg.av_source_year(collect)
            residential_av = M.get(src, 0.0) * cfg.residential_assessment_rate(collect)

            # Developed-lot / lot-inventory AV (Summary J).  Uses the *vacant* lot
            # inventory (lots delivered less lots built out into homes), so a lot
            # and the home later built on it are never both counted.
            lot_av = (dev.vacant_lot_market_value(src)
                      * cfg.lot_inventory_taxable_rate(collect))

            # ── Oil & gas producing property (AD) — assessed at its own ratio ─
            centrally_assessed_av = cfg.centrally_assessed_av

            # ── Commercial property (Z) — assessed at the commercial ratio ───
            csrc = collect - cfg.comm_assessment_lag_years
            if cfg.comm_new_value_add == "Yes":
                comm_mkt = dev.cumulative_commercial_market_value.get(csrc, 0.0)
                commercial_av = comm_mkt * cfg.commercial_assessment_ratio
            else:
                commercial_av = 0.0

            # Debt-service-taxed value (residential + lots + centrally assessed
            # + state) is
            # levied at the DS mill levy; commercial is levied at the commercial
            # mill levy.  Existing builder lot inventory and existing residential value are
            # ONE-TIME certified amounts placed on the certified roll; state
            # assessed is added and the exempt value subtracted every year.
            # The certified roll (certification_date.year) is collected the FOLLOWING
            # year, so the certified base anchors that collection year.
            cert_collect_year = (cfg.certification_date.year + 1
                                 if cfg.certification_date is not None
                                 else cfg.first_collection_year)
            state_av_used = cfg.state_assessed
            exempt_used = cfg.exempt_value

            # Historical certified-value true-up: for a roll year entered in the
            # Inputs side table (collected the FOLLOWING year), the entered county-
            # certified taxable values ARE the actual roll.  Override the DS-taxable
            # categories with the entered figures (a blank cell counts as $0), and
            # skip the greenfield "existing once" / single-year certified plug so
            # nothing is double-counted.
            hist = (cfg.historical_av or {}).get(collect - 1)
            if hist:
                lot_av = hist.get("vacant_land") or 0.0
                residential_av = hist.get("residential") or 0.0
                state_av_used = hist.get("state_assessed") or 0.0
                exempt_used = hist.get("exempt") or 0.0
                existing_once = 0.0
                ds_taxable_av = (residential_av + lot_av + centrally_assessed_av
                                 + state_av_used - exempt_used)
                total_av = ds_taxable_av + commercial_av
            else:
                existing_once = (
                    (cfg.existing_vacant_land + cfg.existing_residential_value)
                    if collect == cert_collect_year else 0.0)
                ds_taxable_av = (residential_av + lot_av + centrally_assessed_av + state_av_used
                                 + existing_once - exempt_used)
                total_av = ds_taxable_av + commercial_av

                # Certified-value true-up: plug Total AV to the county-certified
                # value from the Inputs page (the actual roll) in the certified
                # roll's collection year.  The difference (+/-) adjusts the
                # DS-taxable base so revenue is levied on the certified value.
                if cfg.certification_date is not None and collect == cert_collect_year:
                    cert_adjustment = cfg.current_certified_value - total_av
                    ds_taxable_av += cert_adjustment
                    total_av = ds_taxable_av + commercial_av

            # Existing-value credit / amortization (keyed by roll year = s) flows
            # into the DS-taxable base and pledged revenue: positive credit in the
            # entry year, negative amortization as the land builds out.
            existing_credit = existing_adj.get(s, 0.0)
            if existing_credit:
                ds_taxable_av += existing_credit
                total_av = ds_taxable_av + commercial_av

            # Mill-levy revenue (AI + AM) and the personal property uniform fee
            # (AJ + AN), which § 59-2-405 distributes to taxing entities in the
            # same proportion as ad valorem real property tax.
            mill_revenue = (
                ds_taxable_av / 1000.0 * cfg.effective_ds_mill_levy * cfg.tax_collect_mill_prc
                + commercial_av / 1000.0 * cfg.commercial_mill_levy * cfg.tax_collect_mill_prc
            )
            if total_av < cfg.uniform_fee_av_threshold:
                sot = mill_revenue * (cfg.uniform_fee_prc / 2)
            else:
                sot = mill_revenue * cfg.uniform_fee_prc

            # Net revenue available for SENIOR lien debt service (AX):
            #   mill + uniform fee - county collection cost - senior trustee fee
            #   - annual district administration (inflated, and not charged
            #     before the district is up and running)
            treasurer_fee = mill_revenue * cfg.county_collection_fee
            admin_cost, trustee_fee, trustee_fee_sub = cfg.district_costs(collect)
            if cfg.admin_cost_av_limit and total_av > cfg.admin_cost_av_limit:
                admin_cost = 0.0
            net_senior_revenue = (
                mill_revenue + sot - treasurer_fee - trustee_fee - admin_cost
            )

            # Net revenue available for SUBORDINATE lien debt service (BO):
            #   mill + uniform fee - subordinate trustee fee
            net_sub_revenue = mill_revenue + sot - trustee_fee_sub

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
                uniform_fee_revenue=sot,
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
