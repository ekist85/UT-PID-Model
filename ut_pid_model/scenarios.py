"""
scenarios.py — Base and stress-case scenario runner.

A PID forecast presents the financing under several development
scenarios.  The bonds are sized once in the **base case**; the **stress cases**
slow the development absorption (e.g. 80% and 45% of forecast pace) and test the
*same* debt service against the resulting lower taxable value, pledged revenue,
and debt-service coverage — exactly the Alternative Scenarios A and B in the
accountant's forecast.

``build_scenarios`` returns one ``Scenario`` per pace factor, each carrying the
fully-built model objects needed to render the forecast exhibits.
"""

from __future__ import annotations

from dataclasses import dataclass

from .config import ModelConfig
from .development import DeveloperProjections
from .summary import SummaryModel
from .debt_service import SeniorLienSizer, BondTranche, CallProvisions
from .subordinate import SubordinateLien, SurplusFund, SubLienResult
from .sources_uses import first_financing_sources_uses, SourcesUses


@dataclass
class Scenario:
    """A fully-built development scenario for the forecast exhibits."""
    exhibit: str            # "A", "B", "C"
    label: str              # human-readable description
    pace_factor: float
    cfg: ModelConfig
    dev: DeveloperProjections
    sm: SummaryModel
    senior: BondTranche
    su: SourcesUses
    sub: SubLienResult
    surplus: SurplusFund
    sub_par: float


def build_scenarios(
    cfg: ModelConfig,
    dev: DeveloperProjections | None = None,
    stress_pace_factors: tuple[float, ...] = (0.80, 0.45),
    sub_par: float | None = None,
    senior_dsrf: float | None = None,
    senior_first_principal_year: int | None = None,
    senior_final_year: int | None = None,
) -> list[Scenario]:
    """
    Build the base scenario plus one stress scenario per pace factor.

    The senior and subordinate **par amounts are fixed by the base case**; the
    stress scenarios reuse those bonds and recompute taxable value, pledged
    revenue, surplus fund, subordinate cash flow, and coverage under the slower
    absorption pace (first financing only, matching the forecast exhibits).

    ``dev`` is the development projection loaded from the inputs workbook; the
    stress cases are derived by slowing *its* absorption pace.  When omitted the
    model defaults are used (so nothing is silently hardcoded over the user's
    actual inputs).
    """
    base_dev = dev if dev is not None else DeveloperProjections()
    # Fall back to the config-derived (delivery-driven) defaults — flexible
    # across deals rather than pinned to fixed years/amounts.
    sub_par = cfg.sub_par if sub_par is None else sub_par
    senior_dsrf = cfg.senior_dsrf_deposit if senior_dsrf is None else senior_dsrf
    if senior_first_principal_year is None:
        senior_first_principal_year = cfg.senior_first_principal_year
    if senior_final_year is None:
        senior_final_year = cfg.senior_final_year

    from .debt_service import size_senior_with_dynamic_dsrf
    calls = CallProvisions(
        premium_call_date=cfg.premium_call_date,
        par_call_date=cfg.par_call_date,
        premium_call_price=cfg.premium_call_price,
    )
    dynamic_dsrf = senior_dsrf is None
    dynamic_subpar = sub_par is None

    def _build(pace: float, exhibit: str, label: str,
               base: Scenario | None) -> Scenario:
        dev = base_dev.stressed(pace).build(cfg)
        sm = SummaryModel(cfg, dev).build()
        if base is None:
            # Base case: size senior (with dynamic DSRF) from base-case revenue.
            size_kwargs = dict(
                name="Senior (2025A) Bonds",
                rate=cfg.senior_interest_rate, coverage=cfg.dsc_senior,
                delivery=cfg.delivery,
                first_principal_year=senior_first_principal_year,
                final_year=senior_final_year,
                capi_end_year=cfg.capi_end_date.year, call_provisions=calls,
                reoffering_yield=cfg.senior_reoffering_yield,
                coupon_scale=cfg.senior_coupon_scale,
                yield_scale=cfg.senior_yield_scale,
                term_bonds=cfg.senior_term_bonds,
            )
            sizer = SeniorLienSizer(cfg, sm)
            senior = (size_senior_with_dynamic_dsrf(sizer, **size_kwargs)
                      if dynamic_dsrf else sizer.size(dsrf_deposit=senior_dsrf, **size_kwargs))
            surplus = SurplusFund(cfg, sm).build(senior, None, cfg.first_collection_year, senior.final_year)
            spar = (SubordinateLien(cfg, sm).size_par(
                        senior, cfg.first_collection_year, senior.final_year, surplus_fund=surplus)
                    if dynamic_subpar else sub_par)
            su = first_financing_sources_uses(cfg, senior, sub_par=spar)
        else:
            # Stress case: reuse the base-case bonds (financing is fixed).
            senior, su, spar = base.senior, base.su, base.sub_par
            surplus = SurplusFund(cfg, sm).build(senior, None, cfg.first_collection_year, senior.final_year)
        sub = SubordinateLien(cfg, sm).size(
            spar, senior, cfg.first_collection_year, senior.final_year, surplus_fund=surplus)
        return Scenario(exhibit=exhibit, label=label, pace_factor=pace, cfg=cfg,
                        dev=dev, sm=sm, senior=senior, su=su, sub=sub,
                        surplus=surplus, sub_par=spar)

    base = _build(1.0, "A", "Base Case (100% of forecast absorption pace)", None)
    scenarios = [base]
    letters = ["B", "C", "D", "E"]
    for i, pace in enumerate(stress_pace_factors):
        scenarios.append(_build(
            pace, letters[i],
            f"Alternative Scenario — {pace:.0%} of forecast absorption pace",
            base))
    return scenarios
