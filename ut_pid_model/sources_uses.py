"""
sources_uses.py — Sources & Uses of Funds (first financing and refunding).

Translates each sized bond tranche into its Sources & Uses statement and, most
importantly, the **reimbursement** paid to the developer:

    Reimbursement = Total Sources − DSRF deposit − Capitalized Interest
                    − Underwriter's Discount − Costs of Issuance
                    (− refunding escrow, for the refunding)

The reimbursement is the developer's recovery of eligible public-improvement
costs and is the headline output of a PID financing.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .config import ModelConfig
from .debt_service import BondTranche


@dataclass
class SourcesUses:
    label: str
    sources: dict[str, float] = field(default_factory=dict)
    uses: dict[str, float] = field(default_factory=dict)

    @property
    def total_sources(self) -> float:
        return sum(self.sources.values())

    @property
    def total_uses(self) -> float:
        return sum(self.uses.values())

    @property
    def reimbursement(self) -> float:
        return self.uses.get("Reimbursement", 0.0)

    @property
    def balanced(self) -> bool:
        return abs(self.total_sources - self.total_uses) < 1.0


def allocate_contribution(amount: float, series: str, senior_par: float,
                          sub_par: float, series_c_par: float) -> dict:
    """
    Split a developer cash contribution across the bond series per the toggle:
    Senior / Subordinate / Series C apply it wholly to that series; Proportional
    spreads it across all present series by par.  Returns
    ``{"senior", "subordinate", "series_c"}`` dollar amounts.
    """
    out = {"senior": 0.0, "subordinate": 0.0, "series_c": 0.0}
    if not amount or amount <= 0:
        return out
    s = (series or "").strip().lower()
    if s.startswith("senior"):
        out["senior"] = amount
    elif s.startswith("sub"):
        out["subordinate"] = amount
    elif "series c" in s or s in ("c", "series_c", "seriesc"):
        out["series_c"] = amount
    else:  # proportional (default)
        total = senior_par + sub_par + series_c_par
        if total <= 0:
            out["senior"] = amount
        else:
            out["senior"] = amount * senior_par / total
            out["subordinate"] = amount * sub_par / total
            out["series_c"] = amount * series_c_par / total
    return out


def first_financing_sources_uses(
    cfg: ModelConfig,
    senior: BondTranche,
    sub_par: float,
    series_c_par: float = 0.0,
    contribution: dict | None = None,
) -> SourcesUses:
    """
    Sources & Uses for the original (new-money) financing — senior + subordinate
    (+ optional Series C) liens, plus an optional developer cash contribution
    (a source applied to a chosen series, boosting that series' reimbursement).
    """
    capi = sum(p.capitalized_interest for p in senior.schedule)
    uwd_senior = cfg.uwd_senior * senior.par_amount
    uwd_sub = cfg.uwd_sub * sub_par
    uwd_series_c = cfg.uwd_sub * series_c_par   # cash-flow bond — sub UWD rate
    premium = senior.total_premium   # net premium/(discount) — an extra source
    dc = contribution or {"senior": 0.0, "subordinate": 0.0, "series_c": 0.0}

    su = SourcesUses("First Financing")
    su.sources["Senior Par"] = senior.par_amount
    su.sources["Subordinate Par"] = sub_par
    if series_c_par:
        su.sources["Series C Par"] = series_c_par
    if abs(premium) > 0.5:
        su.sources["Plus: Premium / (Discount)"] = premium
    total_dc = dc["senior"] + dc["subordinate"] + dc["series_c"]
    if total_dc:
        su.sources["Developer Contribution"] = total_dc

    senior_reimb = (
        senior.par_amount + premium - senior.dsrf_deposit - capi - uwd_senior
        - cfg.coi + dc["senior"]
    )
    sub_reimb = sub_par - uwd_sub + dc["subordinate"]
    series_c_reimb = series_c_par - uwd_series_c + dc["series_c"]

    su.uses["Reimbursement"] = senior_reimb + sub_reimb + series_c_reimb
    su.uses["Debt Service Reserve Fund"] = senior.dsrf_deposit
    su.uses["Capitalized Interest"] = capi
    su.uses["Underwriter's Discount"] = uwd_senior + uwd_sub + uwd_series_c
    su.uses["Costs of Issuance"] = cfg.coi
    return su
