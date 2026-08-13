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


def first_financing_sources_uses(
    cfg: ModelConfig,
    senior: BondTranche,
    sub_par: float,
) -> SourcesUses:
    """
    Sources & Uses for the original (new-money) financing — senior + sub liens.
    """
    capi = sum(p.capitalized_interest for p in senior.schedule)
    uwd_senior = cfg.uwd_senior * senior.par_amount
    uwd_sub = cfg.uwd_sub * sub_par
    premium = senior.total_premium   # net premium/(discount) — an extra source

    su = SourcesUses("First Financing")
    su.sources["Senior Par"] = senior.par_amount
    su.sources["Subordinate Par"] = sub_par
    if abs(premium) > 0.5:
        su.sources["Plus: Premium / (Discount)"] = premium

    senior_reimb = (
        senior.par_amount + premium - senior.dsrf_deposit - capi - uwd_senior - cfg.coi
    )
    sub_reimb = sub_par - uwd_sub

    su.uses["Reimbursement"] = senior_reimb + sub_reimb
    su.uses["Debt Service Reserve Fund"] = senior.dsrf_deposit
    su.uses["Capitalized Interest"] = capi
    su.uses["Underwriter's Discount"] = uwd_senior + uwd_sub
    su.uses["Costs of Issuance"] = cfg.coi
    return su
