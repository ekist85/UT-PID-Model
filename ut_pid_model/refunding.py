"""
refunding.py — Refunding (refinancing) of the senior new-money bonds.

This is the capability examined in the Viridian Farm Metropolitan District model:
once the original senior new-money bonds (Series 2024A) become callable, they
are **refinanced** with a new senior refunding series (Series 2029).  Because by
then the assessed-value base — and therefore the pledged mill-levy revenue —
has grown, and because the refunding rate is lower than the original coupon, the
refunding bonds can be sized **larger** than the bonds being refunded.  After
defeasing the old bonds and paying transaction costs, the surplus proceeds
become **additional reimbursement ("new money")** to the developer.

Mechanics (Sources & Uses – Refunding):

  Sources
    • Refunding senior par (sized by the revenue-wrap at the refunding rate)
    • Released debt-service-reserve fund from the refunded series
    • Surplus funds on hand at the refunding date

  Uses
    • Refunding escrow to defease the old senior bonds
        = outstanding callable principal × (call price / 100)
    • (optional) Escrow to defease subordinate bonds
    • Bond insurance / surety
    • Underwriter's discount
    • Costs of issuance
    • Reimbursement  ← the balancing figure ("new money")
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date

from .config import ModelConfig
from .summary import SummaryModel
from .debt_service import BondTranche, SeniorLienSizer, CallProvisions
from .sources_uses import SourcesUses


@dataclass
class RefundingResult:
    refunding_bond: BondTranche
    sources_uses: SourcesUses
    refunded_par_outstanding: float       # callable principal of the old series
    call_price: float                     # redemption price applied (% of par)
    refunding_escrow: float               # cost to defease the old series
    new_money_reimbursement: float        # additional reimbursement generated
    savings_vs_original_rate: float       # PV-style coupon saving indicator


class RefundingAnalysis:
    """
    Sizes the senior refunding bonds and computes the additional reimbursement
    ("new money") created by refinancing the senior new-money bonds.
    """

    def __init__(self, config: ModelConfig, summary: SummaryModel):
        self.cfg = config
        self.sm = summary

    def run(
        self,
        senior_first: BondTranche,
        *,
        surplus_on_hand: float | None = None,
        sub_escrow: float | None = None,
    ) -> RefundingResult:
        cfg = self.cfg
        if surplus_on_hand is None:
            surplus_on_hand = cfg.refunding_surplus_on_hand or 0.0
        if sub_escrow is None:
            sub_escrow = cfg.refunding_sub_escrow or 0.0
        if cfg.refund_financing != "Yes":
            raise RuntimeError("REFUND_FINANCING is 'No' — refunding is disabled.")

        delivery = cfg.delivery_refunding              # first optional redemption date

        # ── Optional redemption of the refunded (old) series ─────────────────
        # The refunding only works once the old bonds are callable.  The price
        # paid to redeem them is the call price applicable on the refunding date
        # (premium, e.g. 103%, between the premium- and par-call dates; par
        # thereafter).  Principal maturing after the call date is defeased.
        call_provisions = senior_first.call_provisions or CallProvisions(
            premium_call_date=cfg.premium_call_date,
            par_call_date=cfg.par_call_date,
            premium_call_price=cfg.premium_call_price,
        )
        call_price = call_provisions.call_price_on(delivery)
        if call_price is None:
            raise RuntimeError(
                f"Senior bonds are call-protected until {call_provisions.premium_call_date}; "
                f"cannot refund on {delivery}."
            )
        refunded_outstanding = senior_first.callable_principal_after(
            date(call_provisions.premium_call_date.year,
                 senior_first.prin_month, senior_first.prin_day)
        )
        refunding_escrow = refunded_outstanding * (call_price / 100.0)

        # ── Released debt-service-reserve fund from the refunded series ───────
        released_dsrf = senior_first.dsrf_deposit

        # ── Size the refunding senior bonds via the revenue-wrap ─────────────
        sizer = SeniorLienSizer(cfg, self.sm)
        first_prin_year = delivery.year + 1
        final_year = delivery.year + cfg.final_mat_yrs_refunding
        # The refunding bonds are call-protected for the same period as the
        # new-money bonds, then callable at PAR — no declining-premium schedule.
        from .config import _edate
        ref_par_call = _edate(delivery, 12 * cfg.premium_call_years)
        refunding_calls = CallProvisions(
            premium_call_date=ref_par_call,
            par_call_date=ref_par_call,
            premium_call_price=100.0,     # par call (no redemption premium)
        )
        refunding_bond = sizer.size(
            name=f"Refunding Series {delivery.year}",
            rate=cfg.senior_refunding_interest_rate,
            coverage=cfg.dsc_refunding,
            delivery=delivery,
            first_principal_year=first_prin_year,
            final_year=final_year,
            capi_end_year=None,          # refunding bonds are not capitalized
            dsrf_deposit=0.0,            # reuses the released reserve
            release_surplus=False,
            call_provisions=refunding_calls,
            reoffering_yield=cfg.senior_refunding_reoffering_yield,
            coupon_scale=cfg.senior_refunding_coupon_scale,
            yield_scale=cfg.senior_refunding_yield_scale,
            term_bonds=cfg.senior_refunding_term_bonds,
        )

        # ── Sources & Uses ───────────────────────────────────────────────────
        su = SourcesUses(f"Refunding (Series {delivery.year})")
        su.sources["Refunding Senior Par"] = refunding_bond.par_amount
        su.sources["Released DSRF (refunded series)"] = released_dsrf
        su.sources["Surplus Funds on Hand"] = surplus_on_hand

        bond_insurance = refunding_bond.total_net_ds * cfg.bond_insurance_rate
        uwd = cfg.uwd_senior_refunding * refunding_bond.par_amount

        new_money = (
            su.total_sources
            - refunding_escrow
            - sub_escrow
            - bond_insurance
            - uwd
            - cfg.coi_refunding
        )

        su.uses["Reimbursement"] = new_money
        su.uses["Refunding Escrow (old senior)"] = refunding_escrow
        if sub_escrow:
            su.uses["Refunding Escrow (subordinate)"] = sub_escrow
        su.uses["Bond Insurance / Surety"] = bond_insurance
        su.uses["Underwriter's Discount"] = uwd
        su.uses["Costs of Issuance"] = cfg.coi_refunding

        # Indicative coupon saving: old coupon vs refunding coupon on refunded par.
        savings = (
            (senior_first.rate - cfg.senior_refunding_interest_rate)
            * refunded_outstanding
            * (final_year - delivery.year)
        )

        return RefundingResult(
            refunding_bond=refunding_bond,
            sources_uses=su,
            refunded_par_outstanding=refunded_outstanding,
            call_price=call_price,
            refunding_escrow=refunding_escrow,
            new_money_reimbursement=new_money,
            savings_vs_original_rate=savings,
        )
