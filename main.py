"""
main.py — Entry point for the Utah Public Infrastructure District financial model
           (UCA 17D-4).

Run:
    python main.py

Outputs:
  - Console summary matching the key workbook outputs
  - CSV exports of the major tables
  - A formatted multi-sheet Excel report
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from ut_pid_model import (
    ModelConfig,
    DeveloperProjections,
    SummaryModel,
    SeniorLienSizer,
    size_senior_with_dynamic_dsrf,
    standard_dsrf,
    CallProvisions,
    SubordinateLien,
    SurplusFund,
    RefundingAnalysis,
    first_financing_sources_uses,
    schedule_dataframe,
    senior_coverage_dataframe,
    build_excel_report,
    build_forecast_report,
    build_scenarios,
    write_inputs_workbook,
    build_memo_html,
)


def run_model(export: bool = True, output_dir: str | None = None,
              inputs_path: str | None = None) -> dict:
    # Deliverables land in a "reimbursement analysis" folder sitting alongside
    # the inputs workbook that drives the run (or in ./output when using the
    # built-in defaults). The caller can still force a location via output_dir.
    if output_dir is None:
        if inputs_path:
            output_dir = os.path.join(
                os.path.dirname(os.path.abspath(inputs_path)),
                "reimbursement analysis")
        else:
            output_dir = "output"
    print("=" * 72)
    print("  Utah Public Infrastructure District Financial Model")
    print("  UCA 17D-4  |  senior / subordinate liens + refunding")
    print("=" * 72)

    # ── 1. Configuration ──────────────────────────────────────────────────
    # Load config + development from an Excel inputs template when given, else
    # use built-in defaults.
    if inputs_path:
        from ut_pid_model import load_inputs_workbook
        cfg, dev = load_inputs_workbook(inputs_path)
        print(f"\n[1] Config loaded from: {inputs_path}")
    else:
        from ut_pid_model import viridian_farm_projections
        cfg, dev = ModelConfig(), viridian_farm_projections()
        print(f"\n[1] Config: {cfg.pid_name or '(district name not set)'}")
    print(f"    {cfg.city} City, {cfg.county} County, Utah")
    print(f"    DS mill levy: {cfg.effective_ds_mill_levy:.3f} "
          f"(cap {cfg.mill_levy_cap:.3f})  |  "
          f"residential taxable ratio: {cfg.resid_taxable_ratio:.2%}  |  "
          f"{cfg.reassess_frequency.lower()} reassessment: {cfg.reassess_rate:.1%}")
    for w in cfg.validate():
        print(f"    ! {w}")
    print(f"    Senior rate {cfg.senior_interest_rate:.2%} ({cfg.dsc_senior:.2f}x)  |  "
          f"Refunding rate {cfg.senior_refunding_interest_rate:.2%} "
          f"({cfg.dsc_refunding:.2f}x)")

    # ── 2. Development & Taxable Value ────────────────────────────────────
    if cfg.absorption_pace_factor != 1.0:
        dev = dev.stressed(cfg.absorption_pace_factor)   # home-sales pacing stress
        print(f"    [Home-sales pace stressed to {cfg.absorption_pace_factor:.0%} of forecast]")
    dev = dev.build(cfg)
    sm = SummaryModel(cfg, dev).build()
    print(f"\n[2] Development: {dev.total_lots} lots, base ASP "
          f"${dev.base_asp:,.0f} @ {cfg.inflation_rate:.0%} inflation")
    print("    Assessed-value build (collection year):")
    _disp = [y for y in range(cfg.first_collection_year, cfg.senior_final_year + 1, 5)
             if sm.row(y)][:6]
    for yr in _disp:
        print(f"      {yr}: AV ${sm.total_av(yr):>14,.0f}  "
              f"net senior rev ${sm.net_senior_revenue(yr):>11,.0f}")

    # ── 3. Senior new-money bonds (revenue-wrap sizing) ────────────────────
    # Optional-redemption provisions: call-protected until the premium-call date,
    # then callable at the premium price (e.g. 103%), declining to par.
    senior_calls = CallProvisions(
        premium_call_date=cfg.premium_call_date,
        par_call_date=cfg.par_call_date,
        premium_call_price=cfg.premium_call_price,
    )
    _size_kwargs = dict(
        name=f"Senior Bonds (Series {cfg.delivery_year}A)",
        rate=cfg.senior_interest_rate,
        coverage=cfg.dsc_senior,
        delivery=cfg.delivery,
        first_principal_year=cfg.senior_first_principal_year,
        final_year=cfg.senior_final_year,
        capi_end_year=cfg.capi_end_date.year,
        call_provisions=senior_calls,
        reoffering_yield=cfg.senior_reoffering_yield,
        coupon_scale=cfg.senior_coupon_scale,
        yield_scale=cfg.senior_yield_scale,
        term_bonds=cfg.senior_term_bonds,
    )
    _sizer = SeniorLienSizer(cfg, sm)
    if cfg.senior_dsrf_deposit is None:
        senior = size_senior_with_dynamic_dsrf(_sizer, **_size_kwargs)
    else:
        senior = _sizer.size(dsrf_deposit=cfg.senior_dsrf_deposit, **_size_kwargs)
    print(f"\n[3] Senior new-money bonds sized: par ${senior.par_amount:,.0f}")
    print(f"    Final maturity {senior.final_year}  |  "
          f"max annual net DS ${senior.max_annual_ds:,.0f}")
    print(f"    DSRF (computed, 3-prong): ${senior.dsrf_deposit:,.0f}")
    print(f"    Optional redemption: call-protected to {senior_calls.premium_call_date} "
          f"→ {senior_calls.premium_call_price:.0f}% premium call "
          f"→ par call {senior_calls.par_call_date}")
    cov = senior_coverage_dataframe(cfg, sm, senior)
    if not cov.empty:
        print(f"    Annual DSC (target {cfg.dsc_senior:.2f}x): "
              f"min {cov['coverage'].min():.2f}x, max {cov['coverage'].max():.2f}x, "
              f"once stabilized ~{cov['coverage'].iloc[-1]:.2f}x")

    # ── 4. Subordinate par — dynamically sized from residual surplus ───────
    first_final = senior.final_year
    _surplus_first = SurplusFund(cfg, sm).build(senior, None, cfg.first_collection_year, first_final)
    if cfg.sub_par is None:
        sub_par = SubordinateLien(cfg, sm).size_par(
            senior, cfg.first_collection_year, first_final, surplus_fund=_surplus_first)
    else:
        sub_par = cfg.sub_par
    print(f"    Subordinate par (sized to residual surplus): ${sub_par:,.0f}")

    # ── 5. Sources & Uses + Reimbursement (first financing) ────────────────
    su = first_financing_sources_uses(cfg, senior, sub_par=sub_par)
    print(f"\n[5] First-financing Sources & Uses (balanced={su.balanced}):")
    for k, v in su.uses.items():
        print(f"      {k:<28} ${v:>13,.0f}")
    print(f"    >>> Developer reimbursement: ${su.reimbursement:,.0f}")

    # ── 6. Refunding — refinance the senior new-money bonds ─────────────────
    # Surplus on hand and the subordinate escrow at the refunding date are
    # derived from the accumulated surplus and the outstanding sub balance.
    refunding = None
    if cfg.refund_financing == "Yes":
        ref_year = cfg.delivery_refunding.year
        if cfg.refunding_surplus_on_hand is None:
            sf_row = {r.year: r for r in _surplus_first.rows}.get(ref_year)
            surplus_on_hand = round(sf_row.reserve_balance, 2) if sf_row else 0.0
        else:
            surplus_on_hand = cfg.refunding_surplus_on_hand
        if cfg.refunding_sub_escrow is None:
            sub_first = SubordinateLien(cfg, sm).size(
                sub_par, senior, cfg.first_collection_year, first_final, surplus_fund=_surplus_first)
            srow = {r["year"]: r for r in sub_first.rows}.get(ref_year)
            sub_escrow = round(((srow["principal_balance"] + srow["accrued_balance"])
                                if srow else 0.0) / 1) if srow else 0.0
        else:
            sub_escrow = cfg.refunding_sub_escrow
        refunding = RefundingAnalysis(cfg, sm).run(
            senior, surplus_on_hand=surplus_on_hand, sub_escrow=sub_escrow)

    # ── 7. Senior surplus fund + subordinate cash-flow lien ────────────────
    # The subordinate lien is the first-financing instrument (sized against this
    # same surplus); the refunding separately escrows it (section 6).
    surplus = _surplus_first
    sub = SubordinateLien(cfg, sm).size(
        sub_par, senior, cfg.first_collection_year, first_final, surplus_fund=surplus)
    print(f"\n[6] Senior surplus fund target: ${surplus.target:,.0f} "
          f"(excess flows to sub lien once full)")
    print(f"\n[7] Subordinate cash-flow note: par ${sub.par_amount:,.0f}  "
          f"({sub.coverage:.2f}x coverage, {sub.rate:.0%} accreting)")
    print(f"    Total debt service ${sub.total_payments:,.0f} "
          f"(interest ${sub.total_interest_paid:,.0f} + "
          f"principal ${sub.total_principal_paid:,.0f})")
    print(f"    Fully repaid: {sub.fully_repaid}  |  "
          f"ending accrued interest ${sub.ending_accrued_interest:,.0f}")

    if refunding is not None:
        print(f"\n[8] REFUNDING — refinance senior new-money bonds:")
        print(f"    Refunding delivered {cfg.delivery_refunding} at "
              f"{refunding.call_price:.0f}% call price (optional redemption)")
        print(f"    Refunding par sized: ${refunding.refunding_bond.par_amount:,.0f} "
              f"@ {cfg.senior_refunding_interest_rate:.2%}")
        print(f"    Callable principal refunded: "
              f"${refunding.refunded_par_outstanding:,.0f}")
        print(f"    Refunding escrow (defeasance): ${refunding.refunding_escrow:,.0f}")
        print(f"    >>> Additional reimbursement (NEW MONEY): "
              f"${refunding.new_money_reimbursement:,.0f}")
        total_reimb = su.reimbursement + refunding.new_money_reimbursement
        print(f"\n    TOTAL developer reimbursement (first + refunding): "
              f"${total_reimb:,.0f}")

    # ── 8. Exports ─────────────────────────────────────────────────────────
    # One flat folder holds the four deliverables: inputs workbook, model
    # output workbook, forecast exhibits, and the reimbursement memo.
    if export:
        os.makedirs(output_dir, exist_ok=True)
        xlsx = build_excel_report(
            cfg, sm, senior, su, refunding, sub, surplus, dev=dev,
            output_path=f"{output_dir}/ut_pid_model_output.xlsx",
        )
        # Forecast exhibits — base case + development stress scenarios (80%/45%).
        scenarios = build_scenarios(cfg, dev, stress_pace_factors=(0.80, 0.45),
                                    sub_par=sub_par)
        forecast_xlsx = build_forecast_report(
            scenarios, output_path=f"{output_dir}/ut_pid_forecast_exhibits.xlsx")
        # Tierra-style reimbursement memo (HTML), populated from the model.
        memo_html = build_memo_html(
            cfg, sm, senior, su, sub, refunding, dev=dev,
            output_path=f"{output_dir}/ut_pid_model_memo.html",
            developer=cfg.developer or "[Developer / Master Developer]")
        print(f"\n[9] Output written to '{output_dir}/':")
        print(f"    Model output:      {xlsx}")
        print(f"    Forecast exhibits: {forecast_xlsx}")
        print(f"    Reimbursement memo:{memo_html}")
        print(f"      scenarios: " + ", ".join(
            f"Exhibit {s.exhibit} ({s.pace_factor:.0%})" for s in scenarios))
        print("\n    Stress-case minimum debt-service coverage (construction era):")
        for s in scenarios:
            net = s.senior.annual_net_ds()
            covs = [s.sm.net_senior_revenue(y) / net[y]
                    for y in sorted(net) if net[y] > 0]
            print(f"      Exhibit {s.exhibit} ({s.pace_factor:>4.0%} pace): "
                  f"min coverage {min(covs):.2f}x")

    print("\n" + "=" * 72)
    print("  Model complete.")
    print("=" * 72)

    return {"cfg": cfg, "dev": dev, "sm": sm, "senior": senior,
            "su": su, "sub": sub, "refunding": refunding}


if __name__ == "__main__":
    run_model()
