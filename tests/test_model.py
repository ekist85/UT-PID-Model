"""
Regression tests for the Utah PID model.

Two things are pinned here:

  * **Utah statute.**  The levy cap, the residential exemption, the tax
    calendar and the reassessment cadence are what make this a Utah model
    rather than the Colorado one; if any of them drift the suite fails.
  * **The reference deal.**  Taxable value and pledged revenue are checked
    against the Tierra pricing-day workbook in this repository, "Financial
    Analysis - Viridian Farms PID (3 MILLS) - Salem_Pricing Day
    (Sept 17 2024).xlsm".
"""

from __future__ import annotations

import os
import sys
from dataclasses import replace
from datetime import date

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ut_pid_model import (CallProvisions, DeveloperProjections, ModelConfig,
                          PID_STATUTORY_LEVY_CAP, RESIDENTIAL_EXEMPTION,
                          RefundingAnalysis, SeniorLienSizer, SubordinateLien,
                          SummaryModel, SurplusFund, build_scenarios,
                          first_financing_sources_uses, lot_inventory_ratio,
                          residential_taxable_ratio_for,
                          size_senior_with_dynamic_dsrf, viridian_farm_projections)

# ── Reference workbook, Summary!AG — total taxable value by ASSESSMENT year.
# The model's `collection_year` is the assessment year + 1.
REFERENCE_TAXABLE_VALUE = {
    2030: 195_362_876.98,
    2031: 197_316_505.75,
    2032: 199_289_670.81,
    2033: 201_282_567.52,
    2034: 203_295_393.19,
    2035: 205_328_347.13,
}

# Reference workbook, Summary!AX — net revenue available for senior debt service.
REFERENCE_NET_REVENUE = {
    2027: 262_445.22,
    2028: 375_624.80,
    2029: 447_485.12,
}

REFERENCE_SENIOR_PAR = 5_690_000.0


@pytest.fixture(scope="module")
def built():
    cfg = ModelConfig()
    dev = viridian_farm_projections().build(cfg)
    sm = SummaryModel(cfg, dev).build()
    return cfg, dev, sm


@pytest.fixture(scope="module")
def senior(built):
    cfg, _dev, sm = built
    calls = CallProvisions(cfg.premium_call_date, cfg.par_call_date,
                           cfg.premium_call_price)
    return size_senior_with_dynamic_dsrf(
        SeniorLienSizer(cfg, sm),
        name="Senior", rate=cfg.senior_interest_rate, coverage=cfg.dsc_senior,
        delivery=cfg.delivery, first_principal_year=cfg.senior_first_principal_year,
        final_year=cfg.senior_final_year, capi_end_year=cfg.capi_end_date.year,
        call_provisions=calls)


# ── Utah statute ──────────────────────────────────────────────────────────────

def test_statutory_levy_cap_is_fifteen_mills():
    assert PID_STATUTORY_LEVY_CAP * 1000 == pytest.approx(15.0)


def test_residential_exemption_is_forty_five_percent():
    assert RESIDENTIAL_EXEMPTION == pytest.approx(0.45)
    assert ModelConfig().resid_taxable_ratio == pytest.approx(0.55)


def test_exemption_history_reaches_45_percent_in_1995():
    assert residential_taxable_ratio_for(1994) > 0.55
    assert residential_taxable_ratio_for(1995) == pytest.approx(0.55)
    assert residential_taxable_ratio_for(2026) == pytest.approx(0.55)


def test_builder_inventory_carries_the_residential_exemption():
    assert lot_inventory_ratio(2025) == pytest.approx(0.55)


def test_controlling_cap_is_the_most_restrictive():
    cfg = ModelConfig(mill_levy_governing_doc=5.0, mill_levy_indenture=3.0)
    assert cfg.mill_levy_cap == pytest.approx(3.0)
    assert ModelConfig(mill_levy_governing_doc=5.0,
                       mill_levy_indenture=None).mill_levy_cap == pytest.approx(5.0)
    assert ModelConfig(mill_levy_governing_doc=40.0,
                       mill_levy_indenture=None).mill_levy_cap == pytest.approx(15.0)


def test_levy_is_held_down_to_the_cap():
    cfg = ModelConfig(mill_levy_ds_target=9.0, mill_levy_indenture=3.0)
    assert cfg.effective_ds_mill_levy == pytest.approx(3.0)
    assert any("controlling cap" in w for w in cfg.validate())


def test_levy_above_the_statute_is_flagged():
    cfg = ModelConfig(mill_levy_ds_target=16.0, mill_levy_governing_doc=20.0,
                      mill_levy_indenture=None)
    assert any("17D-4-303" in w for w in cfg.validate())


def test_default_configuration_is_clean():
    assert ModelConfig().validate() == []


def test_principal_falls_in_march_because_taxes_are_due_in_november():
    cfg = ModelConfig()
    assert cfg.prin_maturity == 3
    assert cfg.int_maturity == 9


def test_call_and_capi_dates_land_on_payment_dates():
    cfg = ModelConfig()
    for d in (cfg.capi_end_date, cfg.premium_call_date, cfg.par_call_date,
              cfg.delivery_refunding):
        assert (d.month, d.day) == (cfg.prin_maturity, cfg.prin_maturity_day_senior)


def test_utah_reassesses_annually():
    cfg = ModelConfig()
    assert cfg.reassess_frequency == "Annual"
    assert DeveloperProjections._is_reassess_year(cfg.first_year + 5, cfg)
    assert DeveloperProjections._is_reassess_year(cfg.first_year + 6, cfg)


def test_biennial_toggle_skips_alternate_rolls():
    cfg = ModelConfig(reassess_frequency="Biennial")
    flags = [DeveloperProjections._is_reassess_year(y, cfg)
             for y in range(cfg.first_year + 4, cfg.first_year + 10)]
    assert flags.count(True) == 3 and flags.count(False) == 3


def test_no_county_treasurer_haircut_by_default():
    assert ModelConfig().county_collection_fee == 0.0


# ── Tie-out to the reference workbook ─────────────────────────────────────────

@pytest.mark.parametrize("assessment_year,expected",
                         sorted(REFERENCE_TAXABLE_VALUE.items()))
def test_taxable_value_matches_reference_workbook(built, assessment_year, expected):
    _cfg, _dev, sm = built
    row = sm.row(assessment_year + 1)
    assert row is not None
    # Lot inventory is carried at the inflated ASP (the current Colorado
    # methodology), where the 2024 workbook used the flat base ASP, so the
    # build-out years differ by a fraction of a percent.
    assert row.total_av == pytest.approx(expected, rel=5e-3)


def test_first_roll_year_is_immaterial_but_tracked(built):
    """
    The first roll (2024) sits ~18% under the workbook — the Colorado value
    build recognises the opening lot inventory differently.  It backs $388 of
    revenue in a year with no debt service, so it is tracked rather than chased.
    """
    _cfg, _dev, sm = built
    row = sm.row(2025)
    assert 100_000 < row.total_av < 140_000
    assert row.net_senior_revenue < 1_000


def test_taxable_value_is_exact_after_buildout(built):
    _cfg, _dev, sm = built
    for year in (2031, 2032, 2033, 2034, 2035):
        assert sm.row(year + 1).total_av == pytest.approx(
            REFERENCE_TAXABLE_VALUE[year], rel=1e-6)


@pytest.mark.parametrize("assessment_year,expected",
                         sorted(REFERENCE_NET_REVENUE.items()))
def test_net_revenue_matches_reference_workbook(built, assessment_year, expected):
    _cfg, _dev, sm = built
    assert sm.row(assessment_year + 1).net_senior_revenue == pytest.approx(
        expected, rel=5e-3)


def test_senior_par_matches_reference_workbook(senior):
    assert senior.par_amount == pytest.approx(REFERENCE_SENIOR_PAR, rel=5e-3)


def test_final_maturity_matches_the_offering(senior):
    assert senior.final_year == 2054


# ── Structural invariants ─────────────────────────────────────────────────────

def test_principal_sums_to_par(senior):
    assert sum(p.principal for p in senior.schedule) == pytest.approx(senior.par_amount)


def test_coverage_meets_target_once_principal_amortises(built, senior):
    cfg, _dev, sm = built
    net = senior.annual_net_ds()
    amortising = [p.payment_date.year for p in senior.schedule if p.principal > 0]
    for y in amortising:
        assert sm.net_senior_revenue(y) / net[y] >= cfg.dsc_senior - 0.02


def test_capitalized_interest_covers_the_capi_years(built, senior):
    cfg, _dev, _sm = built
    for p in senior.schedule:
        if p.payment_date.year <= cfg.capi_end_date.year:
            assert p.capitalized_interest == pytest.approx(p.interest)


def test_sources_equal_uses(built, senior):
    cfg, _dev, sm = built
    sub_par = SubordinateLien(cfg, sm).size_par(
        senior, cfg.first_collection_year, senior.final_year)
    su = first_financing_sources_uses(cfg, senior, sub_par=sub_par)
    assert su.balanced


def test_subordinate_par_is_fully_repaid(built, senior):
    cfg, _dev, sm = built
    sf = SurplusFund(cfg, sm).build(senior, None, cfg.first_collection_year,
                                    senior.final_year)
    par = SubordinateLien(cfg, sm).size_par(
        senior, cfg.first_collection_year, senior.final_year, surplus_fund=sf)
    res = SubordinateLien(cfg, sm).size(
        par, senior, cfg.first_collection_year, senior.final_year, surplus_fund=sf)
    assert par > 0
    assert res.fully_repaid
    assert res.ending_accrued_interest == pytest.approx(0.0, abs=1.0)


def test_refunding_defeases_the_callable_principal(built, senior):
    cfg, _dev, sm = built
    res = RefundingAnalysis(cfg, sm).run(senior, surplus_on_hand=0.0, sub_escrow=0.0)
    assert res.call_price == pytest.approx(103.0)
    assert res.refunding_escrow == pytest.approx(
        res.refunded_par_outstanding * 1.03)


def test_refunding_issues_its_own_subordinate_lien(built, senior):
    """Sized by the same method as the new-money sub, dated on the refunding
    delivery, and fully repaid by its final maturity."""
    cfg, _dev, sm = built
    res = RefundingAnalysis(cfg, sm).run(senior, surplus_on_hand=0.0, sub_escrow=0.0)
    assert res.refunding_sub is not None
    assert res.refunding_sub_par > 0
    assert res.refunding_sub.par_amount == pytest.approx(res.refunding_sub_par)
    assert res.refunding_sub.fully_repaid
    assert res.refunding_sub_par % 1000 == 0          # floored to the $1,000


def test_refunding_sub_first_coupon_is_a_stub_from_the_refunding_delivery(built, senior):
    """Dated 1 March, first pays 15 March — a 14-day stub, not a full year."""
    cfg, _dev, sm = built
    res = RefundingAnalysis(cfg, sm).run(senior, surplus_on_hand=0.0, sub_escrow=0.0)
    first = res.refunding_sub.rows[0]
    assert first["year"] == cfg.delivery_refunding.year
    implied = first["current_interest"] / (res.refunding_sub_par * cfg.sub_interest_rate)
    assert implied == pytest.approx(14 / 360, abs=1e-4)


def test_refunding_new_money_depends_on_the_refunding_sub(built, senior):
    """At 3 mills the senior refunding alone returns LESS than it costs once the
    new-money sub is defeased; the refunding subordinate lien is what makes the
    refunding pay.  If this ever flips, the 'new money' headline needs rereading."""
    cfg, _dev, sm = built
    ref_year = cfg.delivery_refunding.year
    final = senior.final_year
    sf = SurplusFund(cfg, sm).build(senior, None, cfg.first_collection_year, final)
    sub_par = SubordinateLien(cfg, sm).size_par(
        senior, cfg.first_collection_year, final, surplus_fund=sf)
    sub = SubordinateLien(cfg, sm).size(
        sub_par, senior, cfg.first_collection_year, final, surplus_fund=sf)
    srow = {r["year"]: r for r in sub.rows}[ref_year]
    res = RefundingAnalysis(cfg, sm).run(
        senior,
        surplus_on_hand={r.year: r for r in sf.rows}[ref_year].reserve_balance,
        sub_escrow=srow["principal_balance"] + srow["accrued_balance"])
    # Dropping the refunding sub removes its par from sources and its
    # underwriter's discount from uses.
    senior_only = res.new_money_reimbursement - res.refunding_sub_par * (1 - cfg.uwd_sub)
    assert senior_only < 0, senior_only
    assert res.new_money_reimbursement > 0


# ── Utah vs Colorado behaviour ────────────────────────────────────────────────

def test_utah_base_dwarfs_the_colorado_base_for_the_same_homes(built):
    cfg, dev, sm = built
    colorado = SummaryModel(
        replace(cfg, resid_taxable_ratio=0.067, lot_inventory_taxable_ratio=0.29,
                lot_inventory_rate_schedule={1995: 0.29}),
        dev).build()
    ratio = sm.total_av(2032) / colorado.total_av(2032)
    assert 7.0 < ratio < 9.0        # 55% vs 6.7% on a nearly all-residential roll


def test_three_mills_in_utah_beats_sixty_in_colorado_on_capacity(built):
    """A 3-mill Utah levy raises more than a 60-mill Colorado levy would."""
    cfg, dev, sm = built
    colorado = SummaryModel(
        replace(cfg, resid_taxable_ratio=0.067, lot_inventory_taxable_ratio=0.29,
                lot_inventory_rate_schedule={1995: 0.29},
                mill_levy_ds_target=60.0, mill_levy_governing_doc=60.0,
                mill_levy_indenture=None), dev).build()
    assert sm.row(2032).mill_revenue > colorado.row(2032).mill_revenue


def test_summary_stops_at_the_value_build_horizon(built):
    """No tail of zero taxable value past the senior final maturity."""
    cfg, _dev, sm = built
    # The builds are keyed by roll year and collected the next, so the last row
    # is the collection of the final-maturity roll.
    assert sm.rows[-1].collection_year <= cfg.senior_final_year + 1
    assert all(r.total_av > 0 for r in sm.rows if r.collection_year >= 2026)


def test_biennial_reassessment_lands_below_annual(built):
    cfg, _dev, _sm = built
    annual = SummaryModel(cfg, viridian_farm_projections().build(cfg)).build()
    biennial_cfg = replace(cfg, reassess_frequency="Biennial")
    biennial = SummaryModel(
        biennial_cfg, viridian_farm_projections().build(biennial_cfg)).build()
    assert biennial.rows[-1].total_av < annual.rows[-1].total_av


# ── Sensitivity behaviour ─────────────────────────────────────────────────────

def _size(cfg, dev=None):
    dev = (dev or viridian_farm_projections()).build(cfg)
    sm = SummaryModel(cfg, dev).build()
    calls = CallProvisions(cfg.premium_call_date, cfg.par_call_date,
                           cfg.premium_call_price)
    return size_senior_with_dynamic_dsrf(
        SeniorLienSizer(cfg, sm),
        name="S", rate=cfg.senior_interest_rate, coverage=cfg.dsc_senior,
        delivery=cfg.delivery, first_principal_year=cfg.senior_first_principal_year,
        final_year=cfg.senior_final_year, capi_end_year=cfg.capi_end_date.year,
        call_provisions=calls)


def test_higher_coverage_reduces_par(senior):
    assert _size(ModelConfig(dsc_senior=1.50)).par_amount < senior.par_amount


def test_raising_the_levy_within_the_cap_raises_capacity(senior):
    higher = ModelConfig(mill_levy_ds_target=4.0, mill_levy_indenture=4.0)
    assert _size(higher).par_amount > senior.par_amount


def test_raising_the_target_above_the_cap_changes_nothing(senior):
    capped = ModelConfig(mill_levy_ds_target=9.0, mill_levy_indenture=3.0)
    assert _size(capped).par_amount == pytest.approx(senior.par_amount)


def test_slower_absorption_delays_the_tax_base(built):
    """
    Stress cases are about *coverage during construction*, not par.

    Slowing absorption pushes taxable value out several years, which is what
    breaks coverage while the bonds are outstanding.  It does not necessarily
    shrink the bond: total units are preserved and homes close later at
    inflated prices, so the stabilized base can end up slightly larger.  That is
    why the Colorado model sizes once in the base case and tests the fixed debt
    service against the stressed revenue — the same convention is kept here.
    """
    cfg, _dev, sm = built
    slow = SummaryModel(cfg, viridian_farm_projections().stressed(0.6).build(cfg)).build()
    for year in (2028, 2029, 2030, 2031):
        assert slow.total_av(year) < sm.total_av(year)


def test_stress_cases_break_coverage_during_construction(built, senior):
    cfg, dev, _sm = built
    scenarios = build_scenarios(cfg, dev, stress_pace_factors=(0.80, 0.45),
                                sub_par=1_000_000)
    mins = []
    for s in scenarios:
        net = s.senior.annual_net_ds()
        mins.append(min(s.sm.net_senior_revenue(y) / net[y]
                        for y in sorted(net) if net[y] > 0))
    assert mins[0] > mins[1] > mins[2]


def test_stress_scenarios_step_down_in_pace(built):
    cfg, dev, _sm = built
    scenarios = build_scenarios(cfg, dev, stress_pace_factors=(0.80, 0.45),
                                sub_par=1_000_000)
    assert [s.pace_factor for s in scenarios] == [1.0, 0.80, 0.45]
    assert [s.exhibit for s in scenarios] == ["A", "B", "C"]


# ── Outputs ───────────────────────────────────────────────────────────────────

EXPECTED_TABS = [
    "Summary - Light", "Summary - Detail", "Builder Lot Inventory Value",
    "Residential Value", "Development Projections",
    "Sources & Uses - First", "Senior Lien DS - First", "Subordinate Lien",
    "Senior Surplus Fund", "CAPI Fund - First", "O&M Revenue",
    "Sources & Uses - Refunding", "Senior Lien DS - Refunding",
    "Subordinate Lien - Refunding",
    "Senior Lien Coverage", "Call Schedule", "Notes",
]


@pytest.fixture(scope="module")
def deliverables(tmp_path_factory):
    out = tmp_path_factory.mktemp("deliverables")
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from main import run_model
    result = run_model(export=True, output_dir=str(out))
    return out, result


def test_workbook_has_the_colorado_model_tab_set(deliverables):
    import openpyxl
    out, _ = deliverables
    wb = openpyxl.load_workbook(out / "ut_pid_model_output.xlsx")
    assert wb.sheetnames == EXPECTED_TABS


def test_forecast_exhibits_cover_three_scenarios(deliverables):
    import openpyxl
    out, _ = deliverables
    wb = openpyxl.load_workbook(out / "ut_pid_forecast_exhibits.xlsx")
    names = " ".join(wb.sheetnames)
    for prefix in ("A", "B", "C"):
        assert f"Exhibit {prefix}-1" in names or f"{prefix}-1" in names


def test_memo_states_the_utah_framework(deliverables):
    out, _ = deliverables
    html = (out / "ut_pid_model_memo.html").read_text()
    for phrase in ("17D-4", "59-2-103", "45% exemption", "30&nbsp;November",
                   "R884-24P-52", "59-2-405", "59-2-1602"):
        assert phrase in html, phrase
    for stale in ("Gallagheriz", "TABOR", "Specific Ownership"):
        # The memo may explain that Utah has no gallagherization; it must not
        # claim one is applied.
        assert f"is &ldquo;{stale}" not in html


def test_memo_uses_of_funds_splits_by_series(deliverables):
    """The per-series Uses table must foot to each series' par, and to the total."""
    import re
    out, result = deliverables
    html = (out / "ut_pid_model_memo.html").read_text()
    block = re.search(r"Uses of Funds.*?</table>", html, re.S)
    assert block, "memo has no Uses of Funds table"
    body = block.group(0)
    assert "Senior Lien Bonds" in body and "Subordinate Lien" in body
    total_row = re.search(r"Total Uses.*?</tr>", body, re.S).group(0)
    amounts = [int(x.replace(",", "")) for x in re.findall(r"\$([\d,]+)", total_row)]
    assert len(amounts) == 3, amounts                       # senior, sub, total
    senior_par = result["senior"].par_amount
    sub_par = result["sub"].par_amount
    assert amounts[0] == pytest.approx(senior_par + result["senior"].total_premium, abs=1)
    assert amounts[1] == pytest.approx(sub_par, abs=1)
    assert amounts[2] == pytest.approx(amounts[0] + amounts[1], abs=1)


def test_memo_prints_negative_amounts_in_parentheses():
    """Utah's fixed caps let a refunding return less than it costs — the memo
    has to read as ($914,919), not $-914,919."""
    from ut_pid_model.memo import _money
    assert _money(-914_919) == "($914,919)"
    assert _money(914_919) == "$914,919"


def test_subordinate_first_coupon_is_a_stub_from_the_dated_date(built, senior):
    """The note is dated 26 September and first pays 15 March, in the NEXT
    calendar year.  Upstream would charge a full year; Utah gets a stub."""
    from ut_pid_model import SubordinateLien, SurplusFund
    cfg, _dev, sm = built
    final = senior.final_year
    sf = SurplusFund(cfg, sm).build(senior, None, cfg.first_collection_year, final)
    par = SubordinateLien(cfg, sm).size_par(
        senior, cfg.first_collection_year, final, surplus_fund=sf)
    sub = SubordinateLien(cfg, sm).size(
        par, senior, cfg.first_collection_year, final, surplus_fund=sf)
    rows = {r["year"]: r for r in sub.rows}
    yf = {y: (r["current_interest"] / (r["principal_balance"] * cfg.sub_interest_rate)
              if r["principal_balance"] else 0.0)
          for y, r in rows.items()}
    first_pay = cfg.delivery.year + 1                       # 15 March 2025
    assert yf[cfg.delivery.year] == pytest.approx(0.0)      # not yet dated
    # 30/360 from 2024-09-26 to 2025-03-15 = 169/360.
    assert yf[first_pay] == pytest.approx(169 / 360, abs=1e-4)
    assert yf[first_pay + 1] == pytest.approx(1.0)


# ── District O&M expense ─────────────────────────────────────────────────────

def _sized(cfg):
    """(senior, sub_par, sm) for a config — used by the O&M sensitivity tests."""
    dev = viridian_farm_projections().build(cfg)
    sm = SummaryModel(cfg, dev).build()
    calls = CallProvisions(cfg.premium_call_date, cfg.par_call_date,
                           cfg.premium_call_price)
    senior = size_senior_with_dynamic_dsrf(
        SeniorLienSizer(cfg, sm),
        name="Senior", rate=cfg.senior_interest_rate, coverage=cfg.dsc_senior,
        delivery=cfg.delivery, first_principal_year=cfg.senior_first_principal_year,
        final_year=cfg.senior_final_year, capi_end_year=cfg.capi_end_date.year,
        call_provisions=calls)
    sf = SurplusFund(cfg, sm).build(senior, None, cfg.first_collection_year,
                                    senior.final_year)
    sub_par = SubordinateLien(cfg, sm).size_par(
        senior, cfg.first_collection_year, senior.final_year, surplus_fund=sf)
    return senior, sub_par, sm


def test_om_expense_defaults_to_zero_and_changes_nothing():
    """An operating budget is district-specific; the model must not invent one."""
    cfg = ModelConfig()
    assert cfg.om_expense == 0.0
    assert all(r.om_expense == 0.0 for r in SummaryModel(
        cfg, viridian_farm_projections().build(cfg)).build().rows)


def test_om_expense_starts_with_the_other_district_costs_and_inflates():
    cfg = ModelConfig()
    cfg.om_expense = 40_000
    cfg.om_growth_rate = 0.035
    start = cfg.district_cost_start_year or (cfg.delivery.year + 2)
    assert cfg.om_expense_for(start - 1) == 0.0
    assert cfg.om_expense_for(start) == pytest.approx(40_000)
    assert cfg.om_expense_for(start + 10) == pytest.approx(40_000 * 1.035 ** 10)


def test_om_expense_is_netted_from_both_liens():
    """The sub lien's own revenue is `net_sub - net_senior`, so a cost netted
    from the senior side alone would be handed to the sub — an O&M expense would
    then *raise* subordinate capacity.  It has to come off the top."""
    base = SummaryModel(ModelConfig(),
                        viridian_farm_projections().build(ModelConfig())).build()
    cfg = ModelConfig()
    cfg.om_expense = 40_000
    with_om = SummaryModel(cfg, viridian_farm_projections().build(cfg)).build()
    year = 2035
    charge = with_om.row(year).om_expense
    assert charge > 0
    assert (base.net_senior_revenue(year) - with_om.net_senior_revenue(year)
            == pytest.approx(charge))
    assert (base.net_sub_revenue(year) - with_om.net_sub_revenue(year)
            == pytest.approx(charge))


def test_om_expense_reduces_both_senior_and_subordinate_capacity():
    cfg_om = ModelConfig()
    cfg_om.om_expense = 40_000
    senior_0, sub_0, _ = _sized(ModelConfig())
    senior_om, sub_om, _ = _sized(cfg_om)
    assert senior_om.par_amount < senior_0.par_amount
    assert sub_om < sub_0


def test_inputs_workbook_carries_the_om_expense_rows(tmp_path):
    import openpyxl
    from ut_pid_model import load_inputs_workbook, write_inputs_workbook
    path = write_inputs_workbook(output_path=str(tmp_path / "inputs.xlsx"))
    wb = openpyxl.load_workbook(path)
    ws = wb["Inputs"]
    cells = {}
    for row in ws.iter_rows(min_row=5, max_col=4):
        if row[3].value in ("OM_EXPENSE", "OM_GROWTH_RATE"):
            cells[row[3].value] = row[2]
    assert set(cells) == {"OM_EXPENSE", "OM_GROWTH_RATE"}
    cells["OM_EXPENSE"].value = 40_000
    cells["OM_GROWTH_RATE"].value = 0.035
    wb.save(path)
    cfg, _dev = load_inputs_workbook(path)
    assert cfg.om_expense == pytest.approx(40_000)
    assert cfg.om_growth_rate == pytest.approx(0.035)


def test_om_tab_shows_expense_against_revenue(deliverables):
    import openpyxl
    out, _ = deliverables
    ws = openpyxl.load_workbook(out / "ut_pid_model_output.xlsx")["O&M Revenue"]
    hdrs = [ws.cell(row=5, column=c).value for c in range(1, 9)]
    assert "− O&M\nExpense" in hdrs
    assert "O&M Surplus /\n(Deficit)" in hdrs


def test_summary_detail_shows_the_om_column(deliverables):
    import openpyxl
    out, _ = deliverables
    ws = openpyxl.load_workbook(out / "ut_pid_model_output.xlsx")["Summary - Detail"]
    hdrs = [c.value for c in ws[5]]
    assert "− District\nO&M" in hdrs


def test_workbook_carries_no_colorado_labels(deliverables):
    import openpyxl
    out, _ = deliverables
    wb = openpyxl.load_workbook(out / "ut_pid_model_output.xlsx")
    stale = ("TABOR", "Gallagher", "Specific Ownership", "Service Plan",
             "Oil & Gas", "SB24-233", "County Treasurer", "Metro District",
             "Assessed Valuation", "Biennial reassessment",
             "biennial reassessment", "Specific\nOwnership",
             "County\nTreasurer")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    for term in stale:
                        assert term not in cell.value, f"{ws.title}!{cell.coordinate}: {term}"


def test_no_colorado_identifiers_survive_in_the_package():
    """Locals and dict keys are part of the port too — `sot`, `tabor` and
    `treasurer_fee` name Colorado mechanisms Utah does not have."""
    import glob
    import re
    pkg = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "ut_pid_model", "*.py")
    banned = re.compile(r"\b(sot|tabor|treasurer_fee|gallagherize)\b")
    offenders = []
    for path in glob.glob(pkg):
        for n, line in enumerate(open(path, encoding="utf-8"), 1):
            if banned.search(line):
                offenders.append(f"{os.path.basename(path)}:{n}: {line.strip()}")
    assert not offenders, offenders


def test_inputs_workbook_round_trips(tmp_path):
    from ut_pid_model import load_inputs_workbook, write_inputs_workbook
    path = write_inputs_workbook(
        output_path=str(tmp_path / "inputs.xlsx"),
        cfg=ModelConfig(), dev=viridian_farm_projections())
    cfg, dev = load_inputs_workbook(path)
    assert cfg.pid_name == ModelConfig().pid_name
    assert cfg.effective_ds_mill_levy == pytest.approx(3.0)
    assert cfg.resid_taxable_ratio == pytest.approx(0.55)
    assert sum(p.total_units for p in dev.products) == 716
    sm = SummaryModel(cfg, dev.build(cfg)).build()
    assert sm.total_av(2032) == pytest.approx(REFERENCE_TAXABLE_VALUE[2031], rel=1e-6)


def test_inputs_workbook_has_the_utah_reference_tab(tmp_path):
    import openpyxl
    from ut_pid_model import write_inputs_workbook
    path = write_inputs_workbook(output_path=str(tmp_path / "inputs.xlsx"))
    wb = openpyxl.load_workbook(path)
    assert "Utah Property Tax Reference" in wb.sheetnames
    text = " ".join(str(c.value) for row in wb["Utah Property Tax Reference"].iter_rows()
                    for c in row if c.value)
    for phrase in ("59-2-103", "R884-24P-52", "17D-4-303", "30 November"):
        assert phrase in text, phrase


# ── Port fidelity ─────────────────────────────────────────────────────────────

CO_REPO = os.path.join(os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))), "CO-Metro-District-Model")


@pytest.mark.skipif(not os.path.isdir(os.path.join(CO_REPO, "co_metro_model")),
                    reason="Colorado model not checked out alongside this repo")
def test_port_is_up_to_date_with_colorado():
    """
    ``ut_pid_model`` is a port of ``co_metro_model``, not a fork.  If the
    Colorado checkout alongside this one has moved, re-run

        python tools/port_from_colorado.py --source ../CO-Metro-District-Model

    and re-verify the tie-out.  A failure here is the signal, not a defect.
    """
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))), "tools"))
    import port_from_colorado as port
    rc = port.main(["--source", CO_REPO, "--check"])
    assert rc == 0, "port patches no longer apply — Colorado has moved"
