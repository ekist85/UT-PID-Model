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
    2024: 131_917.50,
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
    "Summary - Light", "Summary - Detail", "Development Projections",
    "Sources & Uses - First", "Senior Lien DS - First", "Subordinate Lien",
    "Senior Surplus Fund", "CAPI Fund - First", "O&M Revenue",
    "Sources & Uses - Refunding", "Senior Lien DS - Refunding",
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


def test_workbook_carries_no_colorado_labels(deliverables):
    import openpyxl
    out, _ = deliverables
    wb = openpyxl.load_workbook(out / "ut_pid_model_output.xlsx")
    stale = ("TABOR", "Gallagher", "Specific Ownership", "Service Plan",
             "Oil & Gas", "SB24-233", "County Treasurer")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    for term in stale:
                        assert term not in cell.value, f"{ws.title}!{cell.coordinate}: {term}"


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
