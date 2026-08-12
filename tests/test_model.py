"""
Regression tests.

The taxable-value and pledged-revenue figures below are read off the Tierra
"Financial Analysis - Viridian Farms PID (3 MILLS) - Salem_Pricing Day
(Sept 17 2024).xlsm" workbook in this repository.  The Python model reproduces
them exactly, which is the point of the exercise: if a change here moves those
numbers, the model has stopped agreeing with the reference deal.
"""

from __future__ import annotations

import os
import sys
from dataclasses import replace
from datetime import date

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ut_pid_model import ModelConfig, run_model, viridian_farm_projections
from ut_pid_model.config import PID_STATUTORY_LEVY_CAP
from ut_pid_model.xlfin import days360, edate, price, solve_tic, yearfrac

# Assessment year → total taxable value, from the reference workbook's
# Summary!AG column.
REFERENCE_TAXABLE_VALUE = {
    2024: 131_917.50,
    2025: 5_946_325.50,
    2026: 54_266_366.48,
    2027: 109_036_198.05,
    2028: 147_900_823.60,
    2029: 172_718_643.66,
    2030: 195_362_876.98,
    2031: 197_316_505.75,
    2032: 199_289_670.81,
    2033: 201_282_567.52,
    2034: 203_295_393.19,
    2035: 205_328_347.13,
}

# Assessment year → net revenue available for senior debt service (Summary!AX).
REFERENCE_NET_REVENUE = {
    2026: 102_483.12,
    2027: 262_445.22,
    2028: 375_624.80,
    2029: 447_485.12,
    2030: 512_933.01,
    2031: 517_527.99,
    2032: 522_157.45,
    2033: 526_821.49,
}

REFERENCE_SENIOR_PAR = 5_690_000.0


@pytest.fixture(scope="module")
def result():
    return run_model(ModelConfig(), viridian_farm_projections())


# ── Excel primitives ──────────────────────────────────────────────────────────

def test_edate_clamps_to_month_end():
    assert edate(date(2024, 1, 31), 1) == date(2024, 2, 29)
    assert edate(date(2023, 1, 31), 1) == date(2023, 2, 28)
    assert edate(date(2024, 3, 1), -12) == date(2023, 3, 1)


def test_days360_us_convention():
    assert days360(date(2024, 1, 30), date(2024, 2, 29)) == 29
    assert days360(date(2024, 1, 1), date(2025, 1, 1)) == 360
    assert days360(date(2024, 1, 31), date(2024, 3, 31)) == 60


def test_yearfrac_is_thirty_three_sixty():
    assert yearfrac(date(2024, 9, 26), date(2025, 3, 1)) == pytest.approx(0.4305555, abs=1e-6)


def test_price_at_par_when_coupon_equals_yield():
    p = price(date(2024, 9, 26), date(2034, 3, 1), 0.05875, 0.05875)
    assert p == pytest.approx(100.0, abs=0.5)


def test_price_falls_when_yield_exceeds_coupon():
    settle, maturity = date(2024, 9, 26), date(2044, 3, 1)
    assert price(settle, maturity, 0.05, 0.06) < price(settle, maturity, 0.05, 0.05)


def test_solve_tic_recovers_a_known_rate():
    flows = [50.0] * 19 + [1050.0]
    rate = solve_tic(flows, 1000.0, 0.0)
    assert rate == pytest.approx(0.10, abs=1e-6)


# ── Utah statutory framework ──────────────────────────────────────────────────

def test_statutory_levy_cap_is_fifteen_mills():
    assert PID_STATUTORY_LEVY_CAP * 1000 == pytest.approx(15.0)


def test_residential_exemption_is_forty_five_percent():
    assert ModelConfig().resid_taxable_ratio == pytest.approx(0.55)


def test_levy_above_statutory_cap_is_flagged():
    cfg = ModelConfig(mill_levy_governing_doc=16.0)
    assert any("17D-4-303" in w for w in cfg.validate())


def test_levy_within_cap_is_clean():
    assert ModelConfig().validate() == []


def test_gallagherization_is_flagged_as_a_colorado_mechanism():
    cfg = ModelConfig(gallagherization="Yes")
    assert any("Colorado" in w for w in cfg.validate())


def test_principal_falls_in_march():
    cfg = ModelConfig()
    assert cfg.prin_maturity == 3            # Utah taxes are due 30 November
    assert cfg.int_maturity == 9


# ── Tie-out to the reference workbook ─────────────────────────────────────────

@pytest.mark.parametrize("year,expected", sorted(REFERENCE_TAXABLE_VALUE.items()))
def test_taxable_value_matches_reference_workbook(result, year, expected):
    row = next(r for r in result.summary if r.assessment_date.year == year)
    assert row.senior_taxable_value == pytest.approx(expected, rel=1e-6)


@pytest.mark.parametrize("year,expected", sorted(REFERENCE_NET_REVENUE.items()))
def test_net_revenue_matches_reference_workbook(result, year, expected):
    row = next(r for r in result.summary if r.assessment_date.year == year)
    assert row.senior_net_revenue == pytest.approx(expected, rel=1e-5)


def test_senior_par_matches_reference_workbook(result):
    assert result.senior_par == pytest.approx(REFERENCE_SENIOR_PAR)


def test_sizing_converges(result):
    assert result.converged


def test_final_maturity_matches_the_offering(result):
    assert result.senior_stats.final_maturity == date(2054, 3, 1)


# ── Structural invariants ─────────────────────────────────────────────────────

def test_sources_equal_uses(result):
    sources = result.total_par + result.senior_stats.premium
    uses = (result.total_reimbursement + result.surplus_fund_deposit
            + result.capitalized_interest_deposit + result.senior_uwd
            + result.sub_uwd + result.cfg.coi)
    assert sources == pytest.approx(uses, abs=1.0)


def test_principal_sums_to_par(result):
    assert sum(r.principal for r in result.senior) == pytest.approx(result.senior_par)


def test_bond_value_amortises_to_zero(result):
    assert result.senior[-1].bond_value == pytest.approx(0.0, abs=1.0)


def test_coverage_meets_requirement_once_principal_amortises(result):
    cfg = result.cfg
    amortising = [r for r in result.senior
                  if r.principal > 0 and r.payment_date.month == cfg.prin_maturity]
    assert amortising
    for row in amortising:
        assert row.actual_coverage >= cfg.dsc_senior_lien_bonds - 0.01


def test_subordinate_bonds_are_fully_retired(result):
    assert result.summary[-1].sub_principal_balance == pytest.approx(0.0, abs=1.0)
    assert result.summary[-1].sub_accrued_balance == pytest.approx(0.0, abs=1.0)


def test_capi_fund_drains_to_zero(result):
    assert result.capi[-1].ending_balance == pytest.approx(0.0, abs=1.0)


def test_capi_covers_interest_through_the_capi_end_date(result):
    cfg = result.cfg
    for row in result.senior:
        if row.payment_date <= cfg.capi_end_date:
            assert row.capitalized_interest == pytest.approx(row.interest, rel=1e-9)


# ── Utah vs Colorado behaviour ────────────────────────────────────────────────

def test_utah_value_lag_is_one_year():
    assert ModelConfig().value_lag_years == 1


def test_colorado_style_two_year_lag_delays_the_tax_base():
    dev = viridian_farm_projections()
    utah = run_model(ModelConfig(), dev)
    colorado_lag = run_model(replace(ModelConfig(), value_lag_years=2), dev)
    year = 2026
    u = next(r for r in utah.summary if r.assessment_date.year == year)
    c = next(r for r in colorado_lag.summary if r.assessment_date.year == year)
    assert c.senior_taxable_value < u.senior_taxable_value


def test_colorado_assessment_ratio_shrinks_the_base():
    dev = viridian_farm_projections()
    utah = run_model(ModelConfig(), dev)
    colorado = run_model(replace(ModelConfig(), resid_taxable_ratio=0.067,
                                 developed_lot_value=0.29), dev)
    assert colorado.senior_par < utah.senior_par * 0.3


def test_biennial_reassessment_produces_less_value_than_annual():
    dev = viridian_farm_projections()
    annual = run_model(ModelConfig(), dev)
    biennial = run_model(replace(ModelConfig(), reassess_frequency="Biennial"), dev)
    assert biennial.summary[-1].senior_taxable_value <= \
        annual.summary[-1].senior_taxable_value


# ── Sensitivity behaviour ─────────────────────────────────────────────────────

def test_higher_coverage_reduces_senior_par():
    dev = viridian_farm_projections()
    base = run_model(ModelConfig(), dev)
    tight = run_model(replace(ModelConfig(), dsc_senior_lien_bonds=1.50), dev)
    assert tight.senior_par < base.senior_par


def test_higher_mill_levy_raises_capacity():
    dev = viridian_farm_projections()
    base = run_model(ModelConfig(), dev)
    higher = run_model(replace(ModelConfig(), mill_levy_governing_doc=4.0), dev)
    assert higher.senior_par > base.senior_par


def test_slower_absorption_reduces_reimbursement():
    dev = viridian_farm_projections()
    base = run_model(ModelConfig(), dev)
    slow = run_model(replace(ModelConfig(), hypothetical_scenario="Yes",
                             absorption_scenario=0.6, lot_delivery_scenario=0.6), dev)
    assert slow.total_reimbursement < base.total_reimbursement


def test_sub_par_override_is_honoured():
    out = run_model(replace(ModelConfig(), sub_par_override=1_000_000.0),
                    viridian_farm_projections())
    assert out.sub_par == 1_000_000.0


# ── Outputs ───────────────────────────────────────────────────────────────────

EXPECTED_TABS = [
    "Inputs - First", "Capital Costs", "Costs of Issuance",
    "Sources and Uses - First", "Sources and Uses - Refunding", "Summary",
    "Residential Development", "Comm Development", "Ops Rev & Exp Projection",
    "Senior Lien DS - First", "Sub Lien DS - First (Annual)",
    "Sub Lien DS - First (SA)", "Senior Lien DS - Refunding",
    "CAPI Fund - First", "Scratch--->>>", "DBC Output",
]


def test_workbook_has_the_colorado_tab_set(result, tmp_path):
    import openpyxl

    from ut_pid_model.workbook import build_workbook
    path = build_workbook(result, str(tmp_path / "model.xlsx"))
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == EXPECTED_TABS


def test_workbook_carries_colorado_range_aliases(result, tmp_path):
    import openpyxl

    from ut_pid_model.workbook import build_workbook
    path = build_workbook(result, str(tmp_path / "model.xlsx"))
    wb = openpyxl.load_workbook(path)
    names = set(wb.defined_names)
    for utah, colorado in [("RESID_TAXABLE_RATIO", "TABOR_CURRENT"),
                           ("MILL_LEVY_GOVERNING_DOC", "MILL_LEVY_SERVICE_PLAN"),
                           ("UNIFORM_FEE_PRC", "TAX_COLLECT_SO_PRC"),
                           ("ADMIN_COST_BASE", "OM_CARVEOUT")]:
        assert utah in names and colorado in names


def test_memo_is_written_in_both_formats(result, tmp_path):
    from ut_pid_model.memo import build_memo
    md, docx = build_memo(result, str(tmp_path / "memo"))
    assert os.path.getsize(md) > 4_000
    assert os.path.getsize(docx) > 10_000
    text = open(md).read()
    for phrase in ["17D-4", "59-2-103", "45% primary residential exemption",
                   "30 November", "Sensitivity analysis"]:
        assert phrase in text


def test_inputs_template_round_trips(tmp_path):
    from ut_pid_model.inputs import load_inputs, write_template
    path = write_template(str(tmp_path / "inputs.xlsx"))
    cfg, dev = load_inputs(path)
    out = run_model(cfg, dev)
    assert out.senior_par == pytest.approx(REFERENCE_SENIOR_PAR)
    assert dev.total_units() == 716
