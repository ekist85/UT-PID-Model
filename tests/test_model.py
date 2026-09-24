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
                          allocate_contribution,
                          PID_STATUTORY_LEVY_CAP, RESIDENTIAL_EXEMPTION,
                          RefundingAnalysis, SeniorLienSizer, SubordinateLien,
                          SummaryModel, SurplusFund, build_scenarios,
                          first_financing_sources_uses, lot_inventory_ratio,
                          residential_taxable_ratio_for, senior_coverage_dataframe,
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
        final_year=cfg.senior_final_year, capi_end=cfg.capi_end_date,
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


def test_call_dates_land_on_principal_dates():
    """Bonds are redeemed on a principal payment date, not on an anniversary of
    closing."""
    cfg = ModelConfig()
    for d in (cfg.premium_call_date, cfg.par_call_date, cfg.delivery_refunding):
        assert (d.month, d.day) == (cfg.prin_maturity, cfg.prin_maturity_day_senior)


def test_capi_end_date_is_the_anniversary_of_the_dated_date():
    """36 months from 9/30/2026 is 9/30/2029 — at EITHER frequency.  The end of
    the CAPI period is the anniversary of the dated date and nothing else; it is
    not snapped to a coupon.  Snapping back to the principal month gave 3/1/2029
    (29 months), and snapping to the nearest coupon still gives 3/1/2029 on an
    annual-pay bond, which is the same error wearing a different hat."""
    from ut_pid_model.config import _edate
    for freq in ("Semiannual", "Annual"):
        cfg = ModelConfig(delivery=date(2026, 9, 30), capi_term=36,
                          interest_frequency=freq)
        assert cfg.capi_end_date == date(2029, 9, 30), freq
    # Always exactly the term, whatever the dated date and whatever the term.
    for delivery in (date(2026, 1, 15), date(2026, 5, 31), date(2026, 9, 30),
                     date(2026, 12, 1)):
        for term in (0, 18, 24, 36):
            c = ModelConfig(delivery=delivery, capi_term=term)
            assert c.capi_end_date == _edate(delivery, term), (delivery, term)


@pytest.mark.parametrize("freq", ["Annual", "Semiannual"])
def test_capi_funds_the_whole_term_not_just_the_coupons_inside_it(freq):
    """A 36-month CAPI fund carries the district for 36 months.

    An annual-pay bond dated 9/30/2026 pays 3/1/2027, 3/1/2028 and 3/1/2029
    inside the period — only 29 months of accrual — so funding just those left
    the district paying the 3/1/2030 coupon in full, seven months of which
    belong to the CAPI period.  The fund now pays interest ACCRUED through
    9/30/2029: the straddling coupon is split and the deposit equals exactly
    three years of interest on the par."""
    cfg = ModelConfig(delivery=date(2026, 9, 30), capi_term=36,
                      interest_frequency=freq)
    dev = viridian_farm_projections().build(cfg)
    sm = SummaryModel(cfg, dev).build()
    senior = size_senior_with_dynamic_dsrf(
        SeniorLienSizer(cfg, sm),
        name="Senior", rate=cfg.senior_interest_rate, coverage=cfg.dsc_senior,
        delivery=cfg.delivery, first_principal_year=cfg.senior_first_principal_year,
        final_year=cfg.senior_final_year, capi_end=cfg.capi_end_date,
        call_provisions=CallProvisions(cfg.premium_call_date, cfg.par_call_date,
                                       cfg.premium_call_price))
    rows = [p for p in senior.schedule if p.capitalized_interest]
    # Every coupon inside the period is funded in full; the one straddling the
    # anniversary is funded in part; nothing past it is funded at all.
    whole = [p for p in rows if p.payment_date <= cfg.capi_end_date]
    part = [p for p in rows if p.payment_date > cfg.capi_end_date]
    assert whole and all(
        p.capitalized_interest == pytest.approx(p.interest) for p in whole)
    assert len(part) == 1 and part[0].payment_date == date(2030, 3, 1)
    assert 0 < part[0].capitalized_interest < part[0].interest
    # Three years of interest on the par, to the cent — no more, no less.
    deposit = sum(p.capitalized_interest for p in senior.schedule)
    assert deposit == pytest.approx(
        senior.par_amount * cfg.senior_interest_rate * 3, rel=1e-9)


def test_final_maturity_is_the_last_principal_date(senior):
    """Not the last coupon.  In date order a Utah bond ends on the September
    coupon, which is not a maturity."""
    last_pay = max(p.payment_date for p in senior.schedule)
    last_prin = max(p.payment_date for p in senior.schedule if p.principal)
    assert last_prin.month == 3
    assert last_pay.month == 9 and last_pay > last_prin
    assert senior.final_year == last_prin.year


def test_capitalised_coupons_are_decided_by_date_not_by_year(built, senior):
    """Coupons on or before the CAPI end date are funded in full; at most one
    coupon past it is funded, and only in part; everything beyond that is the
    district's."""
    cfg, _dev, _sm = built
    rows = senior.schedule
    capi = [p for p in rows if p.capitalized_interest]
    assert capi, "nothing capitalized"
    assert all(p.capitalized_interest == pytest.approx(p.interest)
               for p in capi if p.payment_date <= cfg.capi_end_date)
    after = sorted((p for p in rows if p.payment_date > cfg.capi_end_date),
                   key=lambda p: p.payment_date)
    assert after
    assert after[0].capitalized_interest < after[0].interest
    assert all(p.capitalized_interest == 0 for p in after[1:])


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
    # Within ~1%.  The model sizes to exactly its 1.30x target; the workbook's
    # hand-built structure sits at roughly 1.31x on net debt service, so it
    # carries slightly less par for the same revenue.
    assert senior.par_amount == pytest.approx(REFERENCE_SENIOR_PAR, rel=1.2e-2)


def test_first_coupon_is_a_stub_accruing_from_the_dated_date(built, senior):
    """The reference workbook's first senior coupon is $143,929.34 against
    $167,143.75 thereafter — 155/180, the 30/360 days from the 9/26/2024 dating
    to 3/1/2025.  Interest accrues from the dating, not a full half-year."""
    from ut_pid_model.pricing import days_30_360
    cfg, _dev, _sm = built
    rows = sorted(senior.schedule, key=lambda p: p.payment_date)
    first, second = rows[0], rows[1]
    expected = days_30_360(cfg.delivery, first.payment_date) / 180.0
    assert first.interest / second.interest == pytest.approx(expected, rel=1e-9)
    assert first.interest < second.interest


def test_payments_run_in_date_order_and_interest_follows_the_balance(senior):
    """Utah pays principal in March and its other coupon in September, so the
    September coupon is charged on the balance March has already paid down —
    the workbook drops by exactly principal x coupon / 2."""
    rows = senior.schedule
    assert rows == sorted(rows, key=lambda p: p.payment_date)
    for i, p in enumerate(rows[:-1]):
        if p.principal and p.payment_date.month == 3:
            nxt = rows[i + 1]
            assert nxt.payment_date.month == 9
            drop = p.interest - nxt.interest
            assert drop == pytest.approx(p.principal * senior.rate / 2.0, rel=1e-6)
            break


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


def test_refunding_new_money_is_mostly_the_refunding_sub(built, senior):
    """Most of the refunding's 'new money' is the refunding subordinate lien,
    not an interest saving on the senior — worth pinning, because the headline
    reads like a rate story and is not one."""
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
    assert res.new_money_reimbursement > 0
    assert senior_only < res.new_money_reimbursement / 2, senior_only


def test_series_c_is_off_and_invisible(tmp_path):
    """Series C ports across as code but stays inert: toggle off, no Inputs rows,
    no tab, no memo row.  Its Utah statutory basis has not been worked through."""
    import openpyxl
    from ut_pid_model import write_inputs_workbook
    assert ModelConfig().size_series_c == "No"
    path = write_inputs_workbook(output_path=str(tmp_path / "inputs.xlsx"))
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    assert "Series C" not in cell.value, f"{ws.title}!{cell.coordinate}"
                    assert "SERIES_C" not in cell.value, f"{ws.title}!{cell.coordinate}"


def test_developer_contribution_is_a_source_and_defaults_to_zero():
    from ut_pid_model import allocate_contribution
    cfg = ModelConfig()
    assert cfg.developer_contribution == 0.0
    split = allocate_contribution(0.0, cfg.developer_contribution_series,
                                  5_665_000, 1_730_000, 0.0)
    assert sum(split.values()) == 0.0
    split = allocate_contribution(100_000, "Proportional", 5_665_000, 1_730_000, 0.0)
    assert sum(split.values()) == pytest.approx(100_000)
    assert split["series_c"] == 0.0


def test_lot_inventory_follows_deliveries_not_closings():
    """A lot platted in year y-1 is on the roll set 1 January of year y and is
    assessed whether or not a home has been built on it.  Upstream keyed the
    whole tab off home closings, so a builder holding delivered lots showed $0
    of inventory."""
    from ut_pid_model.development import ProductLine
    cfg = ModelConfig()
    p = ProductLine("Test", 100, {2025: 100}, {2027: 25, 2028: 25, 2029: 25, 2030: 25},
                    500_000, 2026)
    dev = DeveloperProjections(products=[p]).build(cfg)
    rows = {r["av_set"]: r for r in dev.lot_inventory_value_build(cfg)}
    # 100 lots delivered in 2025 land on the 2026 roll, before any home closes.
    assert rows[2026]["new_lots"] == pytest.approx(dev.lot_market_value[2025])
    assert rows[2026]["cumulative"] > 0
    # They stay in inventory while no homes close, then draw down as they do.
    assert rows[2027]["cumulative"] == pytest.approx(rows[2026]["cumulative"])
    for y in (2028, 2029, 2030, 2031):
        assert rows[y]["cumulative"] < rows[y - 1]["cumulative"]
    assert rows[2031]["cumulative"] == pytest.approx(0.0, abs=1.0)


def test_lot_inventory_cumulative_is_the_inventory_actually_held(built):
    """The running cumulative telescopes to lots delivered minus homes closed,
    lagged one year onto the roll — so the tab means what its headers say."""
    cfg, dev, _sm = built
    for r in dev.lot_inventory_value_build(cfg):
        assert r["cumulative"] == pytest.approx(
            dev.vacant_lot_market_value(r["av_set"] - 1), abs=1.0), r["av_set"]


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


def test_summary_runs_the_projection_horizon_with_no_zero_tail(built):
    """The projection runs PROJECTION_YEARS from delivery — never past the value
    builds into a tail of zero taxable value and fee-only negative revenue."""
    cfg, dev, sm = built
    horizon = max(cfg.delivery.year + cfg.projection_years, cfg.senior_final_year)
    assert dev.last_year == horizon
    assert sm.rows[-1].collection_year == horizon
    assert all(r.total_av > 0 for r in sm.rows if r.collection_year >= 2026)
    assert sm.rows[-1].net_senior_revenue > 0


def test_projection_horizon_never_truncates_bond_sizing(senior):
    """A short display horizon must not shorten the revenue a bond sizes
    against — the horizon floors at the longest bond maturity."""
    cfg = ModelConfig(projection_years=5)
    dev = viridian_farm_projections().build(cfg)
    assert dev.last_year >= cfg.senior_final_year
    sm = SummaryModel(cfg, dev).build()
    calls = CallProvisions(cfg.premium_call_date, cfg.par_call_date,
                           cfg.premium_call_price)
    short = size_senior_with_dynamic_dsrf(
        SeniorLienSizer(cfg, sm),
        name="Senior", rate=cfg.senior_interest_rate, coverage=cfg.dsc_senior,
        delivery=cfg.delivery, first_principal_year=cfg.senior_first_principal_year,
        final_year=cfg.senior_final_year, capi_end=cfg.capi_end_date,
        call_provisions=calls)
    assert short.par_amount == pytest.approx(senior.par_amount)


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
        final_year=cfg.senior_final_year, capi_end=cfg.capi_end_date,
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


def _deliverable(out, result, suffix=""):
    """Deliverables are named '<date> - Reimbursement Analysis - <district> -
    <N> Lots - Tierra Financial Advisors[ - <tag>].<ext>'."""
    from ut_pid_model import deliverable_basename
    cfg = result["cfg"]
    base = deliverable_basename(cfg, lots=result["dev"].total_lots)
    return out / f"{base}{suffix}"


def _forecast(out, result):
    """The exhibits carry their own label in place of 'Reimbursement Analysis'."""
    from ut_pid_model import deliverable_basename
    base = deliverable_basename(result["cfg"], lots=result["dev"].total_lots,
                                label="Forecast Exhibits")
    return out / f"{base}.xlsx"


WORKBOOK = ".xlsx"
MEMO = " - Memo.html"


def test_workbook_has_the_colorado_model_tab_set(deliverables):
    import openpyxl
    out, _r = deliverables
    wb = openpyxl.load_workbook(_deliverable(out, _r, WORKBOOK))
    assert wb.sheetnames == EXPECTED_TABS


def test_forecast_exhibits_cover_three_scenarios(deliverables):
    import openpyxl
    out, _r = deliverables
    wb = openpyxl.load_workbook(_forecast(out, _r))
    names = " ".join(wb.sheetnames)
    for prefix in ("A", "B", "C"):
        assert f"Exhibit {prefix}-1" in names or f"{prefix}-1" in names


def test_forecast_exhibits_are_labelled_forecast_exhibits(deliverables):
    """The exhibits read '<date> - Forecast Exhibits - <district> ...': the label
    sits where the workbook says 'Reimbursement Analysis', not on the tail, and
    that wording — like Colorado — appears nowhere inside."""
    import openpyxl
    import re
    out, _r = deliverables
    path = _forecast(out, _r)
    assert re.match(r"^\d{1,2}\.\d{1,2}\.\d{4} - Forecast Exhibits - ", path.name), path.name
    assert "Reimbursement Analysis" not in path.name
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                text = str(cell.value or "")
                for stale in ("Reimbursement Analysis", "COLORADO", "Colorado"):
                    assert stale not in text, f"{ws.title}!{cell.coordinate}: {text}"


def test_capi_fund_tab_runs_through_its_last_draw(deliverables):
    """The month-by-month fund roll ends on the month of the last draw — which
    is the straddling coupon, after the end of the period — and no later, so the
    tab neither shows a deposit it never spends nor trails empty months."""
    import openpyxl
    out, _r = deliverables
    senior = _r["senior"]
    wb = openpyxl.load_workbook(_deliverable(out, _r, WORKBOOK))
    ws = wb["CAPI Fund - First"]
    dates = [c[0].value.date() for c in ws.iter_rows(min_row=7, max_col=1)
             if hasattr(c[0].value, "year")]
    last_draw = max(p.payment_date for p in senior.schedule
                    if p.capitalized_interest > 0.005)
    assert dates
    assert (max(dates).year, max(dates).month) == (last_draw.year, last_draw.month)


def test_memo_states_the_utah_framework(deliverables):
    out, _r = deliverables
    html = _deliverable(out, _r, MEMO).read_text()
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
    out, _r = deliverables
    result = _r
    html = _deliverable(out, _r, MEMO).read_text()
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
        final_year=cfg.senior_final_year, capi_end=cfg.capi_end_date,
        call_provisions=calls)
    sf = SurplusFund(cfg, sm).build(senior, None, cfg.first_collection_year,
                                    senior.final_year)
    sub_par = SubordinateLien(cfg, sm).size_par(
        senior, cfg.first_collection_year, senior.final_year, surplus_fund=sf)
    return senior, sub_par, sm


def test_om_expense_is_the_single_district_cost_line():
    """District administration and O&M are one budget for a Utah PID, which
    rarely carries an operations levy to fund either.  One line, one growth
    rate — the $53,060 base is the reference deal's figure."""
    cfg = ModelConfig()
    assert cfg.om_expense == pytest.approx(53_060)
    assert cfg.om_growth_rate == pytest.approx(0.02)
    for gone in ("admin_cost", "admin_growth_rate", "admin_cost_av_limit"):
        assert not hasattr(cfg, gone), gone
    # district_costs now carries the trustee fees only.
    assert len(cfg.district_costs(cfg.delivery.year + 5)) == 2


def test_om_expense_starts_with_the_other_district_costs_and_inflates():
    cfg = ModelConfig(om_expense=40_000, om_growth_rate=0.035)
    start = cfg.district_cost_start_year or (cfg.delivery.year + 2)
    assert cfg.om_expense_for(start - 1) == 0.0
    assert cfg.om_expense_for(start) == pytest.approx(40_000)
    assert cfg.om_expense_for(start + 10) == pytest.approx(40_000 * 1.035 ** 10)


def test_om_expense_stops_above_the_taxable_value_limit():
    cfg = ModelConfig(om_expense=40_000, om_expense_av_limit=100_000_000)
    year = cfg.delivery.year + 5
    assert cfg.om_expense_for(year, 50_000_000) > 0
    assert cfg.om_expense_for(year, 150_000_000) == 0.0


def test_om_expense_is_netted_from_both_liens():
    """The sub lien's own revenue is `net_sub - net_senior`, so a cost netted
    from the senior side alone would be handed to the sub — an operating expense
    would then *raise* subordinate capacity.  It has to come off the top."""
    base_cfg = ModelConfig(om_expense=0.0)
    base = SummaryModel(base_cfg, viridian_farm_projections().build(base_cfg)).build()
    cfg = ModelConfig()                      # the calibrated $53,060
    with_om = SummaryModel(cfg, viridian_farm_projections().build(cfg)).build()
    year = 2035
    charge = with_om.row(year).om_expense
    assert charge > 0
    assert (base.net_senior_revenue(year) - with_om.net_senior_revenue(year)
            == pytest.approx(charge))
    assert (base.net_sub_revenue(year) - with_om.net_sub_revenue(year)
            == pytest.approx(charge))


def test_om_expense_reduces_both_senior_and_subordinate_capacity():
    senior_0, sub_0, _ = _sized(ModelConfig(om_expense=0.0))
    senior_om, sub_om, _ = _sized(ModelConfig())      # $53,060
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
    out, _r = deliverables
    ws = openpyxl.load_workbook(_deliverable(out, _r, WORKBOOK))["O&M Revenue"]
    hdrs = [ws.cell(row=5, column=c).value for c in range(1, 9)]
    assert "− O&M\nExpense" in hdrs
    assert "O&M Surplus /\n(Deficit)" in hdrs


def test_summary_detail_shows_the_om_column(deliverables):
    import openpyxl
    out, _r = deliverables
    ws = openpyxl.load_workbook(_deliverable(out, _r, WORKBOOK))["Summary - Detail"]
    hdrs = [c.value for c in ws[5]]
    assert "− District\nO&M" in hdrs


def test_workbook_carries_no_colorado_labels(deliverables):
    import openpyxl
    out, _r = deliverables
    wb = openpyxl.load_workbook(_deliverable(out, _r, WORKBOOK))
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


def test_every_excel_date_uses_the_m_d_yyyy_format(deliverables, tmp_path):
    """Dates read m/d/yyyy everywhere in the Excel output — every date-valued
    cell in the model workbook, the forecast exhibits and the inputs template."""
    import datetime
    import openpyxl
    from ut_pid_model import write_inputs_workbook
    out, _r = deliverables
    paths = [_deliverable(out, _r, WORKBOOK), _forecast(out, _r),
             tmp_path / "inputs.xlsx"]
    write_inputs_workbook(output_path=str(paths[-1]))
    seen = 0
    for path in paths:
        wb = openpyxl.load_workbook(path)
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, (datetime.datetime, datetime.date)):
                        seen += 1
                        assert cell.number_format == "m/d/yyyy", (
                            f"{path.name}:{ws.title}!{cell.coordinate} "
                            f"= {cell.number_format}")
    assert seen > 100, seen


def test_title_band_reads_m_d_yyyy_and_file_names_stay_filesystem_safe(deliverables):
    """The title band reads m/d/yyyy; file names carry the same date as m.d.yyyy,
    because a slash is a path separator and would be stripped out entirely."""
    import openpyxl
    import re
    out, _r = deliverables
    path = _deliverable(out, _r, WORKBOOK)
    assert re.match(r"^\d{1,2}\.\d{1,2}\.\d{4} - ", path.name), path.name
    assert "/" not in path.name
    wb = openpyxl.load_workbook(path)
    banded = 0
    for ws in wb.worksheets:
        # Development Projections and Notes carry their own district-first
        # header rather than the standard title band (as in Colorado).
        if not str(ws.cell(row=2, column=1).value or "").startswith(
                "Reimbursement Analysis"):
            continue
        banded += 1
        assert re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}",
                            str(ws.cell(row=1, column=1).value)), ws.title
    assert banded >= 14, banded


# ── The revenue wrap sizes at the entered coupons ────────────────────────────

def _sized_at(coupon, inputs_rate=None, built=None):
    """Size the senior lien with a single Term row at `coupon`, off `inputs_rate`."""
    cfg = ModelConfig()
    if inputs_rate is not None:
        cfg.senior_interest_rate = inputs_rate
    dev = viridian_farm_projections().build(cfg)
    sm = SummaryModel(cfg, dev).build()
    kw = dict(name="Senior", rate=cfg.senior_interest_rate, coverage=cfg.dsc_senior,
              delivery=cfg.delivery,
              first_principal_year=cfg.senior_first_principal_year,
              final_year=cfg.senior_final_year, capi_end=cfg.capi_end_date,
              call_provisions=CallProvisions(cfg.premium_call_date, cfg.par_call_date,
                                             cfg.premium_call_price))
    if coupon is not None:
        kw["coupon_scale"] = {cfg.senior_final_year: coupon}
        kw["term_bonds"] = [(cfg.senior_first_principal_year, cfg.senior_final_year, coupon)]
    senior = size_senior_with_dynamic_dsrf(SeniorLienSizer(cfg, sm), **kw)
    return cfg, sm, senior


def test_wrap_at_entered_coupon_is_exact_when_it_equals_the_flat_rate():
    """Entering a scale equal to the Inputs rate must not move the par — the
    weighted-average coupon reduces to the flat rate exactly."""
    _c1, _s1, blank = _sized_at(None)
    cfg, _s2, scaled = _sized_at(ModelConfig().senior_interest_rate)
    assert scaled.par_amount == blank.par_amount


def test_wrap_ignores_the_inputs_rate_once_coupons_are_entered():
    """The Debt Structure coupon drives sizing; the Inputs rate is only the
    fallback for a blank sheet."""
    pars = {r: _sized_at(0.0625, inputs_rate=r)[2].par_amount
            for r in (0.04, 0.05, 0.05875, 0.0625, 0.08)}
    assert len(set(pars.values())) == 1, pars


def test_wrap_at_entered_coupon_holds_the_coverage_target():
    """Sized at 5.875% but paying 6.250%, the deal used to land under target."""
    cfg, sm, senior = _sized_at(0.0625, inputs_rate=0.05875)
    cov = senior_coverage_dataframe(cfg, sm, senior)
    assert cov["coverage"].iloc[-1] >= cfg.dsc_senior


def test_a_higher_coupon_supports_less_par():
    pars = [_sized_at(c, inputs_rate=0.05875)[2].par_amount
            for c in (0.05, 0.05875, 0.0625, 0.07)]
    assert pars == sorted(pars, reverse=True), pars


# ── Bond pricing ─────────────────────────────────────────────────────────────

def _priced_tranche(coupon_scale, term_bonds, yield_scale=None):
    from ut_pid_model.debt_service import BondTranche
    return BondTranche(
        name="Senior", rate=0.05875, coverage=1.30, delivery=date(2026, 9, 1),
        first_principal_year=2027, final_year=2056, prin_month=3, prin_day=1,
        capi_end=None, dsrf_deposit=0.0, dsrf_earn_rate=0.0,
        par_amount=10_000_000.0,
        call_provisions=CallProvisions(date(2031, 3, 1), date(2034, 3, 1), 103.0),
        coupon_scale=coupon_scale, yield_scale=yield_scale, term_bonds=term_bonds)


DBC_2056A = {2032: 35_000, 2033: 80_000, 2034: 90_000, 2035: 100_000, 2036: 110_000,
             2037: 125_000, 2038: 135_000, 2039: 150_000, 2040: 165_000, 2041: 180_000,
             2042: 195_000, 2043: 215_000, 2044: 230_000, 2045: 250_000, 2046: 270_000,
             2047: 295_000, 2048: 320_000, 2049: 345_000, 2050: 370_000, 2051: 400_000,
             2052: 430_000, 2053: 460_000, 2054: 500_000, 2055: 535_000, 2056: 1_195_000}


def _dbc_tranche(price_scale=None):
    """Viridian Farm PID No. 2, 2056A term bond, as DBC sized and priced it."""
    from ut_pid_model.debt_service import BondTranche, PaymentRow, SeniorLienSizer
    par = sum(DBC_2056A.values())
    t = BondTranche(name="2056A", rate=0.0625, coverage=1.30, delivery=date(2026, 9, 30),
                    first_principal_year=2032, final_year=2056, prin_month=3, prin_day=1,
                    capi_end=None, dsrf_deposit=0.0, dsrf_earn_rate=0.0,
                    par_amount=par,
                    call_provisions=CallProvisions(date(2031, 3, 1), None, 103.0),
                    coupon_scale={2056: 0.0625}, term_bonds=[(2032, 2056, 0.06625)],
                    price_scale=price_scale)
    t.schedule = [PaymentRow(payment_date=date(y, m, 1),
                             principal=(DBC_2056A.get(y, 0.0) if m == 3 else 0.0),
                             interest=0.0, capitalized_interest=0.0,
                             dsrf_earnings=0.0, surplus_release=0.0)
                  for y in range(2027, 2057) for m in (3, 9)]
    SeniorLienSizer._apply_coupon_scale(t, None)
    return t


def test_interest_frequency_drives_the_schedule_and_the_price():
    """Annual means one coupon a year, on the principal date, and a price
    discounted at annual compounding; Semiannual is the default and unchanged."""
    from ut_pid_model.pricing import price_from_dated_date
    assert ModelConfig().coupon_frequency == 2
    assert ModelConfig(interest_frequency="Annual").coupon_frequency == 1
    assert ModelConfig(interest_frequency="annual").coupon_frequency == 1

    def sized(freq):
        cfg = ModelConfig(interest_frequency=freq)
        dev = viridian_farm_projections().build(cfg)
        sm = SummaryModel(cfg, dev).build()
        return cfg, size_senior_with_dynamic_dsrf(
            SeniorLienSizer(cfg, sm), name="Senior", rate=cfg.senior_interest_rate,
            coverage=cfg.dsc_senior, delivery=cfg.delivery,
            first_principal_year=cfg.senior_first_principal_year,
            final_year=cfg.senior_final_year, capi_end=cfg.capi_end_date,
            call_provisions=CallProvisions(cfg.premium_call_date, cfg.par_call_date,
                                           cfg.premium_call_price))

    _c2, semi = sized("Semiannual")
    _c1, ann = sized("Annual")
    assert len([p for p in semi.schedule if p.payment_date.year == 2040]) == 2
    assert len([p for p in ann.schedule if p.payment_date.year == 2040]) == 1
    assert all(p.payment_date.month == _c1.prin_maturity for p in ann.schedule)
    # Annual compounding discounts less heavily, so the same discount bond
    # prices HIGHER — which is why it cannot explain a price below ours.
    kw = dict(dated=date(2026, 9, 30), redemption_date=date(2056, 3, 1),
              coupon=0.0625, ytm=0.06625)
    assert price_from_dated_date(freq=1, **kw) > price_from_dated_date(freq=2, **kw)


def test_interest_is_paid_once_a_year_when_annual_twice_when_semiannual():
    """Not just the number of rows — the amounts.  An annual coupon carries the
    whole year; a semiannual one carries half, and the coupon that follows the
    principal date carries it on the reduced balance."""
    from ut_pid_model.pricing import days_30_360

    def sized(freq):
        cfg = ModelConfig(interest_frequency=freq)
        dev = viridian_farm_projections().build(cfg)
        sm = SummaryModel(cfg, dev).build()
        return cfg, size_senior_with_dynamic_dsrf(
            SeniorLienSizer(cfg, sm), name="Senior", rate=cfg.senior_interest_rate,
            coverage=cfg.dsc_senior, delivery=cfg.delivery,
            first_principal_year=cfg.senior_first_principal_year,
            final_year=cfg.senior_final_year, capi_end=cfg.capi_end_date,
            call_provisions=CallProvisions(cfg.premium_call_date, cfg.par_call_date,
                                           cfg.premium_call_price))

    for freq, per_year in (("Annual", 1), ("Semiannual", 2)):
        cfg, s = sized(freq)
        rows = s.schedule
        stub = days_30_360(cfg.delivery, rows[0].payment_date) / 360.0
        # First coupon: the stub, on the full par.
        assert rows[0].interest == pytest.approx(s.par_amount * s.rate * stub, rel=1e-9)
        # Second: a full period at this frequency, still on the full par.
        assert rows[1].interest == pytest.approx(
            s.par_amount * s.rate / per_year, rel=1e-9)
        assert len([p for p in rows if p.payment_date.year == 2040]) == per_year

    # The annual coupon is worth two semiannual ones.
    _ca, ann = sized("Annual")
    assert ann.schedule[1].interest == pytest.approx(
        ann.par_amount * ann.rate, rel=1e-9)


def test_tic_compounds_at_the_bond_frequency():
    """An annual-pay bond's TIC must compound annually, or the Sources & Uses
    statistics quote a rate the bonds do not pay."""
    from ut_pid_model.report import _tic
    d = date(2026, 9, 30)
    flows = [(date(2027 + k, 3, 1), 6.25) for k in range(30)]
    flows[-1] = (flows[-1][0], flows[-1][1] + 100.0)
    annual = _tic(flows, 100.0, d, 1)
    semi = _tic(flows, 100.0, d, 2)
    assert annual != pytest.approx(semi, abs=1e-6)
    # Same cash flows, same PV: the semiannual-compounded rate is the lower
    # nominal one, since it compounds twice as often.
    assert semi < annual


def test_frequency_round_trips_through_the_inputs_page(tmp_path):
    import openpyxl
    from ut_pid_model import load_inputs_workbook, write_inputs_workbook
    path = write_inputs_workbook(output_path=str(tmp_path / "freq.xlsx"))
    wb = openpyxl.load_workbook(path)
    ws = wb["Inputs"]
    cell = next(ws.cell(row=r, column=3) for r in range(5, ws.max_row + 1)
                if ws.cell(row=r, column=4).value == "INTEREST_FREQUENCY")
    assert cell.value == "Semiannual"
    assert any('"Semiannual,Annual"' in (dv.formula1 or "")
               for dv in ws.data_validations.dataValidation)
    cell.value = "Annual"
    wb.save(path)
    cfg, _dev = load_inputs_workbook(path)
    assert cfg.interest_frequency == "Annual"
    assert cfg.coupon_frequency == 1


def test_price_matches_excel_price_on_the_underwriters_run():
    """Viridian Farm PID No. 2, from the underwriter's own sheet:
    PRICE(9/30/2026, 3/1/2056, 6.250%, 6.625%, 100, 1) = 95.14825496,
    which truncates to the 95.148 DBC prints."""
    from ut_pid_model.pricing import clean_price, price_to_worst
    S, M, call = date(2026, 9, 30), date(2056, 3, 1), date(2031, 3, 1)
    assert clean_price(S, M, 0.0625, 0.06625, 100.0, 1) == pytest.approx(
        95.14825496, abs=1e-8)
    assert price_to_worst(S, M, 0.0625, 0.06625, [(call, 103.0)], 1) == pytest.approx(
        95.14825496, abs=1e-8)
    # Frequency matters: semiannual is a different bond and a different price.
    assert price_to_worst(S, M, 0.0625, 0.06625, [(call, 103.0)], 2) != pytest.approx(
        95.14825496, abs=1e-4)


def test_a_bond_reoffered_at_its_coupon_is_quoted_at_par():
    """The 9.000%/9.000% series on the same run prints 100.000.  Excel's PRICE
    returns 99.90 for it off a coupon date — quoting par is the convention."""
    from ut_pid_model.pricing import price_to_worst
    S = date(2026, 9, 30)
    for freq in (1, 2):
        assert price_to_worst(S, date(2056, 3, 15), 0.09, 0.09,
                              [(date(2031, 3, 1), 103.0)], freq) == pytest.approx(100.0)


def test_annual_frequency_reproduces_the_underwriters_oid():
    """End to end on the underwriter's sizing: price, interest and OID."""
    from ut_pid_model.debt_service import BondTranche, PaymentRow, SeniorLienSizer
    par = sum(DBC_2056A.values())
    t = BondTranche(name="2056A", rate=0.0625, coverage=1.30, delivery=date(2026, 9, 30),
                    first_principal_year=2032, final_year=2056, prin_month=3, prin_day=1,
                    capi_end=None, dsrf_deposit=0.0, dsrf_earn_rate=0.0,
                    par_amount=par,
                    call_provisions=CallProvisions(date(2031, 3, 1), None, 103.0),
                    coupon_scale={2056: 0.0625}, term_bonds=[(2032, 2056, 0.06625)],
                    coupon_frequency=1)
    t.schedule = [PaymentRow(payment_date=date(y, 3, 1), principal=DBC_2056A.get(y, 0.0),
                             interest=0.0, capitalized_interest=0.0,
                             dsrf_earnings=0.0, surplus_release=0.0)
                  for y in range(2027, 2057)]
    SeniorLienSizer._apply_coupon_scale(t, None)
    assert t.price_for(2056) == pytest.approx(95.148)
    assert t.schedule[0].interest == pytest.approx(188_226, abs=1)     # 151/360 stub
    assert t.schedule[1].interest == pytest.approx(448_750, abs=1)
    assert sum(p.principal + p.interest for p in t.schedule) == pytest.approx(
        17_139_788, abs=1)
    assert t.total_premium == pytest.approx(-348_373.60, abs=0.01)


def test_price_is_truncated_to_three_decimals():
    """DBC prints the price to three decimals and computes the OID from the
    TRUNCATED price — its -348,373.60 is exactly 7,180,000 x (95.148-100)/100.
    Truncate, not round, so the printed price and the OID always agree."""
    from ut_pid_model.debt_service import _truncate3
    assert _truncate3(95.1715) == 95.171
    assert _truncate3(95.1489999) == 95.148        # truncated, not rounded up
    assert _truncate3(95.148) == 95.148            # an entered price is untouched
    # Floating-point noise must not invent a discount on a par bond.
    assert _truncate3(99.99999999999) == 100.0
    assert _truncate3(100.00000000001) == 100.0


def test_price_and_oid_always_agree():
    """Whatever the price, the OID is that price applied to the principal —
    no full-precision residue between what is printed and what is booked."""
    t = _dbc_tranche()
    price = t.price_for(2056)
    assert price == pytest.approx(round(price, 3), abs=1e-12)
    for year in (2032, 2040, 2056):
        assert t.premium_for(year, DBC_2056A[year]) == pytest.approx(
            (price / 100.0 - 1.0) * DBC_2056A[year], abs=1e-9)


def test_entered_price_overrides_the_calculated_one():
    calc = _dbc_tranche().price_for(2056)
    entered = _dbc_tranche({2056: 95.148}).price_for(2056)
    assert entered == pytest.approx(95.148)
    assert calc != pytest.approx(95.148, abs=1e-4)
    # A term bond carries the entered price on every one of its installments.
    assert _dbc_tranche({2056: 95.148}).price_for(2040) == pytest.approx(95.148)


def test_entered_price_reproduces_the_dbc_oid_exactly():
    """With DBC's price entered, every maturity's OID matches DBC to the cent."""
    t = _dbc_tranche({2056: 95.148})
    expected = {2032: -1_698.20, 2033: -3_881.60, 2040: -8_005.80, 2056: -57_981.40}
    for year, oid in expected.items():
        assert t.premium_for(year, DBC_2056A[year]) == pytest.approx(oid, abs=0.01)
    assert t.total_premium == pytest.approx(-348_373.60, abs=0.01)


def test_price_column_is_appended_so_the_other_columns_do_not_move(tmp_path):
    _p, _wb, ws, row_of = _debt_structure_sheet(tmp_path, "price_col.xlsx")
    hdr = row_of(ModelConfig().senior_first_principal_year) - 1
    assert [ws.cell(row=hdr, column=c).value for c in range(2, 8)] == [
        "Maturity Year", "Par Amount", "Coupon", "Yield",
        "Type (Serial/Term)", "Price (optional)"]


def test_entered_price_round_trips_through_the_sheet(tmp_path):
    from ut_pid_model import load_inputs_workbook
    cfg0 = ModelConfig()
    path, wb, ws, row_of = _debt_structure_sheet(tmp_path, "price_rt.xlsx")
    r = row_of(cfg0.senior_final_year)
    ws.cell(row=r, column=4).value = 0.0625
    ws.cell(row=r, column=5).value = 0.06625
    ws.cell(row=r, column=6).value = "Term"
    ws.cell(row=r, column=7).value = 95.148
    wb.save(path)
    cfg, _dev = load_inputs_workbook(path)
    assert cfg.senior_price_scale == {cfg0.senior_final_year: 95.148}


def test_period_count_is_day_accurate():
    """The dated date's DAY was discarded — 5 July and 31 July priced alike."""
    from ut_pid_model.pricing import semiannual_periods
    mat = date(2056, 3, 1)
    n_early = semiannual_periods(date(2026, 7, 5), mat)
    n_late = semiannual_periods(date(2026, 7, 31), mat)
    assert n_early != n_late
    assert n_early == pytest.approx((30 * 360 - 4 * 30 - 4) / 180.0)
    # A coupon-date settlement is still a whole number of periods.
    assert semiannual_periods(date(2026, 9, 1), mat) == pytest.approx(59.0)


def test_a_bond_reoffered_at_its_coupon_prices_at_par_on_any_dated_date():
    """A new issue settles on its dated date, so no accrued changes hands and
    the price is the plain present value — which is exactly 100 when the yield
    equals the coupon, whatever the fractional first period."""
    from ut_pid_model.pricing import price_to_worst
    for dated in (date(2024, 9, 26), date(2026, 7, 5), date(2026, 9, 1)):
        assert price_to_worst(dated, date(2054, 3, 1), 0.0625, 0.0625,
                              [(date(2029, 3, 1), 103.0)]) == pytest.approx(100.0)


def test_clean_price_matches_excel_price():
    """clean_price is the 30/360 quoted convention — Excel PRICE() with accrued
    subtracted.  Pinned against values computed independently."""
    from ut_pid_model.pricing import clean_price
    mat = date(2056, 3, 1)
    # Settling on a coupon date: the level-coupon present value.
    assert clean_price(date(2026, 9, 1), mat, 0.0625, 0.06625) == pytest.approx(95.1672, abs=5e-4)
    assert clean_price(date(2026, 3, 1), mat, 0.0625, 0.06625) == pytest.approx(95.1407, abs=5e-4)
    # Mid-period, where the accrued subtraction bites.
    assert clean_price(date(2026, 7, 5), mat, 0.0625, 0.06625) == pytest.approx(95.1479, abs=5e-4)
    # Par on a coupon date, and the OID sign convention.
    assert clean_price(date(2026, 9, 1), mat, 0.05, 0.05) == pytest.approx(100.0)
    assert clean_price(date(2026, 9, 1), mat, 0.05, 0.06) < 100.0


def test_price_uses_the_entered_coupon_not_the_flat_sizing_rate():
    """3/1/2056 term, 6.250% coupon / 6.625% yield, callable 3/1/2031 at 103.

    Pricing passed the Inputs-page flat rate (5.875%) as the coupon, which
    priced this at 90.3345 instead of 95.1672 — a 4.8-point error on the OID,
    and the OID is a Source of Funds."""
    t = _priced_tranche({2056: 0.0625}, [(2027, 2056, 0.06625)])
    assert t.price_for(2056) == pytest.approx(95.1672, abs=5e-4)
    # Every installment of the term prices at the term's own coupon and yield.
    assert t.price_for(2040) == pytest.approx(t.price_for(2056), abs=1e-9)
    # A discount bond prices to maturity, not to the premium call.
    from ut_pid_model.pricing import bond_price
    assert bond_price(9.0, 0.0625, 0.06625, 103.0) > t.price_for(2056)


def test_a_par_bond_still_prices_at_par():
    t = _priced_tranche({2056: 0.0625}, [(2027, 2056, 0.0625)])
    assert t.price_for(2056) == pytest.approx(100.0, abs=1e-9)
    assert t.premium_for(2056, 10_000_000) == pytest.approx(0.0, abs=1e-6)


def test_each_term_bond_prices_at_its_own_coupon():
    """The coupon sits on the term's FINAL maturity row, and _schedule_lookup
    carries forward — so a later term must not inherit the earlier term's."""
    t = _priced_tranche({2040: 0.0500, 2056: 0.0625},
                        [(2027, 2040, 0.0525), (2041, 2056, 0.06625)])
    assert t.coupon_for(2035) == pytest.approx(0.0500)
    assert t.coupon_for(2045) == pytest.approx(0.0625)
    assert t.price_for(2045) == pytest.approx(t.price_for(2056), abs=1e-9)
    assert t.price_for(2035) != pytest.approx(t.price_for(2045), abs=1e-6)


def test_blank_debt_structure_prices_at_the_flat_rate():
    """No coupon scale ⇒ the flat rate is the coupon, and a bond reoffered at
    its coupon prices at par.  The base case must not move."""
    t = _priced_tranche(None, None)
    assert t.coupon_for(2040) == pytest.approx(0.05875)
    assert t.price_for(2056) == pytest.approx(100.0, abs=1e-9)


# ── Debt Structure tab (per-maturity par / coupon / yield, serial|term) ──────

def _debt_structure_sheet(tmp_path, name="ds.xlsx"):
    """A fresh inputs workbook plus a helper that finds a senior maturity row."""
    import openpyxl
    from ut_pid_model import write_inputs_workbook
    path = write_inputs_workbook(output_path=str(tmp_path / name))
    wb = openpyxl.load_workbook(path)
    ws = wb["Debt Structure"]
    start = next(r for r in range(1, ws.max_row + 1)
                 if str(ws.cell(row=r, column=2).value or "").upper().startswith("SENIOR BONDS"))

    def row_of(year):
        for r in range(start + 2, ws.max_row + 1):
            if ws.cell(row=r, column=2).value == year:
                return r
        raise AssertionError(f"no senior row for {year}")
    return path, wb, ws, row_of


def test_debt_structure_tab_has_the_par_coupon_yield_type_columns(tmp_path):
    _p, _wb, ws, row_of = _debt_structure_sheet(tmp_path)
    hdr_row = row_of(ModelConfig().senior_first_principal_year) - 1
    assert [ws.cell(row=hdr_row, column=c).value for c in range(2, 7)] == [
        "Maturity Year", "Par Amount", "Coupon", "Yield", "Type (Serial/Term)"]
    # The Type column is a Serial/Term dropdown.
    assert any('"Serial,Term"' in (dv.formula1 or "")
               for dv in ws.data_validations.dataValidation)


def test_debt_structure_number_formats_follow_the_columns(tmp_path):
    """Par Amount moved the later columns one to the right; the dollar and
    percent formats have to have moved with them."""
    _p, _wb, ws, row_of = _debt_structure_sheet(tmp_path, "fmt.xlsx")
    for year in (ModelConfig().senior_first_principal_year,
                 ModelConfig().senior_final_year):
        r = row_of(year)
        fmts = {ws.cell(row=r, column=c).column_letter: ws.cell(row=r, column=c).number_format
                for c in range(2, 7)}
        assert fmts == {"B": "General",      # maturity year
                        "C": "#,##0",        # par amount — dollars
                        "D": "0.000%",       # coupon
                        "E": "0.000%",       # yield
                        "F": "General"}, fmts


def test_debt_structure_scale_round_trips_through_the_sheet(tmp_path):
    """Write a par / coupon / yield scale out and read it back unchanged — the
    end-to-end check that values land in the columns the loader reads."""
    from ut_pid_model import (load_inputs_workbook, write_inputs_workbook,
                              viridian_farm_projections)
    cfg = ModelConfig()
    cfg.senior_par_schedule = {2052: 1_250_000.0, 2053: 1_250_000.0, 2054: 1_500_000.0}
    cfg.senior_coupon_scale = {2052: 0.055, 2053: 0.0575, 2054: 0.06}
    cfg.senior_yield_scale = {2052: 0.0525, 2053: 0.0550, 2054: 0.0636}
    path = write_inputs_workbook(cfg, viridian_farm_projections(),
                                 output_path=str(tmp_path / "rt.xlsx"))
    back, _dev = load_inputs_workbook(path)
    assert back.senior_par_schedule == cfg.senior_par_schedule
    assert back.senior_coupon_scale == cfg.senior_coupon_scale
    assert back.senior_yield_scale == cfg.senior_yield_scale
    assert back.senior_term_bonds is None       # every row is a Serial


def test_blank_debt_structure_is_preliminary_flat_rate_sizing(tmp_path):
    from ut_pid_model import load_inputs_workbook
    path, _wb, _ws, _row_of = _debt_structure_sheet(tmp_path, "blank.xlsx")
    cfg, _dev = load_inputs_workbook(path)
    assert cfg.senior_coupon_scale is None
    assert cfg.senior_yield_scale is None
    assert cfg.senior_term_bonds is None
    assert cfg.senior_par_schedule is None


def test_one_term_row_makes_the_whole_structure_a_term_bond(tmp_path):
    """A single Term row at the final maturity spans first principal year → that
    maturity, and a yield above the coupon prices at a discount."""
    from ut_pid_model import load_inputs_workbook
    cfg0 = ModelConfig()
    path, wb, ws, row_of = _debt_structure_sheet(tmp_path, "term.xlsx")
    r = row_of(cfg0.senior_final_year)
    ws.cell(row=r, column=4).value = 0.06        # coupon
    ws.cell(row=r, column=5).value = 0.0636      # yield
    ws.cell(row=r, column=6).value = "Term"
    wb.save(path)
    cfg, _dev = load_inputs_workbook(path)
    assert cfg.senior_term_bonds == [
        (cfg0.senior_first_principal_year, cfg0.senior_final_year, 0.0636)]
    assert cfg.senior_coupon_scale == {cfg0.senior_final_year: 0.06}
    assert cfg.senior_yield_scale is None        # the term row is not a serial


def test_par_amount_column_overrides_the_revenue_wrap(tmp_path):
    from ut_pid_model import load_inputs_workbook, SummaryModel, SeniorLienSizer
    from ut_pid_model import CallProvisions, size_senior_with_dynamic_dsrf
    path, wb, ws, row_of = _debt_structure_sheet(tmp_path, "par.xlsx")
    manual = {2050: 1_000_000.0, 2051: 1_000_000.0, 2052: 1_000_000.0,
              2053: 1_000_000.0, 2054: 1_000_000.0}
    for y, par in manual.items():
        ws.cell(row=row_of(y), column=3).value = par
    wb.save(path)
    cfg, dev = load_inputs_workbook(path)
    assert cfg.senior_par_schedule == manual
    sm = SummaryModel(cfg, dev.build(cfg)).build()
    senior = size_senior_with_dynamic_dsrf(
        SeniorLienSizer(cfg, sm),
        name="Senior", rate=cfg.senior_interest_rate, coverage=cfg.dsc_senior,
        delivery=cfg.delivery, first_principal_year=cfg.senior_first_principal_year,
        final_year=cfg.senior_final_year, capi_end=cfg.capi_end_date,
        call_provisions=CallProvisions(cfg.premium_call_date, cfg.par_call_date,
                                       cfg.premium_call_price),
        par_schedule=cfg.senior_par_schedule)
    assert senior.par_amount == pytest.approx(sum(manual.values()))
    got = {p.payment_date.year: p.principal for p in senior.schedule if p.principal}
    assert got == pytest.approx(manual)


def test_old_layout_debt_structure_sheet_is_refused(tmp_path):
    """The tab gained a Par Amount column, shifting every later column right.  A
    sheet saved against the old layout parses as numbers, so it has to be caught
    rather than silently sizing a $0.06 bond."""
    from ut_pid_model import load_inputs_workbook
    cfg0 = ModelConfig()
    path, wb, ws, row_of = _debt_structure_sheet(tmp_path, "old.xlsx")
    r = row_of(cfg0.senior_final_year)
    ws.cell(row=r, column=3).value = 0.06        # old Coupon column
    ws.cell(row=r, column=4).value = 0.0636      # old Yield column
    ws.cell(row=r, column=5).value = cfg0.senior_final_year   # old Term Final Maturity
    wb.save(path)
    with pytest.raises(ValueError, match="older layout"):
        load_inputs_workbook(path)


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
