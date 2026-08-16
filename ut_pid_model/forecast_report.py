"""
forecast_report.py — CPA-style "Forecast of Cash Balances and Cash Receipts and
Disbursements" exhibits.

Produces a standalone Excel workbook whose sheets mirror the exhibit set used in
the accountant's forecast for a Utah PID debt-service fund:

  Exhibit A    Summary of Debt Service Fund Activity (master cash flow)
  Exhibit A-1  Summary of Taxable Values and Net Tax Revenues
  Exhibit A-2  Schedule of Estimated Taxable Value — Builder Lot Inventory
  Exhibit A-3  Schedule of Estimated Taxable Value — Residential
  Exhibit A-4  Schedule of Estimated Market Value — Residential Development
  Exhibit A-5  Schedule of Estimated Senior Bonds Debt Service Requirements
  Exhibit A-6  Schedule of Estimated Subordinate Bonds Debt Service Requirements
  Exhibit A-7  Estimated Sources and Uses of Funds

The exhibits are driven by the same model objects as the rest of the package,
so any scenario the model runs can be presented in this regulatory format.
"""

from __future__ import annotations

import os

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from .config import ModelConfig

# ── Styles ──────────────────────────────────────────────────────────────────
_NAVY  = PatternFill("solid", fgColor="1F3864")
_LIGHT = PatternFill("solid", fgColor="DDEBF7")
_GRAY  = PatternFill("solid", fgColor="F2F2F2")
_WHITE = PatternFill("solid", fgColor="FFFFFF")
_TOTAL = PatternFill("solid", fgColor="C6E0B4")

_TOP = Side(style="thin", color="000000")
_TOPBORDER = Border(top=_TOP)
_DBL = Border(top=Side(style="double", color="000000"))

_EX_FONT   = Font(name="Times New Roman", bold=True, size=11, color="000000")
_TITLE     = Font(name="Times New Roman", bold=True, size=10)
_SUBTITLE  = Font(name="Times New Roman", italic=True, size=9)
_HDR       = Font(name="Times New Roman", bold=True, size=8)
_BODY      = Font(name="Times New Roman", size=8)
_BOLD      = Font(name="Times New Roman", bold=True, size=8)
_FOOT      = Font(name="Times New Roman", italic=True, size=7, color="595959")

_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
_R = Alignment(horizontal="right", vertical="center")
_L = Alignment(horizontal="left", vertical="center")

_DOLLAR = '#,##0;(#,##0)'
_RATE3  = '0.0000'
_PCT1   = '0.0%'
_PCT2   = '0.00%'


def _exhibit_header(ws, exhibit_no, cfg, schedule_title, last_col, scenario_label=""):
    titles = [
        f"EXHIBIT {exhibit_no}",
        cfg.pid_name.upper(),
        f"IN {cfg.county.upper()} COUNTY, COLORADO",
        "FORECAST OF CASH BALANCES AND CASH RECEIPTS AND DISBURSEMENTS",
        "FOR DEBT SERVICE FUND ONLY",
        schedule_title.upper(),
    ]
    if scenario_label:
        titles.append(scenario_label.upper())
    for i, t in enumerate(titles, 1):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=last_col)
        c = ws.cell(row=i, column=1, value=t)
        c.font = _EX_FONT if i == 1 else _TITLE
        c.alignment = _C
    return 8  # first data-header row


def _col_headers(ws, row, headers):
    """headers: list of (col, multi-line label, width)."""
    for col, label, width in headers:
        c = ws.cell(row=row, column=col, value=label)
        c.font, c.alignment = _HDR, _C
        c.fill = _LIGHT
        c.border = Border(bottom=_TOP)
        ws.column_dimensions[get_column_letter(col)].width = width


def _w(ws, r, c, v, fmt=None, font=None, align=_R, fill=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font or _BODY
    cell.alignment = align
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


def _footer(ws, row, last_col, notes):
    for note in notes:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
        c = ws.cell(row=row, column=1, value=note)
        c.font = _FOOT
        c.alignment = _L
        row += 1
    return row


# ── Exhibit A-1: Taxable Values & Net Tax Revenues ──────────────────────────
def _build_a1(ws, cfg, sm, prefix="A", scenario_label=""):
    mill = cfg.effective_ds_mill_levy
    # Build the column layout dynamically — only include the centrally assessed,
    # commercial, and state-assessed columns when those values are present.
    cols = [(1, "Collection\nYear", 11),
            (2, f"Residential\nTaxable Value\n(Exhibit {prefix}-3)", 16),
            (3, f"Builder Lot Inventory\nTaxable Value\n(Exhibit {prefix}-2)", 16)]
    n = 3
    og_col = comm_col = state_col = None
    if sm.has_centrally_assessed:
        n += 1; og_col = n
        cols.append((n, f"Centrally Assessed\nTaxable Value\n{cfg.centrally_assessed_ratio:.1%}", 14))
    if sm.has_commercial:
        n += 1; comm_col = n
        cols.append((n, f"Commercial\nTaxable Value\n{cfg.commercial_assessment_ratio:.2%}", 14))
    if sm.has_state_assessed:
        n += 1; state_col = n
        cols.append((n, "State\nTaxable Value", 13))
    tot, dml, col_, sot_, tre, net_ = n + 1, n + 2, n + 3, n + 4, n + 5, n + 6
    cols += [
        (tot, "Total\nTaxable Value", 15),
        (dml, f"Debt\nMill Levy\n{mill:.4f}", 11),
        (col_, f"Mill Levy\nCollections\n{cfg.tax_collect_mill_prc:.2%}", 13),
        (sot_, f"Specific\nOwnership Taxes\n{cfg.uniform_fee_prc:.0%}", 13),
        (tre, f"County\nTreasurer Fee\n{cfg.county_collection_fee:.1%}", 13),
        (net_, "Net Tax\nRevenue", 13),
    ]
    hr = _exhibit_header(ws, f"{prefix}-1", cfg,
                         "Summary of Taxable Values and Net Tax Revenues", net_,
                         scenario_label)
    _col_headers(ws, hr, cols)
    r = hr + 1
    for row in sm.rows:
        treas = -row.mill_revenue * cfg.county_collection_fee
        net = row.mill_revenue + row.uniform_fee_revenue + treas
        fill = _GRAY if (r - hr) % 2 == 0 else _WHITE
        _w(ws, r, 1, row.collection_year, font=_BODY, align=_C, fill=fill)
        _w(ws, r, 2, round(row.residential_av) or None, _DOLLAR, fill=fill)
        _w(ws, r, 3, round(row.lot_av) or None, _DOLLAR, fill=fill)
        if og_col:
            _w(ws, r, og_col, round(row.centrally_assessed_av) or None, _DOLLAR, fill=fill)
        if comm_col:
            _w(ws, r, comm_col, round(row.commercial_av) or None, _DOLLAR, fill=fill)
        if state_col:
            _w(ws, r, state_col, round(row.state_av) or None, _DOLLAR, fill=fill)
        _w(ws, r, tot, round(row.total_av) or None, _DOLLAR, fill=fill)
        _w(ws, r, dml, mill if row.total_av else None, _RATE3, fill=fill)
        _w(ws, r, col_, round(row.mill_revenue) or None, _DOLLAR, fill=fill)
        _w(ws, r, sot_, round(row.uniform_fee_revenue) or None, _DOLLAR, fill=fill)
        _w(ws, r, tre, round(treas) or None, _DOLLAR, fill=fill)
        _w(ws, r, net_, round(net) or None, _DOLLAR, fill=fill)
        r += 1
    ws.freeze_panes = "A9"


# ── Exhibit A-2: Builder Lot Inventory Taxable Value ──────────────────────────────
def _build_a2(ws, cfg, dev, sm, prefix="A", scenario_label=""):
    hr = _exhibit_header(ws, f"{prefix}-2", cfg,
                         "Schedule of Estimated Taxable Value - Builder Lot Inventory", 5, scenario_label)
    _col_headers(ws, hr, [
        (1, "Collection\nYear", 11),
        (2, "Cumulative\nBuilder Lot Inventory\nMarket Value", 16),
        (3, "Builder Lot Inventory\nTaxable Ratio", 14),
        (4, "Taxable Value\nof Builder Lot Inventory", 15),
        (5, "(2-yr lag\napplied)", 12),
    ])
    r = hr + 1
    for row in sm.rows:
        # market value backing the collection-year lot-inventory AV (lagged / held);
        # lot-inventory inventory only — lots built out into homes are excluded
        mkt = dev.vacant_lot_market_value(cfg.av_source_year(row.collection_year))
        rate = cfg.lot_inventory_taxable_rate(row.collection_year)
        fill = _GRAY if (r - hr) % 2 == 0 else _WHITE
        _w(ws, r, 1, row.collection_year, font=_BODY, align=_C, fill=fill)
        _w(ws, r, 2, round(mkt) or None, _DOLLAR, fill=fill)
        _w(ws, r, 3, rate if row.lot_av else None, _PCT2, fill=fill)
        _w(ws, r, 4, round(row.lot_av) or None, _DOLLAR, fill=fill)
        _w(ws, r, 5, "lagged" if row.lot_av else None, font=_FOOT, align=_C, fill=fill)
        r += 1
    ws.freeze_panes = "A9"


# ── Exhibit A-3: Residential Taxable Value ──────────────────────────────
def _build_a3(ws, cfg, dev, sm, prefix="A", scenario_label=""):
    hr = _exhibit_header(ws, f"{prefix}-3", cfg,
                         "Schedule of Estimated Taxable Value - Residential", 5, scenario_label)
    _col_headers(ws, hr, [
        (1, "Collection\nYear", 11),
        (2, f"Cumulative Market\nValue of Homes\n(Exhibit {prefix}-4)", 18),
        (3, f"Annual\nReassessment\n{cfg.reassess_rate:.0%}", 13),
        (4, "Taxable\nRatio\n(Varies)" if cfg.residential_assessment_schedule
            else f"Taxable\nRatio (Utah)\n{cfg.resid_taxable_ratio:.2%}", 13),
        (5, "Residential\nTaxable Value", 16),
    ])
    r = hr + 1
    M = dev.cumulative_home_market_value
    for row in sm.rows:
        mkt = M.get(cfg.av_source_year(row.collection_year), 0.0)
        rate = cfg.residential_assessment_rate(row.collection_year)
        fill = _GRAY if (r - hr) % 2 == 0 else _WHITE
        _w(ws, r, 1, row.collection_year, font=_BODY, align=_C, fill=fill)
        _w(ws, r, 2, round(mkt) or None, _DOLLAR, fill=fill)
        _w(ws, r, 3, cfg.reassess_rate if mkt else None, _PCT1, fill=fill)
        _w(ws, r, 4, rate if row.residential_av else None, _PCT2, fill=fill)
        _w(ws, r, 5, round(row.residential_av) or None, _DOLLAR, fill=fill)
        r += 1
    ws.freeze_panes = "A9"


# ── Exhibit A-4: Residential Market Value / Development ───────────────────────
def _build_a4(ws, cfg, dev, prefix="A", scenario_label=""):
    hr = _exhibit_header(ws, f"{prefix}-4", cfg,
                         "Schedule of Estimated Market Value - Residential Development", 6, scenario_label)
    _col_headers(ws, hr, [
        (1, "Construction/\nRoll Year", 12),
        (2, "Homes\nClosed", 10),
        (3, "Avg. Selling\nPrice (ASP)", 13),
        (4, "Annual Value\nof New Homes", 15),
        (5, "Cumulative Market\nValue of Homes", 17),
        (6, "Lot Market\nValue Placed", 14),
    ])
    r = hr + 1
    M = dev.cumulative_home_market_value
    years = sorted(set(list(dev.home_closings) + list(dev.lot_market_value)))
    for y in years:
        fill = _GRAY if (r - hr) % 2 == 0 else _WHITE
        closed = dev.home_closings.get(y, 0)
        _w(ws, r, 1, y, font=_BODY, align=_C, fill=fill)
        _w(ws, r, 2, closed or None, '#,##0', fill=fill)
        _w(ws, r, 3, round(dev.asp(y)) if closed else None, _DOLLAR, fill=fill)
        _w(ws, r, 4, round(dev.new_home_market_value(y)) or None, _DOLLAR, fill=fill)
        _w(ws, r, 5, round(M.get(y, 0.0)) or None, _DOLLAR, fill=fill)
        _w(ws, r, 6, round(dev.lot_market_value.get(y, 0.0)) or None, _DOLLAR, fill=fill)
        r += 1
    ws.freeze_panes = "A9"


# ── Exhibit A-5: Senior Bonds Debt Service ───────────────────────────────────
def _build_a5(ws, cfg, senior, prefix="A", scenario_label=""):
    hr = _exhibit_header(ws, f"{prefix}-5", cfg,
                         "Schedule of Estimated Senior Bonds Debt Service Requirements", 8, scenario_label)
    _col_headers(ws, hr, [
        (1, "Payment\nDate", 12), (2, "Rate", 9), (3, "Principal", 13),
        (4, "Interest", 13), (5, "Capitalized\nInterest", 13),
        (6, "DSRF\nEarnings", 12), (7, "Net\nPayment", 13),
        (8, "Principal\nOutstanding", 14),
    ])
    r = hr + 1
    bal = senior.par_amount
    for p in senior.schedule:
        fill = _GRAY if (r - hr) % 2 == 0 else _WHITE
        if p.principal:
            bal -= p.principal
        _w(ws, r, 1, p.payment_date, "DD-MMM-YY", font=_BODY, align=_C, fill=fill)
        _w(ws, r, 2, senior.rate if p.payment_date.month == senior.prin_month else None, _PCT2, fill=fill)
        _w(ws, r, 3, round(p.principal) or None, _DOLLAR, fill=fill)
        _w(ws, r, 4, round(p.interest) or None, _DOLLAR, fill=fill)
        _w(ws, r, 5, -round(p.capitalized_interest) or None, _DOLLAR, fill=fill)
        _w(ws, r, 6, -round(p.dsrf_earnings) or None, _DOLLAR, fill=fill)
        _w(ws, r, 7, round(p.net_total) or None, _DOLLAR, fill=fill)
        _w(ws, r, 8, round(bal) or None, _DOLLAR, fill=fill)
        r += 1
    _w(ws, r, 1, "Total", font=_BOLD, align=_L, fill=_TOTAL)
    _w(ws, r, 3, round(senior.par_amount), _DOLLAR, font=_BOLD, fill=_TOTAL)
    _w(ws, r, 7, round(senior.total_net_ds), _DOLLAR, font=_BOLD, fill=_TOTAL)
    for c in (2, 4, 5, 6, 8):
        ws.cell(row=r, column=c).fill = _TOTAL
    ws.freeze_panes = "A9"


# ── Exhibit A-6: Subordinate (cash-flow) Bonds Debt Service ──────────────────
def _build_a6(ws, cfg, sub, prefix="A", scenario_label=""):
    hr = _exhibit_header(ws, f"{prefix}-6", cfg,
                         "Schedule of Estimated Subordinate Bonds Debt Service Requirements", 8, scenario_label)
    _col_headers(ws, hr, [
        (1, "Payment\nYear", 11), (2, "Rate", 9), (3, "Principal", 13),
        (4, "Accrued\nInterest", 13), (5, "Interest\nPaid", 13),
        (6, "Unpaid (Accreted)\nInterest", 15), (7, "Total\nPayment", 13),
        (8, "Cumulative\nOutstanding Balance", 16),
    ])
    r = hr + 1
    for row in sub.rows:
        fill = _GRAY if (r - hr) % 2 == 0 else _WHITE
        outstanding = row["principal_balance"] + row["accrued_balance"]
        _w(ws, r, 1, row["year"], font=_BODY, align=_C, fill=fill)
        _w(ws, r, 2, sub.rate if row["principal_paid"] else None, _PCT2, fill=fill)
        _w(ws, r, 3, round(row["principal_paid"]) or None, _DOLLAR, fill=fill)
        _w(ws, r, 4, round(row["current_interest"]) or None, _DOLLAR, fill=fill)
        _w(ws, r, 5, round(row["interest_paid"]) or None, _DOLLAR, fill=fill)
        _w(ws, r, 6, round(row["accrued_balance"]) or None, _DOLLAR, fill=fill)
        _w(ws, r, 7, round(row["total_paid"]) or None, _DOLLAR, fill=fill)
        _w(ws, r, 8, round(outstanding) or None, _DOLLAR, fill=fill)
        r += 1
    _w(ws, r, 1, "Total", font=_BOLD, align=_L, fill=_TOTAL)
    _w(ws, r, 3, round(sub.total_principal_paid), _DOLLAR, font=_BOLD, fill=_TOTAL)
    _w(ws, r, 5, round(sub.total_interest_paid), _DOLLAR, font=_BOLD, fill=_TOTAL)
    _w(ws, r, 7, round(sub.total_payments), _DOLLAR, font=_BOLD, fill=_TOTAL)
    for c in (2, 4, 6, 8):
        ws.cell(row=r, column=c).fill = _TOTAL
    ws.freeze_panes = "A9"


# ── Exhibit A-7: Sources & Uses ──────────────────────────────────────────────
def _build_a7(ws, cfg, senior, su, sub_par, prefix="A", scenario_label=""):
    hr = _exhibit_header(ws, f"{prefix}-7", cfg, "Estimated Sources and Uses of Funds", 4, scenario_label)
    capi = sum(p.capitalized_interest for p in senior.schedule)
    uwd_sr = cfg.uwd_senior * senior.par_amount
    uwd_sub = cfg.uwd_sub * sub_par
    sr, sb = senior.par_amount, sub_par
    rows = [
        ("Sources of Funds:", None, None, True),
        ("  Par Value of Bonds", sr, sb, False),
        ("Total Sources of Funds", sr, sb, True),
        ("", None, None, False),
        ("Uses of Funds:", None, None, True),
        ("  Project Fund (Reimbursement)",
         su.reimbursement - (sb - uwd_sub), sb - uwd_sub, False),
        ("  Capitalized Interest Fund", capi, 0, False),
        ("  Debt Service Reserve Fund", senior.dsrf_deposit, 0, False),
        ("  Underwriter's Discount", uwd_sr, uwd_sub, False),
        ("  Costs of Issuance", cfg.coi, 0, False),
        ("Total Uses of Funds", sr, sb, True),
    ]
    _col_headers(ws, hr, [(1, "", 34), (2, "Senior\nBonds", 16),
                          (3, "Subordinate\nBonds", 16), (4, "Total", 16)])
    r = hr + 1
    for label, a, b, bold in rows:
        font = _BOLD if bold else _BODY
        fill = _TOTAL if bold and label.startswith("Total") else _WHITE
        _w(ws, r, 1, label, font=font, align=_L, fill=fill)
        if a is not None:
            _w(ws, r, 2, round(a), _DOLLAR, font=font, fill=fill)
            _w(ws, r, 3, round(b), _DOLLAR, font=font, fill=fill)
            _w(ws, r, 4, round(a + b), _DOLLAR, font=font, fill=fill)
        r += 1


# ── Exhibit A: Master Debt Service Fund Activity ─────────────────────────────
def _build_a(ws, cfg, sm, senior, surplus, sub, prefix="A", scenario_label=""):
    hr = _exhibit_header(ws, prefix, cfg, "Summary of Debt Service Fund Activity", 12, scenario_label)
    _col_headers(ws, hr, [
        (1, "Collection\nYear", 10),
        (2, f"Total\nTaxable Value\n(Exhibit {prefix}-1)", 15),
        (3, f"Net Tax\nRevenue\n(Exhibit {prefix}-1)", 13),
        (4, "Interest\nEarnings on\nSurplus (1)", 12),
        (5, f"Trustee\nFee\n${cfg.trustee_fee:,.0f}", 11),
        (6, "Total Net\nRevenue\n(Receipts)", 13),
        (7, f"Senior Bonds\nNet Debt Svc\n(Exhibit {prefix}-5)", 14),
        (8, f"Subordinate\nTransfers\n(Exhibit {prefix}-6)", 13),
        (9, "Surplus Fund\nBalance", 13),
        (10, "Total\nDisbursements", 13),
        (11, "Annual\nSurplus\n(Deficit)", 12),
        (12, "Debt Svc\nCoverage", 11),
    ])
    senior_net = senior.annual_net_ds()
    sub_by_year = {row["year"]: row for row in sub.rows}
    surplus_by_year = {row.year: row for row in surplus.rows}

    r = hr + 1
    prev_surplus_bal = 0.0
    for row in sm.rows:
        y = row.collection_year
        treas = -row.mill_revenue * cfg.county_collection_fee
        net_tax = row.mill_revenue + row.uniform_fee_revenue + treas
        int_earn = prev_surplus_bal * cfg.interest_earn_rate  # 1-yr lag
        trustee = -cfg.trustee_fee if net_tax else 0.0
        receipts = net_tax + int_earn + trustee
        sr_ds = senior_net.get(y, 0.0)
        sub_row = sub_by_year.get(y, {})
        sub_pay = sub_row.get("total_paid", 0.0)
        sf = surplus_by_year.get(y)
        surplus_bal = sf.reserve_balance if sf else 0.0
        disbursements = max(sr_ds, 0.0) + sub_pay
        annual_surplus = receipts - disbursements
        coverage = (receipts / sr_ds) if sr_ds > 0 else None

        fill = _GRAY if (r - hr) % 2 == 0 else _WHITE
        _w(ws, r, 1, y, font=_BODY, align=_C, fill=fill)
        _w(ws, r, 2, round(row.total_av) or None, _DOLLAR, fill=fill)
        _w(ws, r, 3, round(net_tax) or None, _DOLLAR, fill=fill)
        _w(ws, r, 4, round(int_earn) or None, _DOLLAR, fill=fill)
        _w(ws, r, 5, round(trustee) or None, _DOLLAR, fill=fill)
        _w(ws, r, 6, round(receipts) or None, _DOLLAR, fill=fill)
        _w(ws, r, 7, round(sr_ds) or None, _DOLLAR, fill=fill)
        _w(ws, r, 8, round(sub_pay) or None, _DOLLAR, fill=fill)
        _w(ws, r, 9, round(surplus_bal) or None, _DOLLAR, fill=fill)
        _w(ws, r, 10, round(disbursements) or None, _DOLLAR, fill=fill)
        _w(ws, r, 11, round(annual_surplus) or None, _DOLLAR, fill=fill)
        _w(ws, r, 12, coverage, _PCT1, fill=fill)
        prev_surplus_bal = surplus_bal
        r += 1
    _footer(ws, r + 1, 12, [
        "(1) The Surplus Fund earns interest at the assumed rate per annum; "
        "interest earnings are available on a one-year lag.",
        "** Totals may not equal the sum of the components due to rounding.",
        "This financial information should be read only in connection with the "
        "accompanying Summary of Significant Assumptions and Accounting Policies.",
    ])
    ws.freeze_panes = "B9"


def _add_scenario_sheets(wb, scn):
    """Add one exhibit set (Exhibit <prefix> and <prefix>-1..7) for a Scenario."""
    p, lbl = scn.exhibit, scn.label
    cfg = scn.cfg
    _build_a(wb.create_sheet(f"Exhibit {p}"), cfg, scn.sm, scn.senior,
             scn.surplus, scn.sub, p, lbl)
    _build_a1(wb.create_sheet(f"Exhibit {p}-1"), cfg, scn.sm, p, lbl)
    _build_a2(wb.create_sheet(f"Exhibit {p}-2"), cfg, scn.dev, scn.sm, p, lbl)
    _build_a3(wb.create_sheet(f"Exhibit {p}-3"), cfg, scn.dev, scn.sm, p, lbl)
    _build_a4(wb.create_sheet(f"Exhibit {p}-4"), cfg, scn.dev, p, lbl)
    _build_a5(wb.create_sheet(f"Exhibit {p}-5"), cfg, scn.senior, p, lbl)
    _build_a6(wb.create_sheet(f"Exhibit {p}-6"), cfg, scn.sub, p, lbl)
    _build_a7(wb.create_sheet(f"Exhibit {p}-7"), cfg, scn.senior, scn.su,
              scn.sub_par, p, lbl)


def build_forecast_report(
    scenarios,
    output_path: str = "output/ut_pid_forecast_exhibits.xlsx",
) -> str:
    """
    Generate the CPA-style forecast exhibit workbook.

    ``scenarios`` is a list of ``Scenario`` objects (see ``scenarios.py``): the
    base case (Exhibit A) plus any stress cases (Exhibits B, C, ...).  Each
    scenario contributes a full exhibit set (Exhibit <letter> and <letter>-1..7).
    """
    if not isinstance(scenarios, (list, tuple)):
        scenarios = [scenarios]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # drop the default sheet; scenarios add their own
    for scn in scenarios:
        _add_scenario_sheets(wb, scn)

    # Tierra logo + "preliminary, subject to change" footer on every exhibit.
    from .report import _apply_branding, _inject_footer_logos
    for ws in wb.worksheets:
        _apply_branding(ws)

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    wb.save(output_path)
    _inject_footer_logos(output_path)
    return output_path
