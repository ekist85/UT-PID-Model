"""
assessment_report.py — the special assessment tabs of the model workbook.

Three sheets, written only when Inputs!ASSESSMENT_BONDS says Yes:

  * **Assessment Bonds** — the program: what each series funds, what it is
    sized at, its sources and uses, and the value-to-lien test on each one.
  * **Assessment Cash Flow** — the annual roll: par issued, lots funded, homes
    converted, who pays each year's installment (the developer on lots it still
    holds, the homeowner on lots that have converted), and the balance
    outstanding.
  * **Assessment — Homeowner Cost** — what the assessment costs a buyer, per
    product: the balance that stays on the lot, the annual installment, and
    that installment as a share of the home's price.

The mill levy and the assessment are separate securities, so these tabs stand
apart from the senior/subordinate sheets — nothing here nets against the
revenue the senior lien is wrapped around.  The last block of the first sheet
is the one place the two are added up, because that is the number a homeowner
actually pays.
"""
from __future__ import annotations

from openpyxl.utils import get_column_letter

from .report import (_BODY, _BOLD, _BORDER, _CENTER, _DOLLAR, _GRAY, _GREEN,
                     _HDR_FONT, _LEFT, _LIGHT, _RIGHT, _TOTAL, _TOTAL_FONT,
                     _WHITE, _cell, _hdr, _title)

_RATE = '0.000"%"'
_MULT = '0.00"x"'
_NUM = '#,##0'
_DATE = 'm/d/yyyy'


def _section(ws, row, text, last_col):
    ws.cell(row=row, column=1, value=text).font = _HDR_FONT
    for c in range(1, last_col + 1):
        ws.cell(row=row, column=c).fill = _LIGHT
    ws.cell(row=row, column=1).font = _HDR_FONT
    return row + 1


def _drow(ws, row, label, value, fmt=_DOLLAR, fill=_WHITE, font=_BODY):
    _cell(ws, row, 1, label, fill, font, align=_LEFT)
    if value is not None:
        _cell(ws, row, 2, value, fill, font, fmt=fmt,
              align=_CENTER if fmt in (_DATE,) else _RIGHT)
    else:
        ws.cell(row=row, column=2).fill = fill
    return row + 1


# ── Sheet 1: the program ────────────────────────────────────────────────────

def build_assessment_sheet(ws, cfg, ab, sm=None, dev=None) -> None:
    """Program summary, the series table, and the value-to-lien test."""
    cols = [(1, 30), (2, 15), (3, 15), (4, 15), (5, 15), (6, 15), (7, 15),
            (8, 15), (9, 15), (10, 15)]
    for c, w in cols:
        ws.column_dimensions[get_column_letter(c)].width = w
    last_col = 10
    _title(ws, cfg, [
        cfg.pid_name,
        "Special Assessment Bonds — Assessment Area Act (UCA Title 11, Ch. 42)",
        f"{len(ab.series)} series  ·  {ab.total_lots:,.0f} lots assessed  ·  "
        f"{cfg.assessment_rate:.3%}  ·  {cfg.assessment_term_yrs}-year installments",
    ], last_col)

    r = 6
    r = _section(ws, r, "Program", last_col)
    prog = ab.programs[0] if ab.programs else None
    r = _drow(ws, r, "  Total Par Issued", ab.total_par)
    r = _drow(ws, r, "  Lots Assessed", ab.total_lots, fmt=_NUM)
    r = _drow(ws, r, "  Assessment per Lot (weighted)",
              ab.total_par / ab.total_lots if ab.total_lots else 0.0)
    r = _drow(ws, r, "  Capitalized Interest", ab.total_capitalized_interest
              if hasattr(ab, "total_capitalized_interest")
              else sum(s.capitalized_interest for s in ab.series))
    r = _drow(ws, r, "  Debt Service Reserve", ab.total_reserve)
    r = _drow(ws, r, "  Costs of Issuance", ab.total_coi)
    r = _drow(ws, r, "  Net Project Proceeds", ab.total_proceeds, fill=_GREEN, font=_BOLD)
    r = _drow(ws, r, "  Total Interest Over the Program", ab.total_interest)
    if prog is not None and getattr(prog, "units_unfunded_early", 0):
        r = _drow(ws, r, f"  Lots delivered before {prog.first_fundable_year} "
                         "(not financed by the program)",
                  prog.units_unfunded_early, fmt=_NUM)
    r += 1

    # ── The series ──────────────────────────────────────────────────────────
    r = _section(ws, r, "Assessment Series", last_col)
    heads = ["Series", "Dated", "Funds Deliveries", "Lots", "Assessment\nper Lot",
             "Par Amount", "Finished-Lot\nValue", "Value to\nLien", "Required",
             "Test"]
    for i, h in enumerate(heads, start=1):
        _hdr(ws, r, i, h)
    ws.row_dimensions[r].height = 30
    r += 1
    for i, s in enumerate(ab.series):
        fill = _GRAY if i % 2 else _WHITE
        _cell(ws, r, 1, f"Series {s.series_id}", fill, _BODY, align=_LEFT)
        _cell(ws, r, 2, s.dated_date, fill, _BODY, fmt=_DATE, align=_CENTER)
        _cell(ws, r, 3, f"{s.window_start}–{s.window_end}", fill, _BODY, align=_CENTER)
        _cell(ws, r, 4, s.lots, fill, _BODY, fmt=_NUM)
        _cell(ws, r, 5, s.lien_per_lot, fill, _BODY, fmt=_DOLLAR)
        _cell(ws, r, 6, s.par_amount, fill, _BODY, fmt=_DOLLAR)
        _cell(ws, r, 7, s.lot_value, fill, _BODY, fmt=_DOLLAR)
        _cell(ws, r, 8, s.value_to_lien, fill, _BODY, fmt=_MULT)
        _cell(ws, r, 9, s.required_value_to_lien, fill, _BODY, fmt=_MULT)
        _cell(ws, r, 10, "Pass" if s.meets_value_to_lien else "FAIL",
              fill, _BOLD, align=_CENTER)
        r += 1
    _cell(ws, r, 1, "TOTAL", _TOTAL, _TOTAL_FONT, align=_LEFT)
    for c, v, fmt in ((4, ab.total_lots, _NUM), (6, ab.total_par, _DOLLAR),
                      (7, sum(s.lot_value for s in ab.series), _DOLLAR),
                      (8, ab.weighted_value_to_lien, _MULT)):
        _cell(ws, r, c, v, _TOTAL, _TOTAL_FONT, fmt=fmt)
    for c in (2, 3, 5, 9, 10):
        ws.cell(row=r, column=c).fill = _TOTAL
    r += 2

    # ── Sources and uses, per series ────────────────────────────────────────
    r = _section(ws, r, "Sources & Uses by Series", last_col)
    for i, h in enumerate(["", "Par Amount", "Less: Costs of\nIssuance",
                           "Less: Capitalized\nInterest", "Less: Reserve\nFund",
                           "Net Project\nProceeds"], start=1):
        _hdr(ws, r, i, h)
    ws.row_dimensions[r].height = 30
    r += 1
    for i, s in enumerate(ab.series):
        fill = _GRAY if i % 2 else _WHITE
        _cell(ws, r, 1, f"Series {s.series_id}", fill, _BODY, align=_LEFT)
        for c, v in ((2, s.par_amount), (3, -s.coi), (4, -s.capitalized_interest),
                     (5, -s.reserve_fund), (6, s.project_proceeds)):
            _cell(ws, r, c, v, fill, _BODY, fmt=_DOLLAR)
        r += 1
    _cell(ws, r, 1, "TOTAL", _TOTAL, _TOTAL_FONT, align=_LEFT)
    for c, v in ((2, ab.total_par), (3, -ab.total_coi),
                 (4, -sum(s.capitalized_interest for s in ab.series)),
                 (5, -ab.total_reserve), (6, ab.total_proceeds)):
        _cell(ws, r, c, v, _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    r += 2

    # ── Assumptions ─────────────────────────────────────────────────────────
    r = _section(ws, r, "Assumptions", last_col)
    r = _drow(ws, r, "  Special Assessment Interest Rate", cfg.assessment_rate, fmt='0.000%')
    r = _drow(ws, r, "  Installment Term (years)", cfg.assessment_term_yrs, fmt=_NUM)
    r = _drow(ws, r, "  Capitalized Interest (years)", cfg.assessment_capi_years, fmt=_NUM)
    r = _drow(ws, r, "  Reserve Fund", cfg.assessment_dsrf, fmt=None)
    r = _drow(ws, r, "  Reserve % of Three-Prong Amount", cfg.assessment_dsrf_pct, fmt='0.00%')
    r = _drow(ws, r, "  Years Dated Ahead of the Lots Funded", cfg.assessment_fund_timing, fmt='0.0')
    r = _drow(ws, r, "  Delivery Years per Series", cfg.assessment_fund_term, fmt=_NUM)
    r += 1

    # ── What the homeowner actually pays ────────────────────────────────────
    # The only place the two securities are added together.
    r = _section(ws, r, "Total Homeowner Burden — Mill Levy plus Assessment", last_col)
    rows = ab.effective_tax_rate_rows()
    avg = rows[-1] if rows else None
    mills = float(getattr(cfg, "mill_levy_ds_target", 0.0) or 0.0)
    ratio = float(getattr(cfg, "resid_taxable_ratio", 0.55) or 0.55)
    if avg:
        price = float(avg["home_price"])
        levy = price * ratio * mills / 1000.0
        inst = float(avg["annual_installment"])
        r = _drow(ws, r, "  Average Home Price", price)
        r = _drow(ws, r, f"  Mill Levy ({mills:.3f} mills on {ratio:.0%} of market value)", levy)
        r = _drow(ws, r, "  Annual Assessment Installment", inst)
        r = _drow(ws, r, "  TOTAL Annual District Cost", levy + inst, fill=_GREEN, font=_BOLD)
        r = _drow(ws, r, "  As a % of Home Price",
                  (levy + inst) / price if price else 0.0, fmt='0.000%')
    ws.freeze_panes = "B6"


# ── Sheet 2: the annual cash flow ───────────────────────────────────────────

_FLOW_COLS = [
    ("year", "Year", _NUM, 8),
    ("lots_delivered", "Lots\nDelivered", _NUM, 11),
    ("lots_funded", "Lots\nFunded", _NUM, 11),
    ("par_issued", "Par\nIssued", _DOLLAR, 14),
    ("net_proceeds", "Net\nProceeds", _DOLLAR, 14),
    ("homes_converted", "Homes\nConverted", _NUM, 11),
    ("developer_payoff_total", "Developer\nPayoff", _DOLLAR, 14),
    ("balance_transferred_to_homeowners", "Balance to\nHomeowners", _DOLLAR, 15),
    ("scheduled_principal", "Principal", _DOLLAR, 14),
    ("interest", "Interest", _DOLLAR, 14),
    ("total_installments", "Total\nInstallments", _DOLLAR, 14),
    ("paid_from_capitalized_interest", "From CAPI\nFund", _DOLLAR, 13),
    ("developer_paid_installments", "Paid by\nDeveloper", _DOLLAR, 14),
    ("homeowner_paid_installments", "Paid by\nHomeowners", _DOLLAR, 15),
    ("outstanding_end", "Outstanding\n(end)", _DOLLAR, 15),
]


def build_assessment_cashflow_sheet(ws, cfg, ab) -> None:
    """The program's annual roll — one row a year, to the last installment."""
    last_col = len(_FLOW_COLS)
    for i, (_k, _h, _f, w) in enumerate(_FLOW_COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    _title(ws, cfg, [
        cfg.pid_name,
        "Special Assessment Bonds — Annual Cash Flow",
        "Installments are a lien on the lots, not a charge on the mill levy: "
        "nothing here is netted from the revenue the senior lien is sized against",
    ], last_col)

    r = 6
    for i, (_k, h, _f, _w) in enumerate(_FLOW_COLS, start=1):
        _hdr(ws, r, i, h)
    ws.row_dimensions[r].height = 30
    r += 1
    prog = ab.programs[0] if ab.programs else None
    df = prog.annual if prog is not None else None
    if df is None or df.empty:
        return
    n = 0
    for _idx, row in df.iterrows():
        fill = _GRAY if n % 2 else _WHITE
        for c, (key, _h, fmt, _w) in enumerate(_FLOW_COLS, start=1):
            v = row.get(key)
            v = float(v) if v is not None else 0.0
            if c == 1:
                _cell(ws, r, c, int(v), fill, _BODY, fmt=_NUM, align=_CENTER)
            else:
                _cell(ws, r, c, round(v) or None, fill, _BODY, fmt=fmt)
        r += 1
        n += 1
    _cell(ws, r, 1, "Total", _TOTAL, _TOTAL_FONT, align=_CENTER)
    for c, (key, _h, fmt, _w) in enumerate(_FLOW_COLS, start=1):
        if c == 1:
            continue
        if key == "outstanding_end":
            ws.cell(row=r, column=c).fill = _TOTAL
            continue
        _cell(ws, r, c, round(float(df[key].sum())) or None, _TOTAL, _TOTAL_FONT, fmt=fmt)
    ws.freeze_panes = "B7"


# ── Sheet 3: what it costs the buyer ────────────────────────────────────────

def build_assessment_homeowner_sheet(ws, cfg, ab) -> None:
    """Per product: the balance left on the lot and what it costs each year."""
    heads = ["Product", "Homes", "Home Price", "Developer\nAssessment",
             "Balance Left to\nthe Homeowner", "Annual\nInstallment",
             "% of Home\nPrice", "$ per $100\nof Value", "Mill Levy\n(annual)",
             "Total Annual\nDistrict Cost"]
    widths = [28, 10, 14, 14, 16, 13, 11, 11, 13, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    last_col = len(heads)
    _title(ws, cfg, [
        cfg.pid_name,
        "Special Assessment — Cost to the Homeowner",
        "The installment on the balance left on the lot at conversion, beside "
        "the mill levy on the finished home",
    ], last_col)

    mills = float(getattr(cfg, "mill_levy_ds_target", 0.0) or 0.0)
    ratio = float(getattr(cfg, "resid_taxable_ratio", 0.55) or 0.55)
    r = 6
    for i, h in enumerate(heads, start=1):
        _hdr(ws, r, i, h)
    ws.row_dimensions[r].height = 30
    r += 1
    rows = ab.effective_tax_rate_rows()
    for i, row in enumerate(rows):
        total = i == len(rows) - 1
        fill = _TOTAL if total else (_GRAY if i % 2 else _WHITE)
        font = _TOTAL_FONT if total else _BODY
        price = float(row["home_price"])
        levy = price * ratio * mills / 1000.0
        inst = float(row["annual_installment"])
        _cell(ws, r, 1, row["product"], fill, font, align=_LEFT)
        _cell(ws, r, 2, row["homes"], fill, font, fmt=_NUM)
        _cell(ws, r, 3, price, fill, font, fmt=_DOLLAR)
        _cell(ws, r, 4, row["developer_assessment"], fill, font, fmt=_DOLLAR)
        _cell(ws, r, 5, row["resident_assessment"], fill, font, fmt=_DOLLAR)
        _cell(ws, r, 6, inst, fill, font, fmt=_DOLLAR)
        _cell(ws, r, 7, row["effective_tax_rate"], fill, font, fmt='0.000%')
        _cell(ws, r, 8, row["rate_per_100_of_value"], fill, font, fmt='"$"0.000')
        _cell(ws, r, 9, levy, fill, font, fmt=_DOLLAR)
        _cell(ws, r, 10, levy + inst, fill, font, fmt=_DOLLAR)
        r += 1
    r += 1
    ws.cell(row=r, column=1, value=(
        "Developer Assessment is the lien placed on the lot when it is delivered; "
        "the balance left to the homeowner is what stays on it after the developer "
        "pays down at conversion. Where a Resident Annual Installment Payment is "
        "entered, that balance is SOLVED so the level payment to final maturity "
        "equals the entered figure.")).font = _BODY
    ws.freeze_panes = "B7"
