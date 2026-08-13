"""
residential_report.py — "Residential" tab for the model-view workbook.

Reformats the model's residential development into a single vertically-stacked
tab matching the source workbook's "Residential Development" sheet, with these
sections in order, each laid out ``Date | Existing | <product columns> | Total``
with a **Totals row above** the data:

  1. Lot Delivery — Units
  2. Lot Delivery — Value
  3. Finished Lots Used (lots remaining, drawn down as homes close)
  4. Home Closings — Units
  5. Home Closings — Value
  6. Residential Pricing (average selling price)
  7. Residential Taxable Value Creation

One column appears per product line (1 to 8 — e.g. one builder/one product, or
four builders with two products each).  In aggregate mode a single "All
Products" column is shown.
"""

from __future__ import annotations

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

DATE_COL = 1
EXIST_COL = 2
PROD_START = 3

_BLUE = PatternFill("solid", fgColor="1F4E79")
_SECT = PatternFill("solid", fgColor="2E6FA3")
_HDRF = PatternFill("solid", fgColor="D6E4F0")
_TOTF = PatternFill("solid", fgColor="BDD7EE")
_GRAY = PatternFill("solid", fgColor="F2F2F2")
_WHITE = PatternFill("solid", fgColor="FFFFFF")

_TITLEFONT = Font(bold=True, color="FFFFFF", size=12)
_SECTFONT = Font(bold=True, color="FFFFFF", size=10)
_HDRFONT = Font(bold=True, color="1F4E79", size=9)
_BODY = Font(size=9)
_TOTFONT = Font(bold=True, color="1F4E79", size=9)
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
_R = Alignment(horizontal="right", vertical="center")
_L = Alignment(horizontal="left", vertical="center")

_DOLLAR = '#,##0;(#,##0)'
_NUM = '#,##0'


def _cell(ws, r, c, v, fill, font=_BODY, fmt=None, align=_R):
    cell = ws.cell(row=r, column=c, value=v)
    cell.fill, cell.font, cell.border, cell.alignment = fill, font, _BORDER, align
    if fmt and v is not None:
        cell.number_format = fmt
    return cell


def _section(ws, start_row, title, names, years, columns, existing,
             fmt, total_mode="sum"):
    """
    columns: list (one per product) of {year: value} dicts.
    existing: {year: value} for the single Existing column.
    total_mode: 'sum' | 'avg' (weighted by row) | 'balance' (no column totals).
    """
    ncol = len(names)
    total_col = PROD_START + ncol
    last_col = total_col

    ws.merge_cells(start_row=start_row, start_column=DATE_COL,
                   end_row=start_row, end_column=last_col)
    t = ws.cell(row=start_row, column=DATE_COL, value=title)
    t.fill, t.font, t.alignment = _SECT, _SECTFONT, _L

    hr = start_row + 1
    _cell(ws, hr, DATE_COL, "Date", _HDRF, _HDRFONT, align=_C)
    _cell(ws, hr, EXIST_COL, "Existing", _HDRF, _HDRFONT, align=_C)
    for i, nm in enumerate(names):
        _cell(ws, hr, PROD_START + i, nm, _HDRF, _HDRFONT, align=_C)
    _cell(ws, hr, total_col, "Total", _HDRF, _HDRFONT, align=_C)

    # Totals row (above the data)
    tr = hr + 1
    _cell(ws, tr, DATE_COL, "Totals", _TOTF, _TOTFONT, align=_C)
    tot_exist = sum(existing.get(y, 0) for y in years)
    _cell(ws, tr, EXIST_COL, tot_exist or None, _TOTF, _TOTFONT, fmt=fmt)
    col_totals = []
    for i in range(ncol):
        if total_mode == "balance":
            cv = None
        elif total_mode == "avg":
            # weighted average not meaningful as a column sum; show blank
            cv = None
        else:
            cv = sum(columns[i].get(y, 0) for y in years)
        col_totals.append(cv or 0)
        _cell(ws, tr, PROD_START + i, cv or None, _TOTF, _TOTFONT, fmt=fmt)
    if total_mode == "balance" or total_mode == "avg":
        grand = None
    else:
        grand = tot_exist + sum(col_totals)
    _cell(ws, tr, total_col, grand or None, _TOTF, _TOTFONT, fmt=fmt)

    # Data rows
    r = tr + 1
    for j, y in enumerate(years):
        fill = _GRAY if j % 2 else _WHITE
        _cell(ws, r, DATE_COL, y, fill, align=_C)
        ev = existing.get(y, 0)
        _cell(ws, r, EXIST_COL, ev or None, fill, fmt=fmt)
        rowvals = [columns[i].get(y, 0) for i in range(ncol)]
        for i, v in enumerate(rowvals):
            _cell(ws, r, PROD_START + i, v or None, fill, fmt=fmt)
        if total_mode == "avg":
            # weighted average across products (weight = value itself unavailable);
            # show simple average of non-zero product values
            nz = [v for v in rowvals if v]
            total = (sum(nz) / len(nz)) if nz else 0
        else:
            total = ev + sum(rowvals)
        _cell(ws, r, total_col, total or None, fill, fmt=fmt)
        r += 1
    return r + 1


def build_residential_sheet(ws, cfg, dev, sm):
    """
    Residential development — the development information from the inputs only:
    product setup plus the lot-delivery and home-closing schedules, on a single
    horizontal plane (one row per year, deliveries and closings across).
    Derived projections (taxable value, reassessment, etc.) live on Summary.
    """
    prods = dev.product_lines()                 # real products, or one synthetic
    per_product = dev.is_per_product
    names = [p.name[:18] for p in prods]
    ncol = len(names)

    deliv = [p.lot_deliveries for p in prods]
    close = [p.home_closings for p in prods]
    years = sorted(set().union(*([set(d) for d in deliv] + [set(c) for c in close])))
    years = [y for y in years
             if any(deliv[i].get(y, 0) or close[i].get(y, 0) for i in range(ncol))]
    if not years:
        years = [cfg.first_year]

    # ── Column plan: Year, then Lot Deliveries (per product + Total), then
    #    Home Closings (per product + Total).  Aggregate ⇒ one column each. ────
    cols = []   # (group, subheader, {year: value})
    def _tot(src):
        return {y: sum(src[i].get(y, 0) for i in range(ncol)) for y in years}
    if per_product:
        for i, nm in enumerate(names):
            cols.append(("Lot Deliveries — Units", nm, deliv[i]))
        cols.append(("Lot Deliveries — Units", "Total", _tot(deliv)))
        for i, nm in enumerate(names):
            cols.append(("Home Closings — Units", nm, close[i]))
        cols.append(("Home Closings — Units", "Total", _tot(close)))
    else:
        cols.append(("Lot Deliveries — Units", "", deliv[0]))
        cols.append(("Home Closings — Units", "", close[0]))

    last_col = 1 + len(cols)
    ws.column_dimensions[get_column_letter(1)].width = 12
    for c in range(2, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 15

    # ── Title ────────────────────────────────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    title = (f"{cfg.pid_name} — Development Projections" if cfg.pid_name
             else "Development Projections")
    b = ws.cell(row=1, column=1, value=title)
    b.fill, b.font, b.alignment = _BLUE, _TITLEFONT, _C
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    mode = (f"{ncol} product line(s)" if per_product else "aggregate (single stream)")
    n = ws.cell(row=2, column=1,
                value=f"Development inputs only — lot deliveries & home closings by year ({mode})")
    n.font = Font(italic=True, size=8, color="595959"); n.alignment = _C

    # ── Product setup ────────────────────────────────────────────────────────
    r = 4
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
    s = ws.cell(row=r, column=1, value="PRODUCT SETUP")
    s.fill, s.font, s.alignment = _SECT, _SECTFONT, _L
    r += 1
    _cell(ws, r, 1, "Parameter", _HDRF, _HDRFONT, align=_C)
    for i, nm in enumerate(names):
        _cell(ws, r, 2 + i, nm, _HDRF, _HDRFONT, align=_C)
    r += 1
    for label, vals, fmt in [
        ("Total Units", [p.total_units for p in prods], _NUM),
        ("Base ASP ($)", [p.asp_base for p in prods], _DOLLAR),
        ("ASP Base Year", [p.asp_base_year for p in prods], "0"),
    ]:
        _cell(ws, r, 1, label, _WHITE, _BODY, align=_L)
        for i, v in enumerate(vals):
            _cell(ws, r, 2 + i, v, _WHITE, _BODY, fmt=fmt)
        r += 1
    r += 1

    # ── Annual development schedule (single horizontal table) ────────────────
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
    s = ws.cell(row=r, column=1, value="ANNUAL DEVELOPMENT SCHEDULE (FROM INPUTS)")
    s.fill, s.font, s.alignment = _SECT, _SECTFONT, _L
    r += 1

    from itertools import groupby
    gh, sh = r, r + 1                          # group-header and sub-header rows
    ws.merge_cells(start_row=gh, start_column=1, end_row=sh, end_column=1)
    _cell(ws, gh, 1, "Year", _HDRF, _HDRFONT, align=_C)
    c = 2
    for grp, items in groupby(cols, key=lambda x: x[0]):
        items = list(items)
        if len(items) > 1:
            ws.merge_cells(start_row=gh, start_column=c, end_row=gh, end_column=c + len(items) - 1)
            _cell(ws, gh, c, grp, _HDRF, _HDRFONT, align=_C)
            for k, it in enumerate(items):
                _cell(ws, sh, c + k, it[1], _HDRF, _HDRFONT, align=_C)
        else:
            ws.merge_cells(start_row=gh, start_column=c, end_row=sh, end_column=c)
            _cell(ws, gh, c, grp, _HDRF, _HDRFONT, align=_C)
        c += len(items)
    r = sh + 1

    # Totals row
    _cell(ws, r, 1, "Totals", _TOTF, _TOTFONT, align=_C)
    for k, (_, _, vals) in enumerate(cols):
        _cell(ws, r, 2 + k, sum(vals.get(y, 0) for y in years) or None, _TOTF, _TOTFONT, fmt=_NUM)
    r += 1

    # Year rows
    for j, y in enumerate(years):
        fill = _GRAY if j % 2 else _WHITE
        _cell(ws, r, 1, y, fill, _BODY, align=_C)
        for k, (_, _, vals) in enumerate(cols):
            _cell(ws, r, 2 + k, vals.get(y, 0) or None, fill, _BODY, fmt=_NUM)
        r += 1

    # ── Lot → home conversion & value reconciliation ─────────────────────────
    # Demonstrates that a lot's value rolls OFF "lot inventory value" and the
    # finished home rolls ON "cumulative home value" exactly when it closes, so
    # the two are never double-counted (Total = lot inventory + homes).
    RLAST = 14
    for c in range(1, RLAST + 1):
        w = ws.column_dimensions[get_column_letter(c)].width
        if not w or w < 14:
            ws.column_dimensions[get_column_letter(c)].width = 12 if c == 1 else 15
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=RLAST)
    s = ws.cell(row=r, column=1, value="LOT → HOME CONVERSION & VALUE RECONCILIATION")
    s.fill, s.font, s.alignment = _SECT, _SECTFONT, _L
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=RLAST)
    note = ws.cell(row=r, column=1, value=(
        "A lot's value rolls off Lot Inventory Market Value and the finished home rolls on "
        "Cumulative Home Market Value when it closes — never both. Total Market Value = "
        "lot inventory + homes.  Value is ADDED in the closing year and appears ON THE "
        "ASSESSMENT ROLL the following year (Roll Year = closing + 1)."))
    note.font = Font(italic=True, size=8, color="595959"); note.alignment = _L
    ws.row_dimensions[r].height = 24
    r += 1

    gh, sh = r, r + 1
    groups2 = [("", ["Closing\nYear"]),
               ("VACANT LOTS (units)", ["Begin", "+ Delivered", "− Built to Homes", "End"]),
               ("LOT MARKET VALUE ($)", ["Begin", "+ Lot Value Delivered",
                                         "− Lot Value to Homes", "End (Lot Inventory Value)"]),
               ("HOME MARKET VALUE — ADDED (closing year)",
                ["+ Home Value Added", "Cumulative Home Value", "Total Resid. Value"]),
               ("ON THE ASSESSMENT ROLL (next year)",
                ["Roll Year\n(closing + 1)", "Total Resid.\nValue on Roll"])]
    c = 1
    for gname, subs in groups2:
        if gname:
            ws.merge_cells(start_row=gh, start_column=c, end_row=gh, end_column=c + len(subs) - 1)
            _cell(ws, gh, c, gname, _HDRF, _HDRFONT, align=_C)
            for k, sub in enumerate(subs):
                _cell(ws, sh, c + k, sub, _HDRF, _HDRFONT, align=_C)
        else:
            ws.merge_cells(start_row=gh, start_column=c, end_row=sh, end_column=c)
            _cell(ws, gh, c, subs[0], _HDRF, _HDRFONT, align=_C)
        c += len(subs)
    r = sh + 1

    yrs2 = sorted(set(dev.lot_deliveries) | set(dev.home_closings))
    if yrs2:
        y0, y1 = min(yrs2), max(yrs2)
        yend = y1
        while dev.vacant_lot_units(yend) > 0.5 and yend < y1 + 8:
            yend += 1
        recon_years = list(range(y0, yend + 1))
    else:
        recon_years = []

    tot_lot_deliv_val = tot_lot_to_homes = 0.0
    for j, y in enumerate(recon_years):
        fill = _GRAY if j % 2 else _WHITE
        begin = dev.vacant_lot_units(y - 1)
        delivered = dev.lot_deliveries.get(y, 0)
        built = dev.home_closings.get(y, 0)
        end = dev.vacant_lot_units(y)
        # Lot market value flow: opening vacant value + value of lots delivered,
        # less the value rolling into homes as they are built = closing vacant value.
        vac_begin = dev.vacant_lot_market_value(y - 1)
        lot_delivered_val = dev.lot_market_value.get(y, 0.0)
        vac_end = dev.vacant_lot_market_value(y)
        lot_to_homes = vac_begin + lot_delivered_val - vac_end   # value converted to homes
        tot_lot_deliv_val += lot_delivered_val
        tot_lot_to_homes += lot_to_homes
        home_add = dev.new_home_market_value(y)
        cum_home = dev.cumulative_home_market_value.get(y, 0.0)
        total_resid = vac_end + cum_home
        # The market value added in the closing year lands on the assessment roll
        # the FOLLOWING year, so the same total is shown against Roll Year = y + 1.
        rowv = [y,
                round(begin), round(delivered), -round(built), round(end),
                round(vac_begin), round(lot_delivered_val), -round(lot_to_homes), round(vac_end),
                round(home_add), round(cum_home), round(total_resid),
                y + 1, round(total_resid)]
        for k, v in enumerate(rowv):
            fmt = "0" if k in (0, 12) else (_NUM if k <= 4 else _DOLLAR)   # year cols: yyyy
            _cell(ws, r, 1 + k, (v if (v or k in (0, 12)) else None), fill, _BODY,
                  fmt=fmt, align=(_C if k in (0, 12) else _R))
        r += 1

    # Totals: units & lot value delivered vs rolled into homes (each nets the
    # vacant balance to ~$0 once all lots are built out).
    tot_deliv = round(sum(dev.lot_deliveries.values()))
    tot_built = round(sum(dev.home_closings.values()))
    for c in range(1, RLAST + 1):
        ws.cell(row=r, column=c).fill = _TOTF
    _cell(ws, r, 1, "Total", _TOTF, _TOTFONT, align=_C)
    _cell(ws, r, 3, tot_deliv, _TOTF, _TOTFONT, fmt=_NUM)
    _cell(ws, r, 4, -tot_built, _TOTF, _TOTFONT, fmt=_NUM)
    _cell(ws, r, 7, round(tot_lot_deliv_val), _TOTF, _TOTFONT, fmt=_DOLLAR)
    _cell(ws, r, 8, -round(tot_lot_to_homes), _TOTF, _TOTFONT, fmt=_DOLLAR)


    # ── Balance flag: total lots delivered must equal total homes completed ───
    r += 1
    delta = tot_deliv - tot_built
    if delta == 0:
        msg = (f"✓ BALANCED — {tot_deliv:,} lots delivered = {tot_built:,} homes completed; "
               "lot-inventory value clears to $0 (no double-count, no residual).")
        ff = PatternFill("solid", fgColor="C6EFCE"); fn = Font(bold=True, size=9, color="006100")
    else:
        msg = (f"⚠ MISMATCH — {tot_deliv:,} lots delivered vs {tot_built:,} homes completed "
               f"(Δ {delta:+,}). Unbuilt lots leave a residual lot-inventory market value in every "
               "later year. Set total lot deliveries = total home closings.")
        ff = PatternFill("solid", fgColor="FFC7CE"); fn = Font(bold=True, size=9, color="9C0006")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=RLAST)
    fc = ws.cell(row=r, column=1, value=msg); fc.fill = ff; fc.font = fn; fc.alignment = _L
    for c in range(1, RLAST + 1):
        ws.cell(row=r, column=c).fill = ff
    ws.row_dimensions[r].height = 22

    # Per-product balance check (when products are entered individually).
    if per_product:
        bad = [(p.name, sum(p.lot_deliveries.values()), sum(p.home_closings.values()))
               for p in prods
               if round(sum(p.lot_deliveries.values())) != round(sum(p.home_closings.values()))]
        if bad:
            r += 1
            txt = "  Per-product mismatch: " + "; ".join(
                f"{nm}: {d:,} delivered vs {c:,} completed" for nm, d, c in bad)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=RLAST)
            pc = ws.cell(row=r, column=1, value=txt)
            pc.font = Font(bold=True, size=8, color="9C0006"); pc.alignment = _L

    # ── Lot-inventory taxable ratio applied ───────────────────────────────────
    # Utah has no separate lot-inventory class: what matters is whether the
    # primary residential exemption reaches builder-held inventory (Utah Admin.
    # Code R884-24P-52 says it can).  Show the ratio the model applies to
    # lot value by roll year (see the "Utah Property Tax Reference" reference tab),
    # with the residential ratio alongside for comparison (held flat at input).
    RATE_LAST = 3
    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=RATE_LAST)
    s = ws.cell(row=r, column=1, value="VACANT-LAND ASSESSMENT RATE APPLIED (BY ROLL YEAR)")
    s.fill, s.font, s.alignment = _SECT, _SECTFONT, _L
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=RATE_LAST)
    note = ws.cell(row=r, column=1, value=(
        "Utah taxes non-exempt property at 100% of fair market value; builder lot inventory "
        "carries the 45% primary residential exemption where the assessor determines the property "
        "will be a primary residence once occupied (Utah Admin. Code R884-24P-52), so it is taxed "
        "at the same 55%. Source: the Utah Property Tax Reference tab."))
    note.font = Font(italic=True, size=8, color="595959"); note.alignment = _L
    ws.row_dimensions[r].height = 24
    r += 1
    _cell(ws, r, 1, "Roll / Tax Year", _HDRF, _HDRFONT, align=_C)
    _cell(ws, r, 2, "Vacant-Land\nTaxable Ratio", _HDRF, _HDRFONT, align=_C)
    _cell(ws, r, 3, "Residential\nTaxable Ratio", _HDRF, _HDRFONT, align=_C)
    r += 1
    if recon_years:
        rate_years = list(range(min(recon_years), max(max(recon_years), 2027) + 1))
    else:
        rate_years = list(range(cfg.first_year, 2028))
    for j, y in enumerate(rate_years):
        fill = _GRAY if j % 2 else _WHITE
        _cell(ws, r, 1, y, fill, _BODY, fmt="0", align=_C)
        _cell(ws, r, 2, cfg.lot_inventory_ratio(y), fill, _BODY, fmt="0.000%")
        _cell(ws, r, 3, cfg.residential_assessment_rate(y + 1), fill, _BODY, fmt="0.000%")
        r += 1

    ws.freeze_panes = "A4"
