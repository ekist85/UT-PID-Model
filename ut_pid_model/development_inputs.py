"""
development_inputs.py — the Development Inputs sheet (write and read).

Utah owns this sheet outright.  Colorado's version carries a product setup
block and two unit grids; Utah's adds four things that have no Colorado
counterpart, so the layout lives here rather than as a patch against a
Colorado function it no longer resembles:

  * **Pacing Per Month** — a product's absorption rate.  A grid left blank is
    PRORATED from it instead of being read as "no units": homes close at
    ``pacing x 12`` a year from the year after the lots arrive, and lots
    deliver at the same rate from the delivery year, each capped by Total Units
    and (for closings) by the lots on the ground.  Entered grids always win;
    nothing is overwritten.
  * **Special Assessment** — the per-product assessment sizing assumptions
    (developer amount, resident amount, LTV, installment) that
    ``assessment_bonds`` sizes each series from.
  * **Existing Development** — homes, their ASP, and vacant lots already on the
    roll when the district is formed.
  * **A Total column** on each grid — enter a year's total and leave the
    product cells blank, and it is allocated across the products by their share
    of the units still to come.

Products sit in columns D..K with an aggregate Total in C, and every block is
found by its label rather than by a fixed row, so inserting a row on the sheet
does not silently shift which column a product is read from.  (The fixed-column
reader this replaces read C..J, so a workbook laid out this way lost its eighth
product without saying so.)
"""
from __future__ import annotations

import re
from datetime import date

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .development import DeveloperProjections, ProductLine

NPROD = 8                      # product columns on every block
FIRST_PROD_COL = 4             # column D — column C is the aggregate Total
_GRID_ROWS = 24                # year rows written on each grid


# ── Styles (shared with inputs.py) ──────────────────────────────────────────
_BLUE = PatternFill("solid", fgColor="1F4E79")
_SECT = PatternFill("solid", fgColor="2E6FA3")
_INPUT = PatternFill("solid", fgColor="FFF2CC")
_HDRF = PatternFill("solid", fgColor="D6E4F0")
_WHITEFONT = Font(bold=True, color="FFFFFF", size=12)
_SECTFONT = Font(bold=True, color="FFFFFF", size=10)
_HDRFONT = Font(bold=True, color="1F4E79", size=9)
_NOTE = Font(italic=True, size=9, color="595959")
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_L = Alignment(horizontal="left", vertical="center")
_R = Alignment(horizontal="right", vertical="center")
_C = Alignment(horizontal="center", vertical="center")
_NUM = '#,##0'
_DOLLAR = '#,##0'

#: Section labels, in the order they are written.  The loader anchors on these.
SETUP = "PRODUCT SETUP  (one column per product; up to 8)"
ASSESS = "SPECIAL ASSESSMENT  (per-product sizing; blank ⇒ the Inputs-tab figures)"
EXISTING = "EXISTING DEVELOPMENT  (already on the roll when the district is formed)"
LOTDEL = "LOT DELIVERY  (units by year, per product)"
HOMECLOS = "HOME CLOSINGS  (units by year, per product)"
COMMERCIAL = "COMMERCIAL MARKET VALUE  (by year; blank if none)"


# ── Reading helpers ─────────────────────────────────────────────────────────

def parse_ltv(raw, default: float | None = None) -> float | None:
    """
    Read an LTV / value-to-lien cell as a VALUE ÷ LIEN multiple.

    Accepts "3:1", "3,1", "3/1", "3 to 1", "3", 3.0, the lien-to-value fraction
    ("1:3", 0.333 → 3.0) and the time value Excel produces when "3:1" is typed
    into a General cell (03:01 → 3.0).  Blank or unreadable → ``default``.
    """
    from .assessment_bonds import parse_value_to_lien
    if raw in (None, ""):
        return default
    got = parse_value_to_lien(raw, 0.0)
    return got or default


def _num(v, default=None):
    try:
        return float(v) if v not in (None, "") else default
    except (TypeError, ValueError):
        return default


def _units(v) -> int:
    n = _num(v, 0.0) or 0.0
    return int(round(n))


# ── Writing ─────────────────────────────────────────────────────────────────

def write_dev_sheet(ws, dev: DeveloperProjections, cfg=None) -> None:
    """Lay out the Development Inputs sheet from ``dev`` (and ``cfg``'s
    per-product assessment table, when one is loaded)."""
    last_col = FIRST_PROD_COL + NPROD          # L — the computed Total column
    note_col = last_col + 1                    # M — right-hand notes
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 11
    for c in range(FIRST_PROD_COL, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.column_dimensions[get_column_letter(note_col)].width = 60

    def band(r, text, font, fill, align=_C, upto=last_col):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=upto)
        c = ws.cell(row=r, column=2, value=text)
        c.fill = fill
        c.font = font
        c.alignment = align
        return r + 1

    band(1, "Development Inputs", _WHITEFONT, _BLUE)
    band(2, "Per-product grids (up to 8 products). A grid left blank is paced from "
            "Pacing Per Month. Edit yellow cells.", _NOTE, PatternFill())

    products = list(dev.products)
    names = [(products[i].name if i < len(products) else None) for i in range(NPROD)]
    table = dict(getattr(cfg, "assessment_products", {}) or {}) if cfg else {}

    def cellw(r, c, v, hdr=False, inp=False, num=None, align=None):
        cell = ws.cell(row=r, column=c)
        if v is not None:
            cell.value = v
        cell.border = _BORDER
        if hdr:
            cell.fill = _HDRF
            cell.font = _HDRFONT
            cell.alignment = _C
        elif inp:
            cell.fill = _INPUT
            cell.alignment = align or _C
        else:
            cell.alignment = align or _R
        if num:
            cell.number_format = num

    def note(r, text):
        c = ws.cell(row=r, column=note_col, value=text)
        c.font = _NOTE
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

    def prod_header(r, first_label="Parameter"):
        cellw(r, 2, first_label, hdr=True)
        for i in range(NPROD):
            cellw(r, FIRST_PROD_COL + i, f"Product {i + 1}", hdr=True)
        return r + 1

    def rows(r, spec):
        for label, vals, num, align in spec:
            cellw(r, 2, label, align=_L)
            for i in range(NPROD):
                cellw(r, FIRST_PROD_COL + i, vals[i], inp=True, num=num, align=align)
            r += 1
        return r

    def pick(attr, default=None):
        return [getattr(products[i], attr, default) if i < len(products) else None
                for i in range(NPROD)]

    def from_table(key):
        return [(table.get(names[i], {}) or {}).get(key) if names[i] else None
                for i in range(NPROD)]

    # ── Product setup ───────────────────────────────────────────────────────
    r = band(4, SETUP, _SECTFONT, _SECT, align=_L)
    r = prod_header(r)
    r = rows(r, [
        ("Product Name", names, None, _L),
        ("Total Units", pick("total_units"), _NUM, _R),
        ("Base ASP ($)", pick("asp_base"), _DOLLAR, _R),
        ("ASP Base Year", pick("asp_base_year"), None, _C),
        ("Pacing Per Month", pick("pacing_per_month"), None, _C),
    ])
    note(r - 1, "units per month; paces a Lot Delivery or Home Closings column "
                "left blank. Blank ⇒ that grid must be entered")

    # ── Special assessment ──────────────────────────────────────────────────
    r = band(r + 1, ASSESS, _SECTFONT, _SECT, align=_L)
    r = prod_header(r)
    first_assess = r
    r = rows(r, [
        ("Developer Assessment Bond Initial Amount (Per Lot)", from_table("developer"), _DOLLAR, _R),
        ("Resident Assessment Bond Remaining Amount (Per Lot)", from_table("resident"), _DOLLAR, _R),
        ("LTV", from_table("ltv"), None, _C),
        ("Resident Annual Installment Payment", from_table("installment"), _DOLLAR, _R),
    ])
    note(first_assess, "$ Developer Bond per developed lot at lot delivery (par = this × lots "
                       "delivered). BLANK or 0 = size each series to the maximum the finished-lot "
                       "value supports at the Loan-to-Value ratio")
    note(first_assess + 1, "$ left on the lot for the homeowner after conversion; the developer pays "
                           "the balance down to this (plus accrued interest). BLANK or 0 = the "
                           "developer pays off the whole balance")
    note(first_assess + 2, "Required finished-lot value ÷ assessment, entered as X:Y or X,Y "
                           "(e.g. 3:1); reported as a test on each series, does not change the assessment")
    note(first_assess + 3, "$/yr the homeowner pays; entered, it SOLVES the balance left on each lot "
                           "so the level payment to final maturity equals this")

    # ── Existing development ────────────────────────────────────────────────
    r = band(r + 1, EXISTING, _SECTFONT, _SECT, align=_L)
    r = prod_header(r)
    r = rows(r, [
        ("Existing Homes (units)", pick("existing_homes"), _NUM, _R),
        ("Existing Home ASP ($)", pick("existing_home_asp"), _DOLLAR, _R),
        ("Existing Vacant Lots (units)", pick("existing_lots"), _NUM, _R),
    ])
    ws.cell(row=r, column=2, value="Existing Home ASP blank → uses Base ASP above. Existing lots "
                                   "carry the builder-inventory ratio and convert via Home "
                                   "Closings.").font = _NOTE
    r += 2

    # ── The two unit grids ──────────────────────────────────────────────────
    prod_years = set().union(*[set(p.lot_deliveries) | set(p.home_closings)
                               for p in products]) if products else set()
    start = min(prod_years) if prod_years else dev.asp_base_year
    end = max(prod_years) if prod_years else dev.asp_base_year
    grid_years = list(range(start, max(end + 3, start + _GRID_ROWS) + 1))

    def grid(r, title, getter):
        r = band(r, title, _SECTFONT, _SECT, align=_L)
        cellw(r, 2, "Year", hdr=True)
        cellw(r, 3, "Total", hdr=True)
        for i in range(NPROD):
            cellw(r, FIRST_PROD_COL + i, f"Product {i + 1}", hdr=True)
        cellw(r, last_col, "Total", hdr=True)
        r += 1
        first = r
        for y in grid_years:
            vals = [getter(products[i], y) if i < len(products) else None
                    for i in range(NPROD)]
            cellw(r, 2, y, inp=True, align=_C)
            cellw(r, 3, None, inp=True, num=_NUM)
            for i in range(NPROD):
                cellw(r, FIRST_PROD_COL + i, vals[i] or None, inp=True, num=_NUM)
            tc = ws.cell(row=r, column=last_col)
            tc.value = (f"=SUM({get_column_letter(FIRST_PROD_COL)}{r}:"
                        f"{get_column_letter(last_col - 1)}{r})")
            tc.border = _BORDER
            tc.number_format = _NUM
            r += 1
        ws.cell(row=r, column=2, value="Total").font = _HDRFONT
        tc = ws.cell(row=r, column=last_col)
        tc.value = f"=SUM({get_column_letter(last_col)}{first}:{get_column_letter(last_col)}{r - 1})"
        tc.font = _HDRFONT
        tc.number_format = _NUM
        return r + 2

    r = grid(r, LOTDEL, lambda p, y: p.lot_deliveries.get(y))
    r = grid(r, HOMECLOS, lambda p, y: p.home_closings.get(y))

    # ── Commercial ──────────────────────────────────────────────────────────
    r = band(r, COMMERCIAL, _SECTFONT, _SECT, align=_L)
    cellw(r, 2, "Year", hdr=True)
    cellw(r, FIRST_PROD_COL, "Commercial Market Value", hdr=True)
    r += 1
    years = sorted(dev.commercial_market_value) or list(range(start, start + 7))
    for y in years:
        cellw(r, 2, y, align=_C)
        cellw(r, FIRST_PROD_COL, dev.commercial_market_value.get(y), inp=True, num=_DOLLAR)
        r += 1

    ws.freeze_panes = "D6"


# ── Reading ─────────────────────────────────────────────────────────────────

#: Row labels the loader anchors on, longest-first so "Existing Home ASP" is not
#: swallowed by "Existing Home".
_ROW_ANCHORS = [
    ("name", "product name"),
    ("units", "total units"),
    ("asp", "base asp"),
    ("aspyr", "asp base year"),
    ("pacing", "pacing per month"),
    ("sa_dev", "developer assessment bond"),
    ("sa_res", "resident assessment bond"),
    ("sa_inst", "resident annual installment"),
    ("sa_ltv", "ltv"),
    ("exasp", "existing home asp"),
    ("exhomes", "existing homes"),
    ("exlots", "existing vacant lots"),
]

#: Section markers, matched on the start of the label.
_SECTION_ANCHORS = [
    ("lotdel", "lot delivery"),
    ("homeclos", "home closings"),
    ("comm", "commercial market value"),
    ("setup", "product setup"),
    ("assess", "special assessment"),
    ("existing", "existing development"),
]


def _anchors(ws) -> dict:
    found: dict[str, int] = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=2).value
        if not isinstance(v, str):
            continue
        lab = v.strip().lower()
        for key, prefix in _ROW_ANCHORS + _SECTION_ANCHORS:
            if key not in found and lab.startswith(prefix):
                found[key] = r
                break
    return found


def _product_columns(ws, header_row: int) -> list[int]:
    """The columns carrying 'Product 1'..'Product N' on a header row."""
    found = {}
    if header_row and header_row >= 1:
        for c in range(2, min(ws.max_column, 40) + 1):
            v = ws.cell(row=header_row, column=c).value
            if isinstance(v, str):
                m = re.match(r"^\s*product\s*(\d+)\s*$", v, re.I)
                if m:
                    found[int(m.group(1))] = c
    return [found[i] for i in sorted(found)] if found else \
        list(range(FIRST_PROD_COL, FIRST_PROD_COL + NPROD))


def _read_grid(ws, title_row, index_by_col, section_rows, caps=None):
    """
    Read a units-by-year grid into ``{year: {column: units}}``.

    A year whose product cells are blank but whose aggregate **Total** column is
    filled is allocated across the products by their share of the units still
    to come (largest-remainder, never past a product's Total Units) — so a
    schedule can be entered in one column when the product split is not known.
    """
    out: dict[int, dict[int, int]] = {}
    if not title_row:
        return out
    hdr = title_row + 1
    grid_cols, total_cols = {}, []
    for c in range(2, min(ws.max_column, 40) + 1):
        v = ws.cell(row=hdr, column=c).value
        if not isinstance(v, str):
            continue
        m = re.match(r"^\s*product\s*(\d+)\s*$", v, re.I)
        if m and m.group(1) and int(m.group(1)) in index_by_col:
            grid_cols[c] = index_by_col[int(m.group(1))]
        elif v.strip().lower() == "total":
            total_cols.append(c)
    if not grid_cols:                       # header carries product NAMES
        cols = _product_columns(ws, hdr) or list(range(FIRST_PROD_COL, FIRST_PROD_COL + NPROD))
        grid_cols = {c: index_by_col.get(i + 1, c)
                     for i, c in enumerate(cols) if (i + 1) in index_by_col}
    placed = {col: 0 for col in grid_cols.values()}
    for r in range(hdr + 1, ws.max_row + 1):
        if r in section_rows:
            break
        y = ws.cell(row=r, column=2).value
        if not isinstance(y, (int, float)) or not (1900 < y < 2200):
            lbl = ws.cell(row=r, column=2).value
            if isinstance(lbl, str) and lbl.strip().lower() == "total":
                break
            continue
        yi = int(y)
        row = {}
        for c, setup_col in grid_cols.items():
            n = _units(ws.cell(row=r, column=c).value)
            if n:
                row[setup_col] = row.get(setup_col, 0) + n
        if row:
            for col, n in row.items():
                placed[col] = placed.get(col, 0) + n
            out[yi] = {**out.get(yi, {}), **row}
            continue
        agg = 0
        for c in total_cols:
            v = ws.cell(row=r, column=c).value
            if isinstance(v, (int, float)):
                agg = max(agg, _units(v))
        if agg and caps:
            share = _allocate(agg, {col: max(caps.get(col, 0) - placed.get(col, 0), 0)
                                    for col in grid_cols.values()})
            if share:
                for col, n in share.items():
                    placed[col] = placed.get(col, 0) + n
                out[yi] = {**out.get(yi, {}), **share}
    return out


def _allocate(total: int, room: dict[int, int]) -> dict[int, int]:
    """Split ``total`` across columns by their remaining room, largest remainder."""
    pool = sum(max(v, 0) for v in room.values())
    if total <= 0 or pool <= 0:
        return {}
    total = min(total, pool)
    exact = {k: total * max(v, 0) / pool for k, v in room.items()}
    out = {k: int(v) for k, v in exact.items()}
    left = total - sum(out.values())
    for k in sorted(exact, key=lambda k: exact[k] - out[k], reverse=True):
        if left <= 0:
            break
        if out[k] < max(room.get(k, 0), 0):
            out[k] += 1
            left -= 1
    return {k: v for k, v in out.items() if v}


def _pace(grid: dict[int, dict[int, int]], col: int, pace: float, cap: int,
          first_year: int, available: dict[int, int] | None) -> dict[int, int]:
    """
    Fill one product's column from its monthly pace.

    ``pace`` units a month is ``pace x 12`` a year, starting ``first_year``,
    stopping at ``cap`` units.  When ``available`` is given (closings paced
    against lots), a year never closes more homes than there are lots delivered
    through the previous year.
    """
    annual = pace * 12.0
    if annual <= 0 or cap <= 0:
        return {}
    placed, out, y = 0, {}, first_year
    for _ in range(200):
        if placed >= cap:
            break
        room = cap - placed
        if available is not None:
            ready = sum(n for yy, n in available.items() if yy <= y - 1) - placed
            room = min(room, max(int(ready), 0))
        take = int(min(annual, room))
        if take > 0:
            out[y] = take
            placed += take
        elif available is None:
            break
        y += 1
        if y > first_year + 199:
            break
    if out:
        grid.setdefault("_paced", {})
    return out


def load_dev_sheet(wb, overrides_dev: dict, cfg=None) -> DeveloperProjections:
    """
    Read the Development Inputs sheet into ``DeveloperProjections``.

    Also fills ``cfg.assessment_products`` — the per-product assessment sizing
    assumptions ``assessment_bonds`` sizes each series from — when a config is
    passed and the Special Assessment block carries anything.
    """
    kwargs = dict(overrides_dev)
    if "Development Inputs" not in wb.sheetnames:
        return DeveloperProjections(**kwargs)
    ws = wb["Development Inputs"]
    a = _anchors(ws)
    section_rows = {a[k] for k in ("setup", "assess", "existing", "lotdel",
                                   "homeclos", "comm") if k in a}

    cols = _product_columns(ws, a.get("name", 0) - 1 if a.get("name") else 0)
    index_by_col = {i: c for i, c in enumerate(cols, start=1)}

    def cell(key, c, default=None):
        row = a.get(key)
        return _num(ws.cell(row=row, column=c).value, default) if row else default

    def text(key, c):
        row = a.get(key)
        v = ws.cell(row=row, column=c).value if row else None
        return str(v).strip() if v not in (None, "") else None

    # ── Products ────────────────────────────────────────────────────────────
    specs, caps = [], {}
    for i, c in enumerate(cols, start=1):
        name = text("name", c)
        if not name:
            continue
        total_units = _units(cell("units", c, 0))
        specs.append({
            "col": c, "name": name,
            "total_units": total_units,
            "asp_base": cell("asp", c, 0.0) or 0.0,
            "asp_base_year": int(cell("aspyr", c, 0) or 0) or None,
            "pacing_per_month": cell("pacing", c, 0.0) or 0.0,
            "existing_homes": _units(cell("exhomes", c, 0)),
            "existing_home_asp": cell("exasp", c, 0.0) or 0.0,
            "existing_lots": _units(cell("exlots", c, 0)),
        })
        caps[c] = total_units
    if not specs:
        return DeveloperProjections(**kwargs)

    lot_grid = _read_grid(ws, a.get("lotdel"), index_by_col, section_rows, caps)
    home_grid = _read_grid(ws, a.get("homeclos"), index_by_col, section_rows, caps)

    def column(grid, c):
        return {y: by.get(c, 0) for y, by in sorted(grid.items()) if by.get(c)}

    # ── Products, pacing the grids a product left blank ─────────────────────
    products, paced = [], []
    default_first = min(lot_grid) if lot_grid else (
        int(cell("aspyr", specs[0]["col"], 0) or 0) or date.today().year)
    for s in specs:
        lots = column(lot_grid, s["col"])
        homes = column(home_grid, s["col"])
        cap = s["total_units"] or 0
        pace = s["pacing_per_month"] or 0.0
        if not lots and pace > 0 and cap > 0:
            lots = _pace({}, s["col"], pace, cap, default_first, None)
            paced.append(f"{s['name']}: lots {pace:g}/mo from {default_first}")
        if not homes and pace > 0 and lots:
            cap_h = cap or sum(lots.values())
            homes = _pace({}, s["col"], pace, min(cap_h, sum(lots.values())),
                          min(lots) + 1, lots)
            paced.append(f"{s['name']}: homes {pace:g}/mo from {min(lots) + 1}")
        products.append(ProductLine(
            name=s["name"],
            existing_units=cap,
            lot_deliveries=lots,
            home_closings=homes,
            asp_base=s["asp_base"] or 515_000.0,
            asp_base_year=s["asp_base_year"] or 2024,
            pacing_per_month=pace,
            existing_homes=s["existing_homes"],
            existing_home_asp=s["existing_home_asp"],
            existing_lots=s["existing_lots"],
        ))
    if paced:
        print("[dev] grids paced from Pacing Per Month — " + "; ".join(paced) + ".")

    # ── Per-product assessment sizing assumptions ───────────────────────────
    if cfg is not None:
        table = {}
        for s in specs:
            c = s["col"]
            entry = {
                "developer": cell("sa_dev", c),
                "resident": cell("sa_res", c),
                "ltv": parse_ltv(ws.cell(row=a["sa_ltv"], column=c).value) if "sa_ltv" in a else None,
                "installment": cell("sa_inst", c),
            }
            if any(v for v in entry.values()):
                table[s["name"]] = entry
        if table:
            cfg.assessment_products = table

    # ── Commercial ──────────────────────────────────────────────────────────
    commercial = {}
    cstart = a.get("comm")
    if cstart:
        for r in range(cstart + 2, ws.max_row + 1):
            if r in section_rows:
                break
            y = ws.cell(row=r, column=2).value
            v = None
            for c in range(3, FIRST_PROD_COL + 2):
                cv = ws.cell(row=r, column=c).value
                if isinstance(cv, (int, float)) and cv:
                    v = float(cv)
                    break
            if isinstance(y, (int, float)) and v:
                commercial[int(y)] = v

    kwargs["products"] = products
    if commercial:
        kwargs["commercial_market_value"] = commercial
    # Total Units is the build-out count the deal is named for ("597 Lots"), so
    # it comes from the products rather than keeping a default from a different
    # deal.  Fall back to the delivered lots when Total Units is not entered.
    kwargs.setdefault("total_lots", sum(p.total_units for p in products)
                      or sum(sum(p.lot_deliveries.values()) for p in products))
    return DeveloperProjections(**kwargs)
