"""
report.py — Multi-sheet Excel workbook generator.

Produces a formatted .xlsx that mirrors the structure of the source workbook:

  1. Summary            — taxable value & pledged revenue by year
  2. Senior Lien DS     — sized senior new-money serial schedule
  3. Sources & Uses     — first financing + reimbursement
  4. Refunding          — refinancing of the senior new-money bonds + new money
  5. Subordinate Lien   — cash-flow waterfall
"""

from __future__ import annotations

import os
from datetime import date as _date

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as _XLImage

from .config import ModelConfig, residential_taxable_ratio_for

_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "tierra_logo.png")


# ── Bond-statistics helpers ───────────────────────────────────────────────────
def _avg_life(principal_by_date, delivery, par) -> float:
    """Average life (years) = Σ principal·years / total principal."""
    if not par:
        return 0.0
    return sum(p * ((d - delivery).days / 365.25) for d, p in principal_by_date) / par


def _tic(payments, proceeds, delivery):
    """
    True interest cost: the annual rate (semiannual compounding) at which the
    present value of the gross debt-service ``payments`` equals ``proceeds``.
    Solved by bisection.  ``payments`` is a list of (date, amount).
    """
    if proceeds <= 0 or not payments:
        return None

    def pv(rate):
        return sum(cf / (1 + rate / 2) ** (2 * ((d - delivery).days / 365.25))
                   for d, cf in payments)

    lo, hi = 0.0, 0.50
    for _ in range(200):
        mid = (lo + hi) / 2
        if pv(mid) > proceeds:   # PV too high ⇒ rate too low
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


# ── Branding: Tierra logo + "preliminary, subject to change" footer ──────────
def _apply_branding(ws):
    """
    Set the print footer on every page: the Tierra logo (left, injected as a
    header/footer image post-save via ``&L&G``) plus the preliminary notice.
    """
    ws.oddFooter.left.text = "&G"
    ws.oddFooter.center.text = "PRELIMINARY — SUBJECT TO CHANGE"
    ws.oddFooter.right.text = "Prepared by Tierra"
    ws.evenFooter.left.text = "&G"
    ws.evenFooter.center.text = "PRELIMINARY — SUBJECT TO CHANGE"
    ws.evenFooter.right.text = "Prepared by Tierra"


_VML_HF = """<xml xmlns:v="urn:schemas-microsoft-com:vml" \
xmlns:o="urn:schemas-microsoft-com:office:office" \
xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="{n}"/></o:shapelayout>
 <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" \
path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
  <v:stroke joinstyle="miter"/>
  <v:formulas>
   <v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/>
   <v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/>
   <v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/>
   <v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/>
  </v:formulas>
  <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
  <o:lock v:ext="edit" aspectratio="t"/>
 </v:shapetype>
 <v:shape id="LF" o:spid="_x0000_s{spid}" type="#_x0000_t75" \
style='position:absolute;margin-left:0;margin-top:0;width:120pt;height:33pt;z-index:1'>
  <v:imagedata o:relid="rId1" o:title="Tierra"/>
  <o:lock v:ext="edit" rotation="t"/>
 </v:shape>
</xml>"""

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _inject_footer_logos(path):
    """
    Embed the Tierra logo as a print header/footer image (the ``&L&G`` code in
    the footer) on every sheet.  openpyxl can't do this, so the saved workbook
    is post-processed: each sheet gets a VML header/footer drawing pointing at a
    shared logo image, wired up with the right relationships and content types.
    """
    import zipfile
    import re
    if not os.path.exists(_LOGO_PATH):
        return
    with open(_LOGO_PATH, "rb") as f:
        logo = f.read()

    with zipfile.ZipFile(path, "r") as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}

    ct = parts["[Content_Types].xml"].decode("utf-8")
    for ext, mime in [("png", "image/png"),
                      ("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing")]:
        if f'Extension="{ext}"' not in ct:
            ct = ct.replace("</Types>", f'<Default Extension="{ext}" ContentType="{mime}"/></Types>')
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    parts["xl/media/tierra_hf.png"] = logo

    sheets = sorted((n for n in parts if re.match(r"xl/worksheets/sheet\d+\.xml$", n)),
                    key=lambda n: int(re.search(r"(\d+)", n).group(1)))
    for sheet in sheets:
        n = int(re.search(r"sheet(\d+)\.xml$", sheet).group(1))
        parts[f"xl/drawings/vmlDrawingHF{n}.vml"] = _VML_HF.format(n=n, spid=1024 + n).encode("utf-8")
        parts[f"xl/drawings/_rels/vmlDrawingHF{n}.vml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{_R_NS}/image" Target="../media/tierra_hf.png"/>'
            '</Relationships>').encode("utf-8")

        rels_name = f"xl/worksheets/_rels/sheet{n}.xml.rels"
        if rels_name in parts:
            rels = parts[rels_name].decode("utf-8")
            ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
            rid = f"rId{(max(ids) + 1) if ids else 1}"
            rels = rels.replace("</Relationships>",
                f'<Relationship Id="{rid}" Type="{_R_NS}/vmlDrawing" '
                f'Target="../drawings/vmlDrawingHF{n}.vml"/></Relationships>')
        else:
            rid = "rId1"
            rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="{rid}" Type="{_R_NS}/vmlDrawing" '
                    f'Target="../drawings/vmlDrawingHF{n}.vml"/></Relationships>')
        parts[rels_name] = rels.encode("utf-8")

        sx = parts[sheet].decode("utf-8")
        if "<legacyDrawingHF" not in sx:
            sx = sx.replace("</worksheet>",
                f'<legacyDrawingHF xmlns:r="{_R_NS}" r:id="{rid}"/></worksheet>')
        parts[sheet] = sx.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in parts.items():
            zout.writestr(n, data)

# ── Styles ──────────────────────────────────────────────────────────────────
_BLUE  = PatternFill("solid", fgColor="1F4E79")
_LIGHT = PatternFill("solid", fgColor="D6E4F0")
_GRAY  = PatternFill("solid", fgColor="F2F2F2")
_WHITE = PatternFill("solid", fgColor="FFFFFF")
_TOTAL = PatternFill("solid", fgColor="BDD7EE")
_GREEN = PatternFill("solid", fgColor="E2EFDA")
_HILITE = PatternFill("solid", fgColor="FFF2CC")   # highlight (existing land value)

_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_TITLE_FONT  = Font(name="Calibri", bold=True, color="FFFFFF", size=13)
_WHITE_FONT  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
_HDR_FONT    = Font(name="Calibri", bold=True, color="1F4E79", size=9)
_BODY        = Font(name="Calibri", size=9)
_BOLD        = Font(name="Calibri", bold=True, size=9)
_TOTAL_FONT  = Font(name="Calibri", bold=True, color="1F4E79", size=9)
_RATE_FONT   = Font(name="Calibri", italic=True, size=7, color="808080")  # taxable ratios
# Blue font marks figures that came from the certified Inputs template — certified
# taxable values and historical (already-set) taxable ratios — so the user can
# tell inputs-driven numbers from the model's projections at a glance.
_CERT_COLOR  = "0070C0"
_CERT_FONT   = Font(name="Calibri", size=9, color=_CERT_COLOR)
_CERT_BOLD   = Font(name="Calibri", bold=True, size=9, color=_CERT_COLOR)
_RATEFMT     = "0.000%"

_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_RIGHT  = Alignment(horizontal="right", vertical="center")
_LEFT   = Alignment(horizontal="left", vertical="center")

_DOLLAR = '"$"#,##0'
_NUM    = "#,##0"
_PCT    = '0.00"%"'


def _county_line(cfg, suffix: str = "") -> str:
    """Location subtitle — omits the county phrase when no county is set."""
    base = f"{cfg.county} County, Utah" if cfg.county else "Utah"
    return base + suffix


def _title(ws, cfg, lines: list[str], last_col: int):
    """
    Standard three-row title band, in the order requested for the deliverables:
      1. today's date
      2. Reimbursement Analysis · {district} · {N} Lots · Tierra Financial Advisors
      3. the sheet's own subtitle / description (``lines[1:]`` — ``lines[0]`` is the
         district name, now folded into row 2).
    Kept to three physical rows so every sheet's header row (row 5) is unaffected.
    """
    from datetime import date as _dt
    _t = _dt.today()
    today = f"{_t.strftime('%B')} {_t.day}, {_t.year}"
    lots = getattr(cfg, "_doc_lots", None)
    lot_txt = f"{lots:,} Lots" if lots else None
    brand = "  ·  ".join(x for x in [
        "Reimbursement Analysis", cfg.pid_name or None, lot_txt,
        "Tierra Financial Advisors"] if x)
    subtitle = " — ".join(str(x) for x in lines[1:] if x)
    rows = [today, brand, subtitle]
    for i, text in enumerate(rows, 1):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=last_col)
        c = ws.cell(row=i, column=1, value=text)
        c.fill = _BLUE
        c.font = _TITLE_FONT if i == 1 else _WHITE_FONT
        c.alignment = _CENTER
    ws.row_dimensions[len(rows) + 1].height = 6


def _hdr(ws, row, col, label, width=None):
    c = ws.cell(row=row, column=col, value=label)
    c.font, c.fill, c.border, c.alignment = _HDR_FONT, _LIGHT, _BORDER, _CENTER
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return c


def _cell(ws, row, col, value, fill, font=None, fmt=None, align=_RIGHT):
    c = ws.cell(row=row, column=col, value=value)
    c.fill, c.font, c.border, c.alignment = fill, font or _BODY, _BORDER, align
    if fmt:
        c.number_format = fmt
    return c


# ── Sheet 1: Summary ──────────────────────────────────────────────────────────
def _build_summary_sheet(ws, cfg, sm):
    # Optional AV component columns appear only when those values are present.
    base = [(1, "Collection\nYear", 12), (2, "Total\nTaxable Value", 16),
            (3, "Residential\nTaxable Value", 15), (4, "Lot Inventory\nTaxable Value", 14)]
    n = 4
    og = comm = None
    if sm.has_centrally_assessed:
        n += 1; og = n; base.append((n, "Centrally Assessed\nTaxable Value", 14))
    if sm.has_commercial:
        n += 1; comm = n; base.append((n, "Commercial\nTaxable Value", 14))
    mr, nsr, nsub = n + 1, n + 2, n + 3
    base += [(mr, "Mill Levy\nRevenue", 15), (nsr, "Net Senior\nRevenue", 15),
             (nsub, "Net Sub.\nRevenue", 15)]
    _title(ws, cfg, [cfg.pid_name, _county_line(cfg),
                "Taxable Value & Pledged Revenue"], nsub)
    for col, lbl, w in base:
        _hdr(ws, 5, col, lbl, w)
    for i, r in enumerate(sm.rows):
        rw = 6 + i
        fill = _GRAY if i % 2 else _WHITE
        _cell(ws, rw, 1, r.collection_year, fill, align=_CENTER)
        _cell(ws, rw, 2, r.total_av, fill, fmt=_DOLLAR)
        _cell(ws, rw, 3, r.residential_av, fill, fmt=_DOLLAR)
        _cell(ws, rw, 4, r.lot_av, fill, fmt=_DOLLAR)
        if og:
            _cell(ws, rw, og, r.centrally_assessed_av, fill, fmt=_DOLLAR)
        if comm:
            _cell(ws, rw, comm, r.commercial_av, fill, fmt=_DOLLAR)
        _cell(ws, rw, mr, r.mill_revenue, fill, fmt=_DOLLAR)
        _cell(ws, rw, nsr, r.net_senior_revenue, fill, fmt=_DOLLAR)
        _cell(ws, rw, nsub, r.net_sub_revenue, fill, fmt=_DOLLAR)
    ws.freeze_panes = "A6"


# ── Residential AV Build (aggregate value-creation schedule) ──────────────────
def _build_residential_av_sheet(ws, cfg, dev, sm):
    """
    Aggregate residential assessed-value build by year:
      Lots delivered → lot market value → lot taxable value @ rate;
      residential units → home market value (incl. reassessment) →
      cumulative home market value → home taxable value @ Taxable Ratio; → Total Residential Taxable Value.
    Values are shown in the year the value is created; the model collects it on
    the Utah lag — created in a calendar year, on the following 1 January roll,
    billed that 30 November, paying debt service the next 1 March (see Notes).
    """
    M = dev.cumulative_home_market_value
    closings = dev.home_closings
    lot_rate = cfg.lot_inventory_taxable_ratio
    resid_ratio = cfg.resid_taxable_ratio

    _title(ws, cfg, [cfg.pid_name, "Residential Taxable Value — Build",
                f"Lots @ {lot_rate:.0%} of market · Homes @ {resid_ratio:.3%} Residential Exemption · "
                f"reassessment {cfg.reassess_rate:.1%}"], 9)
    hdrs = [
        (1, "AV-Set\nYear", 10),
        (2, "Lots\nDelivered", 11),
        (3, "Lot Market\nValue", 15),
        (4, f"Lot Inventory Taxable Value\n@ {lot_rate:.0%}", 14),
        (5, "Total\nResidential Units", 13),
        (6, "Home Market Value\n(+ reassessment)", 17),
        (7, "Cumulative Home\nMarket Value", 17),
        (8, f"Home Taxable Value\n@ Taxable Ratio {resid_ratio:.3%}", 15),
        (9, "Total Residential\nTaxable Value", 16),
    ]
    for col, lbl, w in hdrs:
        _hdr(ws, 5, col, lbl, w)

    years = [s for s in range(cfg.first_year, cfg.senior_final_year + 1)
             if (M.get(s, 0) > 0 or dev.lot_market_value.get(s, 0) > 0
                 or dev.lot_deliveries.get(s, 0) > 0)]
    cum_units = 0
    tot_lots = tot_lotmv = 0.0
    i = 0
    for s in years:
        rw = 6 + i
        fill = _GRAY if i % 2 else _WHITE
        lots = dev.lot_deliveries.get(s, 0)
        lot_mv = dev.lot_market_value.get(s, 0.0)
        lot_av = lot_mv * cfg.lot_inventory_taxable_rate(s)
        cum_units += int(round(closings.get(s, 0)))
        home_added = M.get(s, 0.0) - M.get(s - 1, 0.0)
        cum_home_mv = M.get(s, 0.0)
        home_av = cum_home_mv * cfg.residential_assessment_rate(s)
        total_res_av = lot_av + home_av
        tot_lots += lots; tot_lotmv += lot_mv

        _cell(ws, rw, 1, s, fill, align=_CENTER)
        _cell(ws, rw, 2, int(round(lots)) or None, fill, fmt=_NUM)
        _cell(ws, rw, 3, round(lot_mv) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 4, round(lot_av) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 5, cum_units or None, fill, fmt=_NUM)
        _cell(ws, rw, 6, round(home_added) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 7, round(cum_home_mv) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 8, round(home_av) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 9, round(total_res_av) or None, fill, _BOLD, fmt=_DOLLAR)
        i += 1

    tot = 6 + len(years)
    _cell(ws, tot, 1, "Total", _TOTAL, _TOTAL_FONT, align=_CENTER)
    _cell(ws, tot, 2, int(round(tot_lots)) or None, _TOTAL, _TOTAL_FONT, fmt=_NUM)
    _cell(ws, tot, 3, round(tot_lotmv) or None, _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, tot, 5, cum_units or None, _TOTAL, _TOTAL_FONT, fmt=_NUM)
    _cell(ws, tot, 7, round(M.get(years[-1], 0.0)) if years else None,
          _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    for c in (4, 6, 8, 9):
        ws.cell(row=tot, column=c).fill = _TOTAL
    ws.freeze_panes = "A6"


# ── Comprehensive Summary tab (AV build by class → total AV → revenue) ────────
def _section_banner(ws, r, text, ncols):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    c = ws.cell(row=r, column=1, value=text)
    c.fill = PatternFill("solid", fgColor="2E6FA3")
    c.font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    c.alignment = _LEFT
    return r + 1


def _build_summary_light_sheet(ws, cfg, sm, dev, senior, refunding_bond=None):
    """
    Summary (Light): only the headline totals per year — ending market value,
    total taxable value, gross & net revenue, senior net debt service and
    coverage.  These are the '=' subtotals from the detailed Summary bridge.
    """
    vac = {r["collection"]: r for r in dev.lot_inventory_value_build(cfg)}
    res = {r["collection"]: r for r in dev.residential_value_build(cfg)}
    _title(ws, cfg, [cfg.pid_name, _county_line(cfg, " — Summary (Light)"),
                "Aggregate per year: market value (beginning & ending) → total taxable "
                "value → gross & net revenue for sizing."], 10)
    band = PatternFill("solid", fgColor="1F4E79")
    groups = [("", ["Assessment\nRoll Year", "Tax Collection\nYear"]),
              ("VACANT LAND — MARKET", ["Beginning", "Ending"]),
              ("RESIDENTIAL — MARKET", ["Beginning", "Ending"]),
              ("", ["State\nAssessed", "Total\nTaxable Value", "Gross\nRevenue",
                    "Net Revenue\n(senior sizing)"])]
    widths = [11, 12, 15, 15, 15, 15, 12, 16, 14, 15]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # Group banners on row 5, column headers on row 6, data from row 7 — clear of the
    # title block (rows 1-3) and its spacer (row 4) so nothing overlaps.
    c = 1
    for gname, subs in groups:
        if gname:
            ws.merge_cells(start_row=5, start_column=c, end_row=5, end_column=c + len(subs) - 1)
            gc = ws.cell(row=5, column=c, value=gname)
            gc.fill = band; gc.font = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
            gc.alignment = _CENTER
            for cc in range(c, c + len(subs)):
                ws.cell(row=5, column=cc).fill = band
        else:
            for cc in range(c, c + len(subs)):
                ws.merge_cells(start_row=5, start_column=cc, end_row=6, end_column=cc)
        for k, sub in enumerate(subs):
            _hdr(ws, 6 if gname else 5, c + k, sub, widths[c + k - 1])
        c += len(subs)
    for i, r in enumerate(sm.rows):
        y = r.collection_year
        rw = 7 + i
        fill = _GRAY if i % 2 else _WHITE
        vb, rb = vac.get(y, {}), res.get(y, {})
        vb_prev, rb_prev = vac.get(y - 1, {}), res.get(y - 1, {})
        # Certified/trued (historical) rows — blue-font their value cells, matching
        # Summary - Detail.  Beginning columns follow the PRIOR row's certified state
        # (a beginning value is the prior year's certified ending).  All dynamic —
        # sourced from the value builds' certified flag, no hardcoded years.
        v_cert, r_cert = bool(vb.get("certified")), bool(rb.get("certified"))
        v_cert_prev, r_cert_prev = bool(vb_prev.get("certified")), bool(rb_prev.get("certified"))
        cert_row = v_cert or r_cert
        vfont = lambda hit: _CERT_FONT if hit else _BODY
        for c in range(1, 11):
            ws.cell(row=rw, column=c).fill = fill
            ws.cell(row=rw, column=c).border = _BORDER
        _cell(ws, rw, 1, y - 1, fill, align=_CENTER)
        _cell(ws, rw, 2, y, fill, align=_CENTER)
        _cell(ws, rw, 3, round(vb_prev.get("cumulative", 0.0)) or None, fill, vfont(v_cert_prev), fmt=_DOLLAR)
        _cell(ws, rw, 4, round(vb.get("cumulative", 0.0)) or None, fill, vfont(v_cert), fmt=_DOLLAR)
        _cell(ws, rw, 5, round(rb.get("beginning", 0.0)) or None, fill, vfont(r_cert_prev), fmt=_DOLLAR)
        _cell(ws, rw, 6, round(rb.get("gross", 0.0)) or None, fill, vfont(r_cert), fmt=_DOLLAR)
        _cell(ws, rw, 7, round(r.state_av) or None, fill, vfont(cert_row), fmt=_DOLLAR)
        _cell(ws, rw, 8, round(r.total_av) or None, fill, _CERT_BOLD if cert_row else _BOLD, fmt=_DOLLAR)
        _cell(ws, rw, 9, round(r.mill_revenue + r.uniform_fee_revenue) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 10, round(r.net_senior_revenue) or None, fill, _BOLD, fmt=_DOLLAR)
    ws.freeze_panes = "C7"


def _build_summary_av_sheet(ws, cfg, dev, sm):
    """
    Summary — Detail: assessed-value components (vacant lot, home, centrally assessed /
    commercial when present, state assessed, exempt) → Total Taxable Value →
    revenue waterfall (mill levy + Uniform Fee − fees) → net revenue for senior sizing.

    Assessed values flow from the "Builder Lot Inventory Value" / "Residential Value" tabs
    (single source); the market-value build and the certified true-up / amortization
    live on those tabs, so this sheet stays a clean reconciliation to Total Taxable Value.
    """
    has_comm = sm.has_commercial
    has_og = sm.has_centrally_assessed
    _vl = sorted({cfg.lot_inventory_taxable_rate(r.collection_year) for r in sm.rows})
    _vl_lbl = (f"{_vl[0]:.1%}–{_vl[-1]:.1%}" if len(_vl) > 1 else (f"{_vl[0]:.1%}" if _vl else ""))
    _rr = sorted({cfg.residential_assessment_rate(r.collection_year) for r in sm.rows})
    _rr_lbl = (f"{_rr[0]:.3%}–{_rr[-1]:.3%}" if len(_rr) > 1 else (f"{_rr[0]:.3%}" if _rr else ""))

    cols = [("lotav", f"Lot Inventory\nTaxable Value\n@ {_vl_lbl}", 14),
            ("homeav", f"Home Taxable Value\n@ {_rr_lbl}", 14)]
    if has_og:
        cols.append(("ogav", "+ Centrally Assessed\nTaxable Value", 12))
    if has_comm:
        cols.append(("commav", "+ Commercial\nTaxable Value", 13))
    cols += [("stateav", "+ State\nAssessed", 12), ("exempt", "− Exempt", 11),
             ("total", "Total\nTaxable Value", 16),
             ("mill", f"Mill-Levy Rev\n@ {cfg.effective_ds_mill_levy:.3f}", 14),
             ("uniform_fee", f"Uniform\nFee\n@ {cfg.uniform_fee_prc:.0%}", 10),
             ("gross", "Gross\nRevenue", 13),
             ("treas", f"− County Collection\n@ {cfg.county_collection_fee:.2%}", 13),
             ("trust", "− Senior\nTrustee", 11),
             ("subtrust", "− Sub\nTrustee", 11),
             ("om", "− District\nAdmin", 11),
             ("omexp", "− District\nO&M", 11),
             ("net", "Net Revenue\n(senior sizing)", 15)]

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 10
    pos = {}
    c = 3
    for key, _lbl, w in cols:
        pos[key] = c
        ws.column_dimensions[get_column_letter(c)].width = w
        c += 1
    last_col = c - 1

    _title(ws, cfg, [cfg.pid_name, _county_line(cfg, " — Summary (Detail)"),
                "Taxable-value components → total taxable value → revenue waterfall → "
                "net revenue for sizing.  Lot-inventory and home taxable value flow from "
                "the Builder Lot Inventory Value / Residential Value tabs."], last_col)

    for r_, lbl in ((1, "Assessment\nRoll Year"), (2, "Tax Collection\nYear")):
        cell = ws.cell(row=5, column=r_, value=lbl)
        cell.fill = _LIGHT; cell.font = _HDR_FONT; cell.alignment = _CENTER; cell.border = _BORDER
        ws.merge_cells(start_row=5, start_column=r_, end_row=6, end_column=r_)
    for key, lbl, _w in cols:
        cell = ws.cell(row=5, column=pos[key], value=lbl)
        cell.fill = _LIGHT; cell.font = _HDR_FONT; cell.alignment = _CENTER; cell.border = _BORDER
        ws.merge_cells(start_row=5, start_column=pos[key], end_row=6, end_column=pos[key])

    # Certified/trued-up years, whose Lot Inventory\nTaxable Value / Home AV came straight from the
    # Inputs template (historical certified values + existing values) — highlighted.
    vac_cert = {r["collection"] for r in dev.lot_inventory_value_build(cfg) if r.get("certified")}
    res_cert = {r["collection"] for r in dev.residential_value_build(cfg) if r.get("certified")}
    state_cert = {R + 1 for R, h in (cfg.historical_av or {}).items()
                  if h.get("state_assessed") is not None}

    bold_keys = {"total", "gross", "net"}
    for i, r in enumerate(sm.rows):
        rw = 7 + i
        y = r.collection_year
        fill = _GRAY if i % 2 else _WHITE
        treas = -r.mill_revenue * cfg.county_collection_fee
        # District costs come from the same helper the revenue engine uses,
        # so the report and the sizing can never disagree.
        _admin, _trustee, _subtrustee = cfg.district_costs(r.collection_year)
        trust = -_trustee
        subtrust = -_subtrustee
        om = -_admin
        vals = {
            "lotav": r.lot_av, "homeav": r.residential_av, "ogav": r.centrally_assessed_av,
            "commav": r.commercial_av, "stateav": r.state_av, "exempt": -cfg.exempt_value,
            "total": r.total_av, "mill": r.mill_revenue, "uniform_fee": r.uniform_fee_revenue,
            "gross": r.mill_revenue + r.uniform_fee_revenue, "treas": treas,
            "trust": trust, "subtrust": subtrust, "om": om,
            "omexp": -r.om_expense, "net": r.net_senior_revenue,
        }
        _cell(ws, rw, 1, r.collection_year - 1, fill, align=_CENTER)
        _cell(ws, rw, 2, r.collection_year, fill, align=_CENTER)
        cert_row = y in vac_cert or y in res_cert
        for key, col in pos.items():
            # Highlight AND blue-font the taxable values / total that came from the
            # certified Inputs template, so inputs-driven AV stands out from projections.
            # State assessed is highlighted on any certified/trued row (so the whole
            # row reads consistently), as well as any year whose state value came
            # straight from the historical inputs table.
            hi = (key == "lotav" and y in vac_cert) or (key == "homeav" and y in res_cert) \
                or (key == "stateav" and (cert_row or y in state_cert)) \
                or (key == "total" and cert_row)
            if hi:
                font = _CERT_BOLD if key in bold_keys else _CERT_FONT
            else:
                font = _BOLD if key in bold_keys else _BODY
            cell_fill = _HILITE if hi else fill
            _cell(ws, rw, col, round(vals.get(key, 0.0)) or None, cell_fill, font, fmt=_DOLLAR)
    ws.freeze_panes = "C7"



# ── Senior-lien DS schedule (full Excel-style columns) ────────────────────────
def _build_ds_sheet(ws, cfg, tranche, title, sm=None, target_coverage=None):
    if target_coverage is None:
        target_coverage = tranche.coverage
    _title(ws, cfg, [cfg.pid_name, title,
                f"Par ${tranche.par_amount:,.0f}  ·  {tranche.rate:.2%}  ·  "
                f"{target_coverage:.2f}x coverage"], 18)
    hdrs = [
        (1, "Payment\nDate", 12), (2, "Rate", 8), (3, "Yield", 8), (4, "Price", 8),
        (5, "Premium/\nOID", 11), (6, "Principal", 13), (7, "Interest", 13),
        (8, "Total", 13), (9, "Annual Gross\nTotal DS", 14),
        (10, "Capitalized\nInterest", 13), (11, "DSRF Surplus\nRelease", 13),
        (12, "Interest\nEarnings", 12), (13, "Net\nTotal", 13),
        (14, "Annual Net\nTotal DS", 14), (15, "Bond\nValue", 14),
        (16, "Senior Lien\nRevenues", 14), (17, "Actual\nCoverage", 11),
        (18, "Target\nCoverage", 11),
    ]
    for col, lbl, w in hdrs:
        _hdr(ws, 5, col, lbl, w)

    # Annual roll-ups (booked on the principal-maturity / December row).
    annual_gross = tranche.annual_gross_ds()
    annual_net = tranche.annual_net_ds()

    prin_month = tranche.prin_month
    cum_prin = 0.0
    _PCT = '0.000%'
    for i, p in enumerate(tranche.schedule):
        rw = 6 + i
        fill = _GRAY if i % 2 else _WHITE
        yr = p.payment_date.year
        is_prin_row = p.payment_date.month == prin_month
        cum_prin += p.principal
        bond_value = max(tranche.par_amount - cum_prin, 0.0)

        # Shade every cell in the row (so blank rate/yield/price/coverage cells on
        # the in-between coupon rows are not left unshaded).
        for c in range(1, 19):
            cc = ws.cell(row=rw, column=c)
            cc.fill = fill
            cc.border = _BORDER

        _cell(ws, rw, 1, p.payment_date, fill, fmt="MM/DD/YYYY", align=_LEFT)
        # Rate / Yield / Price / Premium-OID booked on coupon (December) rows.
        if is_prin_row:
            _cell(ws, rw, 2, tranche.coupon_for(yr), fill, fmt=_PCT)
            _cell(ws, rw, 3, tranche.yield_for(yr), fill, fmt=_PCT)
            _cell(ws, rw, 4, tranche.price_for(yr), fill, fmt='0.000')
            if p.principal:
                _cell(ws, rw, 5, round(tranche.premium_for(yr, p.principal), 0) or None,
                      fill, fmt=_DOLLAR)
        _cell(ws, rw, 6, p.principal or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 7, p.interest or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 8, p.gross_total or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 10, p.capitalized_interest or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 11, p.surplus_release or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 12, p.dsrf_earnings or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 13, round(p.net_total, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 15, round(bond_value, 0) or None, fill, fmt=_DOLLAR)

        # Annual figures + coverage on the December (principal-maturity) row.
        if is_prin_row:
            ag = annual_gross.get(yr, 0.0)
            an = annual_net.get(yr, 0.0)
            _cell(ws, rw, 9, round(ag, 0) or None, fill, fmt=_DOLLAR)
            _cell(ws, rw, 14, round(an, 0) or None, fill, fmt=_DOLLAR)
            if sm is not None:
                rev = sm.net_senior_revenue(yr)
                _cell(ws, rw, 16, round(rev, 0) or None, fill, fmt=_DOLLAR)
                if an > 0:
                    _cell(ws, rw, 17, rev / an, fill, fmt='0.00"x"')
                _cell(ws, rw, 18, target_coverage, fill, fmt='0.00"x"')

    # Totals row
    tot = 6 + len(tranche.schedule)
    _cell(ws, tot, 1, "Total", _TOTAL, _TOTAL_FONT, align=_CENTER)
    if abs(tranche.total_premium) > 0.5:
        _cell(ws, tot, 5, round(tranche.total_premium, 0), _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, tot, 6, tranche.par_amount, _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, tot, 8, sum(p.gross_total for p in tranche.schedule),
          _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, tot, 9, sum(annual_gross.values()), _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, tot, 14, tranche.total_net_ds, _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    for c in range(2, 19):
        if c not in (6, 8, 9, 14):
            ws.cell(row=tot, column=c).fill = _TOTAL
    ws.freeze_panes = "B6"


# ── First-financing Sources & Uses (official-statement format) ───────────────
def _build_su_first_sheet(ws, cfg, senior, sub_result, surplus_fund=None, dev=None):
    """
    Sources & Uses laid out in the standard official-statement format: a
    Senior-lien column, a Subordinate cash-flow column, and a Total, with the
    Sources / Uses blocks, Key Assumptions, Bond Statistics, and Taxing-Authority
    assumption sections beneath.
    """
    LBL, SR, SB, TOT = 1, 2, 3, 4
    widths = [(LBL, 44), (SR, 19), (SB, 21), (TOT, 18)]
    for c, w in widths:
        ws.column_dimensions[get_column_letter(c)].width = w

    sub_par = sub_result.par_amount if sub_result else (cfg.sub_par or 0.0)
    senior_par = senior.par_amount
    yr = cfg.delivery.year

    capi = sum(p.capitalized_interest for p in senior.schedule)
    dsrf = senior.dsrf_deposit
    prem_sr = senior.total_premium
    uwd_sr = cfg.uwd_senior * senior_par
    uwd_sb = cfg.uwd_sub * sub_par
    coi = cfg.coi
    reimb_sr = senior_par + prem_sr - dsrf - capi - uwd_sr - coi
    reimb_sb = sub_par - uwd_sb
    src_sr, src_sb = senior_par + prem_sr, sub_par

    # ── Title block ──────────────────────────────────────────────────────────
    def _band(r, text, font=_TITLE_FONT, fill=_BLUE):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOT)
        c = ws.cell(row=r, column=1, value=text); c.fill = fill; c.font = font
        c.alignment = _CENTER
        for cc in range(1, TOT + 1):
            ws.cell(row=r, column=cc).fill = fill
    _band(1, cfg.pid_name or "Metropolitan District")
    loc = " / ".join([x for x in [cfg.city, _county_line(cfg)] if x])
    _band(2, loc, font=_WHITE_FONT)
    _band(3, f"Development Projections at {cfg.effective_ds_mill_levy:.3f} Mills for Debt Service",
          font=_WHITE_FONT)
    rating = "Investment Grade" if cfg.ig_rated == "Yes" else "Non-Rated"
    lots = dev.total_lots if dev is not None else ""
    _band(4, f"Scenario: {rating} / {cfg.reassess_rate:.2%} Reassessment"
             + (f" / {lots} Lots" if lots else ""), font=_WHITE_FONT)
    ws.cell(row=6, column=1, value=f"--- Series {yr} Financing ---").font = _BOLD

    # ── Helpers ──────────────────────────────────────────────────────────────
    def section(r, text):
        ws.cell(row=r, column=1, value=text).font = _HDR_FONT
        for cc in range(1, TOT + 1):
            ws.cell(row=r, column=cc).fill = _LIGHT
        ws.cell(row=r, column=1).font = _HDR_FONT

    def colhdr(r):
        for c, lbl in [(SR, f"Senior Lien Bonds\nSeries {yr}A"),
                       (SB, f"Subordinate Lien\nCashflow Bonds\nSeries {yr}B"),
                       (TOT, "Total")]:
            cell = ws.cell(row=r, column=c, value=lbl)
            cell.fill = _LIGHT; cell.font = _HDR_FONT; cell.border = _BORDER; cell.alignment = _CENTER
        ws.row_dimensions[r].height = 42

    def line(r, label, sv, bv, tv=None, *, fmt=_DOLLAR, total=False, reimb=False, indent=True):
        fill = _TOTAL if total else (_GREEN if reimb else _WHITE)
        font = _TOTAL_FONT if total else (_BOLD if reimb else _BODY)
        _cell(ws, r, LBL, ("  " if indent else "") + label, fill, font, align=_LEFT)
        for c, v in [(SR, sv), (SB, bv), (TOT, tv if tv is not None else
                     ((sv or 0) + (bv or 0)))]:
            if v is None:                       # not applicable — leave cleanly blank
                ws.cell(row=r, column=c).fill = fill
            else:                               # show the value, including real zeros
                disp = round(v) if fmt == _DOLLAR else v
                _cell(ws, r, c, disp, fill, font, fmt=fmt)

    # ── Sources and Uses ─────────────────────────────────────────────────────
    section(8, "Sources and Uses"); r = 10
    colhdr(r); r += 1
    line(r, "Par Amount of Bonds", senior_par, sub_par); r += 1
    if abs(prem_sr) > 0.5:
        line(r, "Plus: Premium / (Discount)", prem_sr, 0.0); r += 1
    if surplus_fund is not None:
        line(r, "Existing Senior Surplus Fund Balance", 0.0, 0.0); r += 1
    line(r, "TOTAL SOURCES OF FUNDS:", src_sr, src_sb, total=True, indent=False); r += 2

    section(r, "Uses of Funds"); r += 2
    colhdr(r); r += 1
    line(r, "Estimated Reimbursement Amount", reimb_sr, reimb_sb, reimb=True); r += 1
    line(r, "Debt Service Reserve Fund", dsrf, 0.0); r += 1
    line(r, "Capitalized Interest", capi, 0.0); r += 1
    line(r, "Underwriters' Discount", uwd_sr, uwd_sb); r += 1
    line(r, "Costs of Issuance", coi, None, tv=coi); r += 1
    line(r, "TOTAL USES OF FUNDS:",
         reimb_sr + dsrf + capi + uwd_sr + coi, reimb_sb + uwd_sb,
         total=True, indent=False); r += 2

    # ── Bond statistics (computed) ───────────────────────────────────────────
    sr_prin = [(p.payment_date, p.principal) for p in senior.schedule if p.principal]
    sr_gross = [(p.payment_date, p.gross_total) for p in senior.schedule]
    sr_first_int = min((p.payment_date for p in senior.schedule), default=None)
    sr_first_mat = min((p.payment_date for p in senior.schedule if p.principal), default=None)
    sr_final = max((p.payment_date for p in senior.schedule), default=None)
    sr_avg_life = _avg_life(sr_prin, cfg.delivery, senior_par)
    sr_total_ds = sum(p.gross_total for p in senior.schedule)
    sr_max_ds = max(senior.annual_gross_ds().values()) if senior.schedule else 0.0
    sr_arb_tic = _tic(sr_gross, senior_par + prem_sr - uwd_sr, cfg.delivery)
    sr_allin_tic = _tic(sr_gross, senior_par + prem_sr - uwd_sr - coi, cfg.delivery)

    sb_avg_life = sb_total_ds = sb_max_ds = 0.0
    sb_arb_tic = None
    sb_first_mat = sb_final = None
    if sub_result and sub_result.rows:
        pm, pd = cfg.prin_maturity, cfg.prin_maturity_day_sub
        sb_prin = [(_date(rr["year"], pm, pd), rr["principal_paid"])
                   for rr in sub_result.rows if rr["principal_paid"]]
        sb_gross = [(_date(rr["year"], pm, pd), rr["total_paid"])
                    for rr in sub_result.rows if rr["total_paid"]]
        sb_avg_life = _avg_life(sb_prin, cfg.delivery, sub_par)
        sb_total_ds = sub_result.total_payments
        sb_max_ds = max((rr["total_paid"] for rr in sub_result.rows), default=0.0)
        sb_arb_tic = _tic(sb_gross, sub_par - uwd_sb, cfg.delivery)
        sb_first_mat = min((d for d, _ in sb_prin), default=None)
        sb_final = max((d for d, _ in sb_gross), default=None)

    # ── Key assumptions ──────────────────────────────────────────────────────
    section(r, "Key Assumptions:"); r += 2
    DATEFMT = 'mm/dd/yyyy'
    def drow(r, label, sv, bv=None, tv=None, fmt=_DOLLAR):
        _cell(ws, r, LBL, label, _WHITE, _BODY, align=_LEFT)
        # A district-wide value (only the Total given) is shown once, beside the
        # label, rather than leaving the per-lien columns as empty boxes.
        if sv is None and bv is None and tv is not None:
            _cell(ws, r, SR, tv, _WHITE, _BODY, fmt=fmt,
                  align=(_CENTER if fmt == DATEFMT else _RIGHT))
            return
        for c, v in [(SR, sv), (SB, bv), (TOT, tv)]:
            if v is None:                       # not applicable — leave cleanly blank
                continue
            _cell(ws, r, c, v, _WHITE, _BODY, fmt=fmt,
                  align=(_CENTER if fmt == DATEFMT else _RIGHT))
    drow(r, "Delivery Date", cfg.delivery, cfg.delivery, cfg.delivery, fmt=DATEFMT); r += 1
    drow(r, "First Interest Date", sr_first_int, fmt=DATEFMT); r += 1
    drow(r, "First Maturity Date", sr_first_mat, sb_first_mat, fmt=DATEFMT); r += 1
    drow(r, "Final Maturity Date", sr_final, sb_final, fmt=DATEFMT); r += 1
    drow(r, "First Par Call Date", cfg.par_call_date, fmt=DATEFMT); r += 1
    drow(r, f"Capitalized Interest Period ({cfg.capi_term}mos)", cfg.capi_end_date, fmt=DATEFMT); r += 1
    drow(r, "Debt Service Coverage", cfg.dsc_senior, cfg.dsc_sub, fmt='0.00"x"'); r += 1
    drow(r, "Reassessment", None, None, cfg.reassess_rate, fmt='0.00%'); r += 1
    if surplus_fund is not None:
        drow(r, "Senior Lien Bonds Surplus Fund Target", surplus_fund.target, fmt=_DOLLAR); r += 1
    r += 1

    # ── Bond statistics ──────────────────────────────────────────────────────
    section(r, "Bond Statistics:"); r += 2
    drow(r, "Average Life (years)", sr_avg_life, sb_avg_life or None, fmt='0.00'); r += 1
    drow(r, "Arbitrage TIC", sr_arb_tic, sb_arb_tic, fmt='0.000%'); r += 1
    drow(r, "All-in TIC", sr_allin_tic, fmt='0.000%'); r += 1
    drow(r, "Maximum Annual Debt Service", sr_max_ds, sb_max_ds or None, fmt=_DOLLAR); r += 1
    drow(r, "Total Debt Service", sr_total_ds, sb_total_ds or None, fmt=_DOLLAR); r += 1
    r += 1

    # ── Taxing-authority & fee assumptions ───────────────────────────────────
    section(r, "Taxing Authority and Fee Assumptions:"); r += 2
    ws.cell(row=r, column=LBL, value="Residential Taxable Ratio").font = _BOLD; r += 1
    drow(r, "  Governing Document Adjusted Base Rate", None, None, cfg.resid_taxable_ratio_prior, fmt='0.000%'); r += 1
    drow(r, "  Financing Plan Assumption", None, None, cfg.resid_taxable_ratio, fmt='0.000%'); r += 1
    ws.cell(row=r, column=LBL, value="Debt Service Mills").font = _BOLD; r += 1
    drow(r, "  Governing Document Mill Levy Cap", None, None, cfg.mill_levy_governing_doc, fmt='0.000'); r += 1
    drow(r, "  Targeted (Effective) Mill Levy", None, None, cfg.effective_ds_mill_levy, fmt='0.000'); r += 1
    ws.cell(row=r, column=LBL, value="Other Revenue and Fee Assumptions").font = _BOLD; r += 1
    drow(r, "  Personal Property Uniform Fee", None, None, cfg.uniform_fee_prc, fmt='0.00%'); r += 1
    drow(r, "  Mill-Levy Collection Rate", None, None, cfg.tax_collect_mill_prc, fmt='0.00%'); r += 1
    drow(r, "  County Collection Cost", None, None, cfg.county_collection_fee, fmt='0.000%'); r += 1
    drow(r, "  Annual Trustee Fee", None, None, cfg.trustee_fee, fmt=_DOLLAR); r += 1


# ── Refunding Sources & Uses (official-statement format) ─────────────────────
def _build_su_refunding_sheet(ws, cfg, refunding_result, dev=None):
    """
    Refunding Sources & Uses in the same official-statement format as the
    new-money sheet: a Refunding Senior column and a Total, with Sources / Uses,
    Key Assumptions, Bond Statistics, and Taxing-Authority sections.
    """
    LBL, AMT, TOT = 1, 2, 3
    for c, w in [(LBL, 44), (AMT, 24), (TOT, 18)]:
        ws.column_dimensions[get_column_letter(c)].width = w

    rb = refunding_result.refunding_bond
    su = refunding_result.sources_uses
    yr = rb.delivery.year
    par = rb.par_amount
    prem = rb.total_premium
    sub_par = su.sources.get("Refunding Subordinate Par", 0.0)
    uwd_senior = cfg.uwd_senior_refunding * par
    uwd = su.uses.get("Underwriter's Discount", uwd_senior)   # senior + subordinate
    coi = cfg.coi_refunding
    released_dsrf = su.sources.get("Released DSRF (refunded series)", 0.0)
    surplus_oh = su.sources.get("Surplus Funds on Hand", 0.0)
    esc_sr = su.uses.get("Refunding Escrow (old senior)", 0.0)
    esc_sb = su.uses.get("Refunding Escrow (subordinate)", 0.0)
    insurance = su.uses.get("Bond Insurance / Surety", 0.0)
    new_money = refunding_result.new_money_reimbursement

    def _band(r, text, font=_TITLE_FONT, fill=_BLUE):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=TOT)
        c = ws.cell(row=r, column=1, value=text); c.font = font; c.alignment = _CENTER
        for cc in range(1, TOT + 1):
            ws.cell(row=r, column=cc).fill = fill
    _band(1, cfg.pid_name or "Metropolitan District")
    loc = " / ".join([x for x in [cfg.city, _county_line(cfg)] if x])
    _band(2, loc, font=_WHITE_FONT)
    _band(3, f"Development Projections at {cfg.effective_ds_mill_levy:.3f} Mills for Debt Service",
          font=_WHITE_FONT)
    ws.cell(row=5, column=1, value=f"--- Series {yr} Refunding Financing ---").font = _BOLD

    def section(r, text):
        ws.cell(row=r, column=1, value=text).font = _HDR_FONT
        for cc in range(1, TOT + 1):
            ws.cell(row=r, column=cc).fill = _LIGHT
        ws.cell(row=r, column=1).font = _HDR_FONT

    def colhdr(r):
        for c, lbl in [(AMT, f"Refunding Senior Lien\nBonds — Series {yr}"), (TOT, "Total")]:
            cell = ws.cell(row=r, column=c, value=lbl)
            cell.fill = _LIGHT; cell.font = _HDR_FONT; cell.border = _BORDER; cell.alignment = _CENTER
        ws.row_dimensions[r].height = 42

    def line(r, label, v, *, fmt=_DOLLAR, total=False, reimb=False, indent=True):
        fill = _TOTAL if total else (_GREEN if reimb else _WHITE)
        font = _TOTAL_FONT if total else (_BOLD if reimb else _BODY)
        _cell(ws, r, LBL, ("  " if indent else "") + label, fill, font, align=_LEFT)
        if v is None:
            for c in (AMT, TOT):
                ws.cell(row=r, column=c).fill = fill
            return
        for c in (AMT, TOT):
            disp = round(v) if fmt == _DOLLAR else v
            _cell(ws, r, c, disp, fill, font, fmt=fmt)

    def drow(r, label, v, fmt=_DOLLAR):
        _cell(ws, r, LBL, label, _WHITE, _BODY, align=_LEFT)
        if v is None:                           # not applicable — leave cleanly blank
            return
        _cell(ws, r, AMT, v, _WHITE, _BODY, fmt=fmt,
              align=(_CENTER if fmt == 'mm/dd/yyyy' else _RIGHT))

    # Sources and Uses
    section(7, "Sources and Uses"); r = 9
    colhdr(r); r += 1
    line(r, "Par Amount of Bonds (Senior)", par); r += 1
    if sub_par:
        line(r, "Par Amount of Bonds (Subordinate)", sub_par); r += 1
    if abs(prem) > 0.5:
        line(r, "Plus: Premium / (Discount)", prem); r += 1
    line(r, "Released DSRF (Refunded Series)", released_dsrf); r += 1
    line(r, "Surplus Funds on Hand", surplus_oh); r += 1
    line(r, "TOTAL SOURCES OF FUNDS:", su.total_sources, total=True, indent=False); r += 2

    section(r, "Uses of Funds"); r += 2
    colhdr(r); r += 1
    line(r, "Estimated Reimbursement Amount (New Money)", new_money, reimb=True); r += 1
    line(r, "Refunding Escrow — Senior (defease old bonds)", esc_sr); r += 1
    if esc_sb:
        line(r, "Refunding Escrow — Subordinate", esc_sb); r += 1
    if insurance:
        line(r, "Bond Insurance / Surety", insurance); r += 1
    line(r, "Underwriters' Discount", uwd); r += 1
    line(r, "Costs of Issuance", coi); r += 1
    line(r, "TOTAL USES OF FUNDS:", su.total_uses, total=True, indent=False); r += 2

    # Statistics
    prin = [(p.payment_date, p.principal) for p in rb.schedule if p.principal]
    gross = [(p.payment_date, p.gross_total) for p in rb.schedule]
    first_int = min((p.payment_date for p in rb.schedule), default=None)
    first_mat = min((p.payment_date for p in rb.schedule if p.principal), default=None)
    final_mat = max((p.payment_date for p in rb.schedule), default=None)
    par_call = rb.call_provisions.par_call_date if rb.call_provisions else None

    section(r, "Key Assumptions:"); r += 2
    DATEFMT = 'mm/dd/yyyy'
    drow(r, "Delivery Date", rb.delivery, fmt=DATEFMT); r += 1
    drow(r, "First Interest Date", first_int, fmt=DATEFMT); r += 1
    drow(r, "First Maturity Date", first_mat, fmt=DATEFMT); r += 1
    drow(r, "Final Maturity Date", final_mat, fmt=DATEFMT); r += 1
    drow(r, "First Par Call Date", par_call, fmt=DATEFMT); r += 1
    drow(r, "Debt Service Coverage", cfg.dsc_refunding, fmt='0.00"x"'); r += 1
    drow(r, "Reassessment", cfg.reassess_rate, fmt='0.00%'); r += 1
    drow(r, "Callable Principal Refunded", refunding_result.refunded_par_outstanding); r += 1
    drow(r, f"Call Price Applied", refunding_result.call_price / 100.0, fmt='0.00%'); r += 1
    r += 1

    section(r, "Bond Statistics:"); r += 2
    drow(r, "Average Life (years)", _avg_life(prin, rb.delivery, par), fmt='0.00'); r += 1
    drow(r, "Arbitrage TIC", _tic(gross, par + prem - uwd_senior, rb.delivery), fmt='0.000%'); r += 1
    drow(r, "All-in TIC", _tic(gross, par + prem - uwd_senior - coi, rb.delivery), fmt='0.000%'); r += 1
    drow(r, "Maximum Annual Debt Service",
         max(rb.annual_gross_ds().values()) if rb.schedule else 0.0); r += 1
    drow(r, "Total Debt Service", sum(p.gross_total for p in rb.schedule)); r += 1
    r += 1

    section(r, "Taxing Authority and Fee Assumptions:"); r += 2
    ws.cell(row=r, column=LBL, value="Residential Taxable Ratio").font = _BOLD; r += 1
    drow(r, "  Governing Document Adjusted Base Rate", cfg.resid_taxable_ratio_prior, fmt='0.000%'); r += 1
    drow(r, "  Financing Plan Assumption", cfg.resid_taxable_ratio, fmt='0.000%'); r += 1
    ws.cell(row=r, column=LBL, value="Debt Service Mills").font = _BOLD; r += 1
    drow(r, "  Governing Document Mill Levy Cap", cfg.mill_levy_governing_doc, fmt='0.000'); r += 1
    drow(r, "  Targeted (Effective) Mill Levy", cfg.effective_ds_mill_levy, fmt='0.000'); r += 1
    ws.cell(row=r, column=LBL, value="Other Revenue and Fee Assumptions").font = _BOLD; r += 1
    drow(r, "  Personal Property Uniform Fee", cfg.uniform_fee_prc, fmt='0.00%'); r += 1
    drow(r, "  Mill-Levy Collection Rate", cfg.tax_collect_mill_prc, fmt='0.00%'); r += 1
    drow(r, "  County Collection Cost", cfg.county_collection_fee, fmt='0.000%'); r += 1
    drow(r, "  Annual Trustee Fee", cfg.trustee_fee, fmt=_DOLLAR); r += 1


# ── Capitalized Interest (CAPI) Fund ─────────────────────────────────────────
def _build_capi_fund_sheet(ws, cfg, senior):
    """
    Capitalized-interest fund draw-down: the CAPI deposit (a use of bond
    proceeds) pays senior interest during the capitalized-interest period,
    drawn down on each interest payment date until exhausted.
    """
    capi_rows = [p for p in senior.schedule if p.capitalized_interest > 0.005]
    deposit = sum(p.capitalized_interest for p in senior.schedule)
    _title(ws, cfg, [cfg.pid_name, "Capitalized Interest (CAPI) Fund — Senior Lien",
                f"${deposit:,.0f} funds ~{cfg.capi_term} months of interest through "
                f"{cfg.capi_end_date:%b %Y}; balance earns {cfg.interest_earn_rate:.2%}/yr"], 6)
    for c, lbl, w in [(1, "Date", 13), (2, "Beginning\nBalance", 16),
                      (3, "Periodic\nRate", 11), (4, "Interest\nEarned", 14),
                      (5, "Interest\nFunded (Draw)", 16), (6, "Ending\nBalance", 16)]:
        _hdr(ws, 5, c, lbl, w)

    # Interest draws happen on the senior interest-payment dates; spread the fund
    # over the whole capitalized-interest period one month at a time, earning the
    # periodic (monthly) interest rate on the running balance.
    draws = {(p.payment_date.year, p.payment_date.month): p.capitalized_interest
             for p in capi_rows}
    monthly_rate = cfg.interest_earn_rate / 12.0

    # Opening deposit row (delivery date).
    _cell(ws, 6, 1, senior.delivery, _WHITE, fmt="MM/DD/YYYY", align=_LEFT)
    _cell(ws, 6, 2, round(deposit), _WHITE, fmt=_DOLLAR)
    for c in (3, 4, 5):
        ws.cell(row=6, column=c).fill = _WHITE
    _cell(ws, 6, 6, round(deposit), _WHITE, fmt=_DOLLAR)

    bal = float(deposit)
    tot_earned = 0.0
    cur = senior.delivery
    rw, i = 7, 0
    while cur < cfg.capi_end_date:
        ny = cur.year + (1 if cur.month == 12 else 0)
        nm = 1 if cur.month == 12 else cur.month + 1
        cur = _date(ny, nm, cfg.prin_maturity_day_senior)
        fill = _GRAY if i % 2 else _WHITE
        beg = bal
        earned = beg * monthly_rate
        draw = draws.get((cur.year, cur.month), 0.0)
        bal = max(beg + earned - draw, 0.0)
        tot_earned += earned
        _cell(ws, rw, 1, cur, fill, fmt="MM/DD/YYYY", align=_LEFT)
        _cell(ws, rw, 2, round(beg), fill, fmt=_DOLLAR)
        _cell(ws, rw, 3, monthly_rate, fill, fmt='0.0000%')
        _cell(ws, rw, 4, round(earned) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 5, round(draw) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 6, round(bal), fill, fmt=_DOLLAR)
        rw += 1
        i += 1

    _cell(ws, rw, 1, "Total", _TOTAL, _TOTAL_FONT, align=_LEFT)
    for c in (2, 3, 6):
        ws.cell(row=rw, column=c).fill = _TOTAL
    _cell(ws, rw, 4, round(tot_earned), _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, rw, 5, round(deposit), _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    ws.freeze_panes = "A6"


# ── Sources & Uses sheet ──────────────────────────────────────────────────────
def _build_su_sheet(ws, cfg, su, title, subtitle=""):
    _title(ws, cfg, [cfg.pid_name, title, subtitle], 3)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 4
    ws.column_dimensions["C"].width = 18

    row = 5
    ws.cell(row=row, column=1, value="SOURCES OF FUNDS").font = _HDR_FONT
    ws.cell(row=row, column=1).fill = _LIGHT
    ws.cell(row=row, column=3).fill = _LIGHT
    row += 1
    for k, v in su.sources.items():
        _cell(ws, row, 1, k, _WHITE, align=_LEFT)
        _cell(ws, row, 3, v, _WHITE, fmt=_DOLLAR)
        row += 1
    _cell(ws, row, 1, "Total Sources", _TOTAL, _TOTAL_FONT, align=_LEFT)
    _cell(ws, row, 3, su.total_sources, _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    row += 2

    ws.cell(row=row, column=1, value="USES OF FUNDS").font = _HDR_FONT
    ws.cell(row=row, column=1).fill = _LIGHT
    ws.cell(row=row, column=3).fill = _LIGHT
    row += 1
    for k, v in su.uses.items():
        is_reimb = k == "Reimbursement"
        fill = _GREEN if is_reimb else _WHITE
        _cell(ws, row, 1, k, fill, _BOLD if is_reimb else _BODY, align=_LEFT)
        _cell(ws, row, 3, v, fill, _BOLD if is_reimb else _BODY, fmt=_DOLLAR)
        row += 1
    _cell(ws, row, 1, "Total Uses", _TOTAL, _TOTAL_FONT, align=_LEFT)
    _cell(ws, row, 3, su.total_uses, _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)


# ── Subordinate cash-flow sheet ───────────────────────────────────────────────
def _build_sub_sheet(ws, cfg, sub_result, series_label=""):
    _title(ws, cfg, [cfg.pid_name,
                f"Subordinate Lien Cash-Flow Bonds{series_label}  ·  "
                f"{sub_result.coverage:.2f}x coverage",
                f"Par ${sub_result.par_amount:,.0f}  ·  {sub_result.rate:.2%}  "
                f"(unpaid interest accretes)"], 15)
    # Two-row header: section banners (row 5) over grouped columns (row 6).
    NC = 19
    band = PatternFill("solid", fgColor="1F4E79")
    sub_groups = [
        ("", [(1, "Payment\nDate", 12), (2, "Rate", 8), (3, "Yield", 8), (4, "Price", 8)]),
        ("HOW “AVAILABLE TO SUB” IS DERIVED  (revenue the sub lien actually receives)",
         [(5, "Senior Residual\n(net rev − sr DS)", 14), (6, "Surplus\nReserve Bal.", 13),
          (7, "+ Excess to Sub\n(above target)", 13), (8, "+ Sub Incremental\n(sub−sr fees)", 13),
          (9, "+ Reserve\nRelease", 12), (10, "Available\nto Sub", 14)]),
        ("SUBORDINATE DEBT SERVICE",
         [(11, "After\nCoverage", 12), (12, "Interest\nPaid", 12), (13, "Accrued Int.\nBalance", 13),
          (14, "Principal\nPaid", 12), (15, "Principal\nBalance", 13), (16, "Total\nPaid", 12)]),
        ("COVERAGE",
         [(17, "Actual\nCoverage", 10), (18, "Target\nCoverage", 10), (19, "Unused\nRevenues", 13)]),
    ]
    for gname, cols in sub_groups:
        if gname:
            first, last = cols[0][0], cols[-1][0]
            ws.merge_cells(start_row=5, start_column=first, end_row=5, end_column=last)
            gc = ws.cell(row=5, column=first, value=gname)
            gc.fill = band; gc.font = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
            gc.alignment = _CENTER
            for cc in range(first, last + 1):
                ws.cell(row=5, column=cc).fill = band
            for col, lbl, w in cols:          # sub-headers on row 6
                _hdr(ws, 6, col, lbl, w)
        else:                                  # ungrouped: header spans rows 5-6
            for col, lbl, w in cols:
                ws.merge_cells(start_row=5, start_column=col, end_row=6, end_column=col)
                _hdr(ws, 5, col, lbl, w)

    _PCT = '0.000%'
    bold = {5, 10, 16}   # senior residual, available to sub, total paid
    # Run the same time period as the senior lien bonds (first sub year through the
    # senior final maturity), filling years after payoff so the schedule spans the
    # full ~30-year horizon rather than stopping when the note is repaid.
    by_year = {r["year"]: r for r in sub_result.rows}
    years = list(range(sub_result.first_year, sub_result.final_year + 1))
    for i, y in enumerate(years):
        rw = 7 + i
        fill = _GRAY if i % 2 else _WHITE
        for c in range(1, NC + 1):
            cc = ws.cell(row=rw, column=c); cc.fill = fill; cc.border = _BORDER
        _cell(ws, rw, 1, _date(y, cfg.prin_maturity, cfg.prin_maturity_day_sub),
              fill, fmt="MM/DD/YYYY", align=_LEFT)
        if i == 0:                          # single accreting rate, sold at par
            _cell(ws, rw, 2, sub_result.rate, fill, fmt=_PCT)
            _cell(ws, rw, 3, sub_result.rate, fill, fmt=_PCT)
            _cell(ws, rw, 4, 100.0, fill, fmt='0.000')
        r = by_year.get(y)
        if r:
            avail = r["available_to_sub"]; paid = r["total_paid"]
            cells = {
                5: r["senior_residual"], 6: r["surplus_balance"],
                7: r["excess_to_sub"], 8: r["sub_incremental"], 9: r["reserve_release"],
                10: avail, 11: r["coverage_adj_available"], 12: r["interest_paid"],
                13: r["accrued_balance"], 14: r["principal_paid"], 16: paid,
                19: (avail - paid),
            }
            for c, v in cells.items():
                _cell(ws, rw, c, round(v) or None, fill,
                      _BOLD if c in bold else _BODY, fmt=_DOLLAR)
            _cell(ws, rw, 15, r["principal_balance"], fill, fmt=_DOLLAR)  # show 0 too
            if paid > 0:
                _cell(ws, rw, 17, avail / paid, fill, fmt='0.00"x"')
            _cell(ws, rw, 18, cfg.dsc_sub, fill, fmt='0.00"x"')
        else:
            _cell(ws, rw, 15, 0, fill, fmt=_DOLLAR)   # note fully repaid

    tot = 7 + len(years)
    for c in range(1, NC + 1):
        ws.cell(row=tot, column=c).fill = _TOTAL
    tot_avail = sum(r["available_to_sub"] for r in sub_result.rows)
    _cell(ws, tot, 1, "Total", _TOTAL, _TOTAL_FONT, align=_CENTER)
    _cell(ws, tot, 10, round(tot_avail), _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, tot, 12, sub_result.total_interest_paid, _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, tot, 14, sub_result.total_principal_paid, _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, tot, 16, sub_result.total_payments, _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, tot, 19, round(tot_avail - sub_result.total_payments), _TOTAL, _TOTAL_FONT, fmt=_DOLLAR)
    ws.freeze_panes = "A7"


# ── Senior surplus fund sheet ─────────────────────────────────────────────────
def _build_surplus_sheet(ws, cfg, surplus_fund):
    _title(ws, cfg, [cfg.pid_name, "Senior Surplus / Debt-Service-Reserve Fund",
                f"Target ${surplus_fund.target:,.0f}  ·  earns {cfg.interest_earn_rate:.2%} "
                f"on balance  (excess flows to subordinate lien)"], 7)
    hdrs = [(1, "Year", 8), (2, "Senior\nResidual", 15),
            (3, f"Interest\nEarned @ {cfg.interest_earn_rate:.2%}", 14),
            (4, "Deposit to\nReserve", 14), (5, "Reserve\nBalance", 14),
            (6, "Excess to\nSub Lien", 14), (7, "Reserve\nRelease", 14)]
    for col, lbl, w in hdrs:
        _hdr(ws, 5, col, lbl, w)
    for i, r in enumerate(surplus_fund.rows):
        rw = 6 + i
        fill = _GRAY if i % 2 else _WHITE
        _cell(ws, rw, 1, r.year, fill, align=_CENTER)
        _cell(ws, rw, 2, round(r.senior_residual, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 3, round(r.interest_earned, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 4, round(r.deposit_to_reserve, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 5, round(r.reserve_balance, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 6, round(r.excess_to_sub, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 7, round(r.reserve_release, 0) or None, fill, fmt=_DOLLAR)
    ws.freeze_panes = "A6"


# ── Lot-inventory value build (Wells-Fargo-style presentation) ──────────────────
def _build_lot_inventory_value_sheet(ws, cfg, dev):
    rows = dev.lot_inventory_value_build(cfg)
    _title(ws, cfg, [cfg.pid_name, "Builder Lot Inventory Value — Residential",
                "Value of new lots → less lots rolled into homes → net (lagged) → "
                "certified-value adjustments → cumulative 100% lot value → taxable "
                "value @ the residential exemption."], 10)
    hdrs = [(1, "Roll", 9), (2, "Collection\nYear", 10),
            (3, "Value of\nNew Lots", 14), (4, "− Lots to\nHomes", 14),
            (5, "Net Value\nwith Lag", 14), (6, "Adjustments", 14),
            (7, "Cumulative\nFinished Lot Value", 16), (8, "100% Lot\nValue", 15),
            (9, "Taxable\nRatio", 11), (10, "Taxable\nValue of Lots", 15)]
    for col, lbl, w in hdrs:
        _hdr(ws, 5, col, lbl, w)
    # Historical (already-set) roll years — those set before the bonds are dated
    # (roll year < delivery year) plus any entered in the certified inputs table.
    # Their taxable ratios are blue-fonted; projected-forward ratios keep the
    # normal font.  Dynamic off the dated date — no hardcoded years.
    hist_cut = max([cfg.delivery.year - 1, *(cfg.historical_av or {})])
    for i, r in enumerate(rows):
        rw = 6 + i
        fill = _GRAY if i % 2 else _WHITE
        # Highlight every non-zero adjustment — the first-year true-up plug and the
        # recognition / amortization plugs alike.
        hi = _HILITE if r["adjustment"] else fill
        # Certified taxable value (from inputs) → blue font; historical ratio → blue font.
        assessed_font = _CERT_BOLD if r.get("certified") else _BOLD
        ratio_font = _CERT_FONT if (hist_cut is not None and r["av_set"] <= hist_cut) else _BODY
        _cell(ws, rw, 1, r["av_set"], fill, align=_CENTER)
        _cell(ws, rw, 2, r["collection"], fill, align=_CENTER)
        _cell(ws, rw, 3, round(r["new_lots"]) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 4, round(r["lots_to_homes"]) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 5, round(r["net"]) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 6, round(r["adjustment"]) or None, hi, fmt=_DOLLAR)
        _cell(ws, rw, 7, round(r["cumulative"]) or None, fill, _BOLD, fmt=_DOLLAR)
        _cell(ws, rw, 8, round(r["cumulative"]) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 9, r["ratio"], fill, ratio_font, fmt=_RATEFMT)
        _cell(ws, rw, 10, round(r["assessed"]) or None, fill, assessed_font, fmt=_DOLLAR)
    ws.freeze_panes = "C6"
    return {r["collection"]: r["assessed"] for r in rows}


# ── Residential value build (Wells-Fargo-style presentation) ──────────────────
def _build_residential_value_sheet(ws, cfg, dev):
    rows = dev.residential_value_build(cfg)
    _title(ws, cfg, [cfg.pid_name, "Residential Value — Projected Taxable Value",
                "Beginning market value → + new home value added to rolls → + annual "
                "reassessment → + certified-value adjustments → gross market value → "
                "taxable value @ the residential exemption."], 10)
    hdrs = [(1, "Roll", 9), (2, "Tax Rev\nYear", 10),
            (3, "Beginning\nMarket Value", 15), (4, "New Market\nValue Added", 15),
            (5, "Mkt Value\nAdded to Rolls", 15), (6, "Annual\nReassessment", 14),
            (7, "Adjustments", 14), (8, "Gross Market\nValue", 16),
            (9, "Taxable\nRatio", 11), (10, "Taxable\nValue", 15)]
    for col, lbl, w in hdrs:
        _hdr(ws, 5, col, lbl, w)
    # Historical (already-set) roll years — set before the bonds are dated (roll
    # year < delivery year) plus any entered in the certified inputs table → blue ratio.
    hist_cut = max([cfg.delivery.year - 1, *(cfg.historical_av or {})])
    for i, r in enumerate(rows):
        rw = 6 + i
        fill = _GRAY if i % 2 else _WHITE
        hi = _HILITE if r["adjustment"] else fill
        assessed_font = _CERT_BOLD if r.get("certified") else _BOLD
        ratio_font = _CERT_FONT if (hist_cut is not None and r["av_set"] <= hist_cut) else _BODY
        _cell(ws, rw, 1, r["av_set"], fill, align=_CENTER)
        _cell(ws, rw, 2, r["collection"], fill, align=_CENTER)
        _cell(ws, rw, 3, round(r["beginning"]) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 4, round(r["new_added"]) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 5, round(r["added_to_rolls"]) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 6, round(r["reassess"]) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 7, round(r["adjustment"]) or None, hi, fmt=_DOLLAR)
        _cell(ws, rw, 8, round(r["gross"]) or None, fill, _BOLD, fmt=_DOLLAR)
        _cell(ws, rw, 9, r["ratio"], fill, ratio_font, fmt=_RATEFMT)
        _cell(ws, rw, 10, round(r["assessed"]) or None, fill, assessed_font, fmt=_DOLLAR)
    ws.freeze_panes = "C6"
    return {r["collection"]: r["assessed"] for r in rows}


# ── Senior lien annual debt-service coverage ──────────────────────────────────
def _build_coverage_sheet(ws, cfg, sm, senior, refunding_bond=None):
    from .debt_service import senior_coverage_schedule
    _RED = PatternFill("solid", fgColor="FFC7CE")
    _title(ws, cfg, [cfg.pid_name, "Senior Lien — Annual Debt Service Coverage",
                f"Target {cfg.dsc_senior:.2f}x (refunding {cfg.dsc_refunding:.2f}x)"], 6)
    hdrs = [(1, "Collection\nYear", 12), (2, "Net Senior\nRevenue", 16),
            (3, "Senior Net\nDebt Service", 16), (4, "Coverage\nFactor", 13),
            (5, "Target", 11), (6, "Meets\nTarget?", 11)]
    for col, lbl, w in hdrs:
        _hdr(ws, 5, col, lbl, w)
    rows = senior_coverage_schedule(cfg, sm, senior, refunding_bond)
    for i, r in enumerate(rows):
        rw = 6 + i
        meets = r["meets_target"]
        fill = _GRAY if i % 2 else _WHITE
        flag = fill if meets else _RED
        _cell(ws, rw, 1, r["year"], fill, align=_CENTER)
        _cell(ws, rw, 2, r["net_senior_revenue"], fill, fmt=_DOLLAR)
        _cell(ws, rw, 3, r["senior_net_ds"], fill, fmt=_DOLLAR)
        _cell(ws, rw, 4, r["coverage"], flag, fmt='0.00"x"')
        _cell(ws, rw, 5, r["target"], fill, fmt='0.00"x"')
        c = _cell(ws, rw, 6, "Yes" if meets else "No", flag, align=_CENTER)
    ws.freeze_panes = "A6"


# ── Operations & Maintenance (O&M) revenue projection ─────────────────────────
def _build_om_sheet(ws, cfg, sm):
    """
    O&M revenue and expense: the operations mill levy applied to total taxable
    value, collected at the collection rate, plus the personal property uniform
    fee — the total available each year for operations & maintenance — against
    the district's modelled O&M expense, and the surplus or deficit between
    them.  A Utah PID usually runs no operations levy, so the deficit shown here
    is what the debt-service levy is carrying.
    """
    ops_mill = cfg.mill_levy_ops_target
    coll = cfg.tax_collect_mill_prc
    _title(ws, cfg, [cfg.pid_name, "Operations & Maintenance (O&M) Revenue and Expense",
                f"Operations mill levy {ops_mill:.3f} mills @ {coll:.1%} collection"
                + (f"  ·  O&M expense ${cfg.om_expense:,.0f} base, inflating at "
                   f"{cfg.om_growth_rate:.1%}" if cfg.om_expense else
                   "  ·  no O&M expense entered")], 8)
    hdrs = [(1, "Collection\nYear", 12), (2, "Total\nTaxable Value", 16),
            (3, "Operations\nMill Levy", 13),
            (4, f"Total Collections\n@ {coll:.1%}", 16),
            (5, f"Uniform Fee\n@ {cfg.uniform_fee_prc:.0%}", 14),
            (6, "Total Available\nfor O&M", 16),
            (7, "− O&M\nExpense", 14),
            (8, "O&M Surplus /\n(Deficit)", 16)]
    for col, lbl, w in hdrs:
        _hdr(ws, 5, col, lbl, w)

    tot_coll = tot_uniform_fee = tot_avail = 0.0
    tot_om = tot_net_om = 0.0
    for i, r in enumerate(sm.rows):
        rw = 6 + i
        fill = _GRAY if i % 2 else _WHITE
        collections = r.total_av / 1000.0 * ops_mill * coll
        # SOT is a percentage of mill-levy collections (halved below the AV
        # threshold), matching the pledged-revenue waterfall.
        so_rate = (cfg.uniform_fee_prc / 2 if r.total_av < cfg.uniform_fee_av_threshold
                   else cfg.uniform_fee_prc)
        uniform_fee = collections * so_rate
        avail = collections + uniform_fee
        om_expense = r.om_expense
        net_om = avail - om_expense
        tot_coll += collections; tot_uniform_fee += uniform_fee; tot_avail += avail
        tot_om += om_expense; tot_net_om += net_om
        _cell(ws, rw, 1, r.collection_year, fill, align=_CENTER)
        _cell(ws, rw, 2, round(r.total_av, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 3, ops_mill, fill, fmt='0.000', align=_CENTER)
        _cell(ws, rw, 4, round(collections, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 5, round(uniform_fee, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 6, round(avail, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 7, -round(om_expense, 0) or None, fill, fmt=_DOLLAR)
        _cell(ws, rw, 8, round(net_om, 0) or None, fill, font=_BOLD, fmt=_DOLLAR)

    rw = 6 + len(sm.rows)
    _cell(ws, rw, 1, "Total", _TOTAL, font=_TOTAL_FONT, align=_CENTER)
    _cell(ws, rw, 2, None, _TOTAL)
    _cell(ws, rw, 3, None, _TOTAL)
    _cell(ws, rw, 4, round(tot_coll, 0) or None, _TOTAL, font=_TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, rw, 5, round(tot_uniform_fee, 0) or None, _TOTAL, font=_TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, rw, 6, round(tot_avail, 0) or None, _TOTAL, font=_TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, rw, 7, -round(tot_om, 0) or None, _TOTAL, font=_TOTAL_FONT, fmt=_DOLLAR)
    _cell(ws, rw, 8, round(tot_net_om, 0) or None, _TOTAL, font=_TOTAL_FONT, fmt=_DOLLAR)
    ws.freeze_panes = "A6"


def build_excel_report(
    cfg: ModelConfig,
    sm,
    senior,
    first_su,
    refunding_result=None,
    sub_result=None,
    surplus_fund=None,
    dev=None,
    output_path: str = "output/ut_pid_model_output.xlsx",
) -> str:
    wb = openpyxl.Workbook()

    # Total lots for the standard title band (Reimbursement Analysis · district ·
    # N Lots · Tierra Financial Advisors).
    cfg._doc_lots = dev.total_lots if dev is not None else None

    ws = wb.active
    if dev is not None:
        # Two summary tabs: a light headline landing page, then the full detail.
        ws.title = "Summary - Light"
        _build_summary_light_sheet(
            ws, cfg, sm, dev, senior,
            refunding_result.refunding_bond if refunding_result is not None else None)
        _build_summary_av_sheet(wb.create_sheet("Summary - Detail"), cfg, dev, sm)
        _build_lot_inventory_value_sheet(wb.create_sheet("Builder Lot Inventory Value"), cfg, dev)
        _build_residential_value_sheet(wb.create_sheet("Residential Value"), cfg, dev)
        from .residential_report import build_residential_sheet
        build_residential_sheet(wb.create_sheet("Development Projections"), cfg, dev, sm)
    else:
        ws.title = "Summary"
        _build_summary_sheet(ws, cfg, sm)

    _build_ds_sheet(wb.create_sheet("Senior Lien DS - First"), cfg, senior,
                    "Senior Lien Bonds (New Money)", sm=sm,
                    target_coverage=cfg.dsc_senior)

    _build_capi_fund_sheet(wb.create_sheet("CAPI Fund - First"), cfg, senior)

    _build_su_first_sheet(wb.create_sheet("Sources & Uses - First"), cfg, senior,
                          sub_result, surplus_fund, dev)

    _build_coverage_sheet(
        wb.create_sheet("Senior Lien Coverage"), cfg, sm, senior,
        refunding_result.refunding_bond if refunding_result is not None else None)

    _build_om_sheet(wb.create_sheet("O&M Revenue"), cfg, sm)

    if refunding_result is not None:
        _build_ds_sheet(wb.create_sheet("Senior Lien DS - Refunding"), cfg,
                        refunding_result.refunding_bond,
                        "Senior Lien Refunding Bonds — Refinanced New Money",
                        sm=sm, target_coverage=cfg.dsc_refunding)
        _build_su_refunding_sheet(wb.create_sheet("Sources & Uses - Refunding"),
                                  cfg, refunding_result, dev)

    if surplus_fund is not None:
        _build_surplus_sheet(wb.create_sheet("Senior Surplus Fund"), cfg, surplus_fund)

    if sub_result is not None:
        _build_sub_sheet(wb.create_sheet("Subordinate Lien"), cfg, sub_result)

    if refunding_result is not None and getattr(refunding_result, "refunding_sub", None) is not None:
        _build_sub_sheet(wb.create_sheet("Subordinate Lien - Refunding"), cfg,
                         refunding_result.refunding_sub, series_label=" — Refunding")

    # Optional-redemption (call) schedule — derived output.
    _build_call_schedule_sheet(wb.create_sheet("Call Schedule"), cfg, refunding_result)

    # Notes & assumptions — always last.
    _build_notes_sheet(wb.create_sheet("Notes"), cfg, senior, sub_result, refunding_result)

    # Reorder tabs: first-financing block, then the refunding block in the same
    # order. Tabs not listed (Senior Lien Coverage, Notes) follow at the end,
    # with Notes always last.
    desired = [
        "Summary - Light", "Summary - Detail", "Builder Lot Inventory Value",
        "Residential Value", "Summary", "Development Projections",
        "Sources & Uses - First", "Senior Lien DS - First",
        "Subordinate Lien", "Senior Surplus Fund", "CAPI Fund - First", "O&M Revenue",
        "Sources & Uses - Refunding", "Senior Lien DS - Refunding",
        "Subordinate Lien - Refunding",
        "Senior Lien Coverage", "Call Schedule", "Notes",
    ]
    order = {name: i for i, name in enumerate(desired)}
    wb._sheets.sort(key=lambda ws: order.get(ws.title, len(desired)))

    # Tierra logo + "preliminary, subject to change" footer on every sheet.
    for ws in wb.worksheets:
        _apply_branding(ws)

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    wb.save(output_path)
    _inject_footer_logos(output_path)
    return output_path


# ── Optional-redemption (call) schedule (output) ─────────────────────────────
def _build_call_schedule_sheet(ws, cfg, refunding_result=None):
    """
    Optional-redemption (call) schedule in the standard official-statement
    format: each "Date of Redemption" is a 12-month window with the applicable
    redemption premium, stepping down 1.00%/yr from the initial premium to par
    (and par thereafter).  The call-protection period runs from the delivery
    date entered on the Inputs tab, so the first call date moves with delivery.
    """
    from .config import _edate

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 22

    _title(ws, cfg, [cfg.pid_name,
                _county_line(cfg, " — Optional Redemption (Call) Schedule"),
                "Callable at par plus the redemption premium below, stepping down "
                "1.00%/yr to par; call protection runs from the delivery date"], 3)

    month_name = ["", "January", "February", "March", "April", "May", "June", "July",
                  "August", "September", "October", "November", "December"]
    _WRAP_L = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def _table(r, title, subtitle, prem_date, par_only=False):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        c = ws.cell(row=r, column=1, value=title)
        c.fill = _BLUE; c.font = _WHITE_FONT; c.alignment = _LEFT
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        sc = ws.cell(row=r, column=1, value=subtitle)
        sc.font = _BODY; sc.alignment = _WRAP_L
        ws.row_dimensions[r].height = 30
        r += 1
        ws.cell(row=r, column=1).fill = _LIGHT
        ws.cell(row=r, column=1).border = _BORDER
        _hdr(ws, r, 2, "Date of Redemption")
        _hdr(ws, r, 3, "Redemption Premium")
        r += 1

        m, d = prem_date.month, prem_date.day
        # A par-only call has no declining-premium schedule — callable at 100% on
        # and after the first call date.
        steps = 0 if par_only else cfg.call_premium_step_years
        initial = cfg.premium_call_price - 100.0
        end_m = 12 if m == 1 else m - 1               # day before the anniversary
        last_day = "30" if end_m in (4, 6, 9, 11) else ("28" if end_m == 2 else "31")
        for k in range(steps):                        # one row per premium year
            fill = _GRAY if k % 2 else _WHITE
            y0 = prem_date.year + k
            premium = max(initial - k, 0.0)
            window = (f"{month_name[m]} {d}, {y0} to and including "
                      f"{month_name[end_m]} {last_day}, {y0 + 1}")
            _cell(ws, r, 1, None, fill)
            _cell(ws, r, 2, window, fill, align=_LEFT)
            _cell(ws, r, 3, premium / 100.0, fill, fmt='0.00%', align=_CENTER)
            r += 1
        # Par thereafter (the only row for a par-only call)
        fill = _GRAY if steps % 2 else _WHITE
        par_year = prem_date.year + steps
        _cell(ws, r, 1, None, fill)
        _cell(ws, r, 2, f"{month_name[m]} {d}, {par_year}, and thereafter", fill,
              _BOLD, align=_LEFT)
        _cell(ws, r, 3, 0.0, fill, _BOLD, fmt='0.00%', align=_CENTER)
        r += 1
        return r + 1

    r = _table(
        5, "SENIOR LIEN BONDS (New Money)",
        f"Delivered {cfg.delivery:%B %d, %Y}; callable on and after "
        f"{cfg.premium_call_date:%B %d, %Y} at par plus accrued interest and a premium of:",
        cfg.premium_call_date)
    if refunding_result is not None:
        rd = cfg.delivery_refunding
        ref_par = _edate(rd, 12 * cfg.premium_call_years)
        r = _table(
            r, "REFUNDING BONDS",
            f"Delivered {rd:%B %d, %Y}; callable on and after {ref_par:%B %d, %Y} "
            f"at par plus accrued interest (no redemption premium):",
            ref_par, par_only=True)
    ws.freeze_panes = "A5"


# ── Notes & Assumptions (last sheet) ──────────────────────────────────────────
def _build_notes_sheet(ws, cfg, senior, sub_result=None, refunding_result=None):
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 60
    ws.merge_cells("A1:C1")
    _ntitle = (f"{cfg.pid_name} — Notes & Significant Assumptions" if cfg.pid_name
               else "Notes & Significant Assumptions")
    t = ws.cell(row=1, column=1, value=_ntitle)
    t.fill = _BLUE; t.font = _TITLE_FONT; t.alignment = _CENTER

    def _pct(x): return f"{x:.3%}" if x is not None else "—"
    sub_par = sub_result.par_amount if sub_result else cfg.sub_par
    dsrf = senior.dsrf_deposit if senior else 0.0
    capi = sum(p.capitalized_interest for p in senior.schedule) if senior else 0.0

    rows = [
        ("SECTION", "Taxable Value & Revenue", ""),
        ("note", "Primary residential taxable ratio", _pct(cfg.resid_taxable_ratio)),
        ("note", "Taxable-value lag", f"{cfg.av_lag_years} years"
            + (" (two-year level-of-value hold)" if cfg.hold_value_flat else "")),
        ("note", f"{cfg.reassess_frequency} reassessment (residential / commercial)",
            f"{_pct(cfg.reassess_rate)} / {_pct(cfg.reassess_comm_rate)}"),
        ("note", "Debt-service mill levy"
                 + f" (cap {cfg.mill_levy_cap:.3f})",
         f"{cfg.effective_ds_mill_levy:.4f} mills"),
        ("note", "Mill-levy collection / specific-ownership tax",
            f"{_pct(cfg.tax_collect_mill_prc)} / {_pct(cfg.uniform_fee_prc)}"),
        ("note", "County collection cost", _pct(cfg.county_collection_fee)),
        ("note", "District administration (base / growth)",
            f"${cfg.admin_cost:,.0f} / {_pct(cfg.admin_growth_rate)}"),
        ("note", "District O&M expense (base / growth)",
            (f"${cfg.om_expense:,.0f} / {_pct(cfg.om_growth_rate)}"
             if cfg.om_expense else "none entered")),
        ("note", "Home price inflation", _pct(cfg.inflation_rate)),
        ("SECTION", "Bond Structure", ""),
        ("note", "Senior coupon / coverage",
            f"{_pct(cfg.senior_interest_rate)} / {cfg.dsc_senior:.2f}x"),
        ("note", "Subordinate coupon / coverage",
            f"{_pct(cfg.sub_interest_rate)} / {cfg.dsc_sub:.2f}x (cash-flow, accretes)"),
        ("note", "Refunding coupon / coverage",
            f"{_pct(cfg.senior_refunding_interest_rate)} / {cfg.dsc_refunding:.2f}x"),
        ("note", "Capitalized-interest period", f"{cfg.capi_term} months"),
        ("note", "Optional redemption — senior",
            f"callable {cfg.premium_call_date} at {cfg.premium_call_price:.0f}%, "
            f"premium steps down 1.00%/yr to par by {cfg.par_call_date}"),
        ("note", "Optional redemption — refunding",
            f"callable at par {cfg.premium_call_years} years after delivery "
            f"(no redemption premium)"),
        ("note", "Interest-earnings rate (reserve / surplus)", _pct(cfg.interest_earn_rate)),
        ("SECTION", "Dynamically Sized Amounts", ""),
        ("note", "Senior par (sized to revenue at coverage)",
            f"${senior.par_amount:,.0f}" if senior else "—"),
        ("note", "Debt service reserve fund (3-prong: 10% par / max net DS / 125% avg net DS)",
            f"${dsrf:,.0f}"),
        ("note", "Capitalized interest fund", f"${capi:,.0f}"),
        ("note", "Subordinate par (sized to residual surplus)", f"${sub_par:,.0f}"),
    ]
    if refunding_result is not None:
        rows += [
            ("note", "Refunding par (sized to grown revenue)",
                f"${refunding_result.refunding_bond.par_amount:,.0f}"),
            ("note", "Refunding 'new money' reimbursement",
                f"${refunding_result.new_money_reimbursement:,.0f}"),
        ]
    rows += [
        ("SECTION", "Methodology Notes", ""),
        ("text", "1. Senior principal is sized year-by-year so net debt service is "
            "covered by pledged net revenue at the target coverage ratio "
            "(a revenue-wrap), with the par solved by bisection.", ""),
        ("text", "2. The DSRF, subordinate par, and refunding surplus / escrow are "
            "computed dynamically from the district's future taxable value and the "
            "resulting tax revenue — they are not user inputs.", ""),
        ("text", "3. The subordinate lien is a cash-flow bond: unpaid interest accretes "
            "at the subordinate rate and is repaid from residual surplus after the "
            "senior lien's reserve is satisfied.", ""),
        ("text", "4. Bonds are priced to worst call; per-maturity coupon/yield scales and "
            "term bonds may be supplied on the Debt Structure input sheet.", ""),
        ("text", "5. This information should be read only in connection with the full "
            "model and its forecast exhibits.", ""),
    ]

    r = 3
    for kind, label, val in rows:
        if kind == "SECTION":
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            c = ws.cell(row=r, column=2, value=label)
            c.fill = _LIGHT; c.font = _HDR_FONT; c.alignment = _LEFT
        elif kind == "text":
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            c = ws.cell(row=r, column=2, value=label)
            c.font = _BODY; c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            ws.row_dimensions[r].height = 28
        else:
            ws.cell(row=r, column=2, value=label).font = _BODY
            ws.cell(row=r, column=3, value=val).font = _BOLD
        r += 1
