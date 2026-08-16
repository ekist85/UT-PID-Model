"""
memo.py — Generate a Tierra-style reimbursement-analysis memo (HTML) for a
Utah Public Infrastructure District financing, populated from the model outputs.

Mirrors the Texas MUD and Arizona CFD reimbursement memos in look and structure
(right-floated absorption table, assumption bullets, bond/reimbursement table,
sources & uses), but states Utah assumptions — the levy held under the
§ 17D-4-303 cap, the 45% primary residential exemption applied to homes and to
builder lot inventory, annual reassessment, capitalized interest, the 3-prong DSRF, the
subordinate cash-flow note and the senior refunding — and pulls the bond program,
development schedule and reimbursement figures from the model objects.
"""

from __future__ import annotations

import base64
import os
from datetime import date


def _logo_data_uri() -> str:
    """Tierra wordmark as a base64 PNG data URI (self-contained).  '' if missing."""
    try:
        path = os.path.join(os.path.dirname(__file__), "assets", "tierra_logo.png")
        with open(path, "rb") as fh:
            return "data:image/png;base64," + base64.b64encode(fh.read()).decode("ascii")
    except Exception:
        return ""


# Right-floated table style that Word honors (from the Texas / AZ memos).  Word
# wraps body text around a TABLE carrying mso-table-float (it ignores float on a
# div), so this lives inline on the <table>; a <div style="clear:both"> after the
# wrapping paragraphs drops the following block back to full width.
_FLOAT_TBL = ("float:right; margin-left:14px; margin-bottom:6px; "
              "mso-table-float:right; mso-table-anchor-vertical:paragraph; "
              "mso-table-anchor-horizontal:column; mso-table-left:right; mso-table-top:0in;")

_CSS = """
@page WordSection1 { size: 8.5in 11.0in; margin: 0.5in;
  mso-margin-top-alt: 0.5in; mso-margin-bottom-alt: 0.5in;
  mso-margin-left: 0.5in; mso-margin-right: 0.5in;
  mso-footer: f1; mso-footer-margin: 0.3in; mso-paper-source: 0; }
div.WordSection1 { page: WordSection1; }
@page { size: letter; margin: 0.5in; }
@media screen { body { margin: 0.5in; } }
.wd-footer { display: none; mso-element: footer; }
body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000; line-height: 1.35; margin: 0; }
.page-header { width: 100%; border-collapse: collapse; border-bottom: 2px solid #002060; margin-bottom: 14px; }
.page-header td { padding: 0 0 5px 0; vertical-align: middle; }
.brand { font-size: 16pt; font-weight: bold; color: #002060; }
.hdr-right { text-align: right; font-style: italic; font-size: 8pt; color: #002060; line-height: 1.6; }
h1 { font-size: 13pt; color: #002060; margin: 4px 0 2px 0; }
p { margin: 0 0 7px 0; }
ul { margin: 3px 0 10px 0; padding-left: 22px; }
li { margin-bottom: 3px; }
.re-line { font-weight: bold; margin: 8px 0; }
td.cap { text-align: right; font-size: 8pt; font-style: italic; color: #555;
  border-bottom: none; padding-top: 4px; white-space: nowrap; }
table.data { border-collapse: collapse; font-size: 9pt; white-space: nowrap; }
table.data th { background: #002060; color: #fff; padding: 4px 14px; text-align: center; font-weight: bold; }
table.data td { padding: 3px 14px; border-bottom: 1px solid #d9d9d9; text-align: right; }
table.data td.c { text-align: center; }
table.data td.l { text-align: left; }
table.data tr.total td { background: #BDD7EE; font-weight: bold; color: #002060; border-top: 2px solid #002060; }
table.data tr.sub td { background: #EBF3FB; font-style: italic; }
.note { font-size: 8pt; font-style: italic; color: #555; }
.sig { margin-top: 18px; }
.disclaimer { font-size: 8pt; color: #777; margin-top: 16px; border-top: 1px solid #ccc; padding-top: 6px; }
"""


def _money(x):
    # Accounting style: negatives in parentheses.  A refunding under Utah's
    # fixed levy caps can genuinely return less than it costs, so this is a
    # real case for a PID rather than a defensive flourish.
    return f"(${abs(x):,.0f})" if x < 0 else f"${x:,.0f}"


def _md(d):
    """m/d/yyyy — non-padded month and day."""
    return f"{d.month}/{d.day}/{d.year}"


def build_memo_html(cfg, sm, senior, su, sub=None, refunding=None, dev=None,
                    output_path: str = "output/ut_pid_model_memo.html",
                    developer: str = "[Developer / Master Developer]") -> str:
    """Render the reimbursement memo HTML from the model objects and write it.

    ``cfg`` ModelConfig, ``sm`` built SummaryModel, ``senior`` the senior tranche,
    ``su`` the first-financing SourcesUses, ``sub`` the SubLienResult (optional),
    ``refunding`` the RefundingResult (optional), ``dev`` DeveloperProjections.
    """
    dev = dev or sm.dev
    _t = date.today()
    today = f"{_t.strftime('%B')} {_t.day}, {_t.year}"
    developer = developer if developer and developer != "[Developer / Master Developer]" \
        else (cfg.developer or "[Developer / Master Developer]")
    district = cfg.pid_name or "Utah Public Infrastructure District"
    # Addressee block — developer name over street and city/state/zip, each line
    # falling back to a bracketed placeholder when the input is blank.
    addr_lines = [developer,
                  cfg.developer_address or "[Address]",
                  cfg.developer_city_state_zip or "[City/State/Zip]"]
    addressee_html = "<br>".join(addr_lines)

    # ── Development / absorption summary ───────────────────────────────────
    closings = {y: int(round(v)) for y, v in dev.home_closings.items() if round(v)}
    deliveries = {y: int(round(v)) for y, v in dev.lot_deliveries.items() if round(v)}
    total_homes = sum(closings.values())
    total_lots = sum(deliveries.values())
    wasp = getattr(dev, "base_asp", 0.0)
    peak_av = max((r.total_av for r in sm.rows), default=0.0)
    total_mkt = total_homes * wasp

    years = sorted(set(closings) | set(deliveries))
    abs_rows = "".join(
        f"<tr style='background:{'#F2F2F2' if i % 2 else '#FFFFFF'}'>"
        f"<td style='text-align:center;padding:0 8px'>{y}</td>"
        f"<td style='text-align:right;padding:0 8px'>{closings.get(y, 0):,}</td>"
        f"<td style='text-align:right;padding:0 8px'>{deliveries.get(y, 0):,}</td></tr>"
        for i, y in enumerate(years))

    # ── Rates / structure pulled from the config ──────────────────────────
    mill_ds = cfg.mill_levy_ds_target
    mill_ops = cfg.mill_levy_ops_target
    mill_total = mill_ds + mill_ops
    eff_mill = cfg.effective_ds_mill_levy


    # Lot-inventory / nonresidential taxable ratio phases by roll year
    # (SB24-233); show the range applied over roll years that carry lot-inventory
    # value (so an early 29% year with no lots doesn't widen the range).
    _vl = sorted({cfg.lot_inventory_ratio(y)
                  for y in range(cfg.first_year, cfg.senior_final_year)
                  if dev.vacant_lot_market_value(y) > 0})
    _vl_txt = (f"{_vl[0]:.1%}" if len(_vl) == 1
               else f"{_vl[0]:.1%}&ndash;{_vl[-1]:.1%}") if _vl else f"{cfg.lot_inventory_taxable_ratio:.1%}"

    from .config import PID_STATUTORY_LEVY_CAP
    statutory_mills = PID_STATUTORY_LEVY_CAP * 1000.0
    _mill_bullet = (
        f"A district levy of {mill_total:.3f} mills is assumed &mdash; {mill_ds:.3f} mills for "
        f"debt service"
        + (f" and {mill_ops:.3f} mills for operations" if mill_ops else "")
        + f". The levy is capped at the most restrictive of {statutory_mills:.3f} mills "
          f"(&sect;&nbsp;17D-4-303), the governing document, and the indentures &mdash; "
          f"{cfg.mill_levy_cap:.3f} mills here, so {eff_mill:.3f} mills is levied. Utah caps are "
          f"fixed rates per dollar of taxable value and do not float with the residential "
          f"exemption, so there is no Colorado-style &ldquo;Gallagherization&rdquo; adjustment. "
          f"Because the required mill levy stays within those caps, the District need not give "
          f"Truth in Taxation notice or hold a hearing to impose it.")

    assumptions = [
        (f"Development delivers {total_homes:,} homes on {total_lots:,} lots (weighted-average "
         f"sales price &asymp; {_money(wasp)}); value phases onto the tax roll as homes close."),
        (f"Lot values are estimated at {cfg.platted_lot_value:.0%} of the average selling price "
         f"(platted-lot value); a delivered lot&rsquo;s value rolls off lot-inventory value and the "
         f"finished home rolls on as each home closes."),
        (f"Utah taxable ratios: {cfg.resid_taxable_ratio:.2%} on primary residential value "
         f"(the 45% exemption, &sect;&nbsp;59-2-103) and "
         f"{_vl_txt} on builder lot inventory, which carries the same exemption where the "
         f"assessor determines the property will be a primary residence once occupied "
         f"(Utah Admin. Code R884-24P-52)."),
        _mill_bullet,
        (f"County assessors revalue annually (&sect;&nbsp;59-2-303.1); value already on the roll "
         f"grows at {cfg.reassess_rate:.1%} a year. Value created in a calendar year lands on the "
         f"following 1&nbsp;January roll, is billed that November, and pays debt service the next "
         f"1&nbsp;March &mdash; a {cfg.av_lag_years}-year lag from creation to the payment it "
         f"supports."),
        (f"Home prices and market values escalate at {cfg.inflation_rate:.1%} annually."),
        (f"A {cfg.tax_collect_mill_prc:.1%} tax-collection rate and a {cfg.interest_earn_rate:.2%} "
         f"interest-earnings rate on fund balances are assumed."),
        (f"Senior new-money bonds carry a {cfg.senior_interest_rate:.2%} interest rate, a "
         f"{cfg.final_mat_yrs}-year final maturity, and are sized to a minimum "
         f"{cfg.dsc_senior:.2f}x debt-service coverage ratio."),
        (f"The senior debt-service reserve fund is sized to the 3-prong test (least of 10% of par, "
         f"maximum annual net debt service, or 125% of average annual net debt service); interest "
         f"earnings on fund balances offset debt service."),
    ]
    assumptions.append(
        f"Utah property taxes are levied on value as of 1&nbsp;January and are due in a single "
        f"payment on 30&nbsp;November, so district principal is structured on 1&nbsp;March with "
        f"interest on 1&nbsp;March and 1&nbsp;September &mdash; each year&rsquo;s collections are "
        f"in hand before the payment they support. Delinquent taxes carry a penalty of the greater "
        f"of 2.5% or $10, and the Public Infrastructure District Act permits the District to impose "
        f"an additional 7% annual penalty; no penalty revenue is credited here.")
    if cfg.uniform_fee_prc:
        assumptions.append(
            f"Registered personal property pays a uniform fee in lieu of ad valorem tax "
            f"(&sect;&nbsp;59-2-405), distributed to taxing entities in the same proportion as real "
            f"property tax; {cfg.uniform_fee_prc:.2%} of the levy is credited.")
    else:
        assumptions.append(
            "No credit is taken for the personal property uniform fee "
            "(&sect;&nbsp;59-2-405), which would otherwise be distributed to the District in the "
            "same proportion as real property tax.")
    assumptions.append(
        f"County assessing and collecting is funded by a separate statewide levy on property "
        f"(&sect;&nbsp;59-2-1602) rather than a deduction from the District&rsquo;s distribution, "
        f"so {cfg.county_collection_fee:.2%} is netted from pledged revenue "
        f"(a Colorado metropolitan district would lose roughly 1.50% to the county treasurer).")
    if cfg.capi_term > 0:
        assumptions.append(
            f"Capitalized interest (up to {cfg.capi_term} months) is funded from bond proceeds to "
            f"cover debt service during the initial absorption ramp-up.")
    if sub is not None and sub.par_amount > 0:
        assumptions.append(
            f"A subordinate-lien cash-flow note ({sub.rate:.2%}, interest accretes) is sized to the "
            f"residual surplus at a {sub.coverage:.2f}x coverage factor and repaid from surplus revenues.")
    if refunding is not None:
        rd = cfg.delivery_refunding
        assumptions.append(
            f"The senior new-money bonds are refunded at the first optional call "
            f"({_md(cfg.premium_call_date)}, {cfg.premium_call_price:.0f}% call price) at "
            f"{cfg.senior_refunding_interest_rate:.2%} / {cfg.dsc_refunding:.2f}x coverage to "
            f"generate additional new-money reimbursement.")
    assumptions_html = "".join(f"<li>{a}</li>" for a in assumptions)

    # Certified value already on the rolls (from the inputs), when set.
    cv_html = ""
    if cfg.current_certified_value > 0:
        _asof = f" as of {_md(cfg.certification_date)}" if cfg.certification_date else ""
        cv_html = (f"<p>The District&rsquo;s certified taxable value{_asof} is "
                   f"<strong>{_money(cfg.current_certified_value)}</strong>; the model trues Total "
                   f"Taxable Value to this certified roll in the certification year.</p>")

    # ── Bond program / reimbursement table ─────────────────────────────────
    sr_par = senior.par_amount
    sub_par = sub.par_amount if sub is not None else 0.0
    first_reimb = su.reimbursement
    first_par = sr_par + sub_par
    cum = first_reimb
    rows = [
        (f"<tr><td class='l'>Senior New-Money Bonds, Series {cfg.delivery.year}A</td>"
         f"<td class='c'>{_md(cfg.delivery)}</td><td>{_money(sr_par)}</td>"
         f"<td class='c'>{cfg.senior_interest_rate:.2%}</td><td>&mdash;</td><td>&mdash;</td></tr>"),
    ]
    if sub is not None and sub_par > 0:
        rows.append(
            f"<tr><td class='l'>Subordinate Lien Cash-Flow Note</td>"
            f"<td class='c'>{_md(cfg.delivery)}</td><td>{_money(sub_par)}</td>"
            f"<td class='c'>{sub.rate:.2%}</td><td>&mdash;</td><td>&mdash;</td></tr>")
    rows.append(
        f"<tr class='sub'><td class='l'>First Financing &mdash; Net Reimbursement</td>"
        f"<td class='c'></td><td>{_money(first_par)}</td><td class='c'></td>"
        f"<td>{_money(first_reimb)}</td><td>{_money(cum)}</td></tr>")
    total_reimb = first_reimb
    if refunding is not None:
        rd = cfg.delivery_refunding
        rb = refunding.refunding_bond
        add = refunding.new_money_reimbursement
        cum += add
        total_reimb += add
        rows.append(
            f"<tr><td class='l'>Senior Refunding Bonds, Series {rd.year}</td>"
            f"<td class='c'>{_md(rd)}</td><td>{_money(rb.par_amount)}</td>"
            f"<td class='c'>{cfg.senior_refunding_interest_rate:.2%}</td>"
            f"<td>{_money(add)}</td><td>{_money(cum)}</td></tr>")
    rows.append(
        f"<tr class='total'><td class='c'>Total Developer Reimbursement</td><td></td><td></td>"
        f"<td></td><td>{_money(total_reimb)}</td><td>{_money(total_reimb)}</td></tr>")
    bond_rows = "".join(rows)

    # ── Sources & Uses (first financing) ──────────────────────────────────
    # Uses of Funds split per series (senior new-money vs. subordinate cash-flow
    # note), mirroring the "Sources & Uses" tab so the reimbursement split is
    # visible — same derivation as the workbook.
    yr = cfg.delivery.year
    has_sub = sub is not None and sub_par > 0
    _capi = sum(p.capitalized_interest for p in senior.schedule)
    _dsrf = senior.dsrf_deposit
    _prem_sr = senior.total_premium
    _uwd_sr = cfg.uwd_senior * sr_par
    _uwd_sb = cfg.uwd_sub * sub_par
    _coi = cfg.coi
    _reimb_sr = sr_par + _prem_sr - _dsrf - _capi - _uwd_sr - _coi
    _reimb_sb = sub_par - _uwd_sb
    # (label, senior, sub) — sub = None means "not applicable" (blank cell).
    uses_data = [
        ("Estimated Reimbursement Amount", _reimb_sr, _reimb_sb),
        ("Debt Service Reserve Fund", _dsrf, 0.0),
        ("Capitalized Interest", _capi, 0.0),
        ("Underwriter's Discount", _uwd_sr, _uwd_sb),
        ("Costs of Issuance", _coi, None),
    ]
    total_sr = _reimb_sr + _dsrf + _capi + _uwd_sr + _coi
    total_sb = _reimb_sb + _uwd_sb
    total_uses = total_sr + total_sb

    def _row_total(sv, bv):
        return _money((sv or 0.0) + (bv or 0.0)) if bv is not None else _money(sv)

    if has_sub:
        u_rows = "".join(
            f"<tr><td class='l'>{lbl}</td><td>{_money(sv)}</td>"
            f"<td>{_money(bv) if bv is not None else ''}</td>"
            f"<td>{_row_total(sv, bv)}</td></tr>"
            for lbl, sv, bv in uses_data)
        uses_table = (
            '<table class="data" style="width:5.4in">\n'
            '  <tr><th style="text-align:left">Uses of Funds — First Financing</th>'
            f'<th>Senior Lien Bonds<br>Series {yr}A</th>'
            f'<th>Subordinate Lien<br>Cash-Flow Note Series {yr}B</th><th>Total</th></tr>\n'
            f'  {u_rows}\n'
            f'  <tr class="total"><td class="l">Total Uses</td><td>{_money(total_sr)}</td>'
            f'<td>{_money(total_sb)}</td><td>{_money(total_uses)}</td></tr>\n'
            '</table>')
    else:
        u_rows = "".join(
            f"<tr><td class='l'>{lbl}</td><td>{_row_total(sv, bv)}</td></tr>"
            for lbl, sv, bv in uses_data)
        uses_table = (
            '<table class="data" style="width:3.3in">\n'
            '  <tr><th style="text-align:left">Uses of Funds — First Financing</th>'
            '<th>Amount</th></tr>\n'
            f'  {u_rows}\n'
            f'  <tr class="total"><td class="l">Total Uses</td><td>{_money(total_uses)}</td></tr>\n'
            '</table>')

    logo_uri = _logo_data_uri()
    brand_cell = (f'<img src="{logo_uri}" alt="Tierra Financial Advisors" width="60" height="45" '
                  f'style="width:60px;height:45px;vertical-align:middle">' if logo_uri
                  else '<span class="brand">Tierra Financial Advisors</span>')
    header_html = (
        f'<table class="page-header"><tr><td>{brand_cell}</td>'
        f'<td class="hdr-right">Reimbursement Analysis &ndash; {district}<br>{today}</td>'
        f'</tr></table>')

    html = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<title>Reimbursement Analysis — {district}</title><style>{_CSS}</style></head><body>

<div class="wd-footer" id="f1">
  <p style="font-style:italic;font-size:8pt;color:#555;margin:0;text-align:left">Preliminary, subject to change.</p>
</div>

<div class="WordSection1">

{header_html}

<h1>Reimbursement Analysis — {district}</h1>
<p>{today}</p>
<p>{addressee_html}</p>
<p>Dear {developer},</p>
<p class="re-line">RE: {district} &ndash; Reimbursement Analysis</p>

<table class="data" style="width:2.1in; {_FLOAT_TBL}">
  <tr><th>Year</th><th>Homes</th><th>Lots</th></tr>
  {abs_rows}
  <tr class="total"><td class="c">Total</td><td>{total_homes:,}</td><td>{total_lots:,}</td></tr>
  <tr><td class="cap" colspan="3">WASP &asymp; {_money(wasp)} · Peak Taxable Value &asymp; {_money(peak_av)}</td></tr>
</table>

<p>We have prepared the attached reimbursement analysis for {district} (the
&ldquo;District&rdquo;) as requested by the developer. The analysis sizes a senior / subordinate
Utah public infrastructure district financing and a subsequent senior refunding, held within the
levy caps that bind a public infrastructure district under the Public Infrastructure District
Act, Title&nbsp;17D, Chapter&nbsp;4, Utah Code. The following is a summary of the assumptions used
in the analysis:</p>

{cv_html}

<ul>{assumptions_html}</ul>
<div style="clear:both;"></div>

<p>The following table reflects the assumptions above. The estimated bond reimbursement amount is
the net proceeds remaining after subtracting the debt-service reserve fund, capitalized interest,
underwriter&rsquo;s discount and costs of issuance.</p>

<table class="data">
  <tr><th>Financing</th><th>Delivery Date</th><th>Par Amount</th><th>Rate</th>
      <th>Est. Reimbursement</th><th>Cumulative Reimb.</th></tr>
  {bond_rows}
</table>

<p style="margin-top:10px">The first financing (senior new-money bonds{' plus the subordinate-lien note' if sub is not None and sub_par > 0 else ''})
applies its sources as follows:</p>

{uses_table}

<p style="margin-top:12px">Please call us if you have any questions or if we can be of any further assistance.</p>
<div class="sig">Sincerely,<br><br>Evan Kist, CFA&nbsp;&nbsp;|&nbsp;&nbsp;Tierra Financial Advisors<br>
M: 817-357-9192&nbsp;&nbsp;E: edkist@tierrafa.com</div>

<div class="disclaimer">Preliminary and subject to change. This analysis is based on developer-provided
assumptions and is for discussion purposes only; it is not a recommendation or an offer to sell securities.</div>

</div>
</body></html>"""

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write(html)
    return output_path
