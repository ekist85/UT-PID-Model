"""
inputs.py — read a model run from an Excel inputs workbook, and write the
blank template for one.

The template has three sheets, following the convention the Texas reimbursement
model uses in this organisation:

    Inputs                   column B = label, C = value, D = range name
    Residential Development  one row per product, one column per year
    Commercial Development   one row per commercial parcel, one column per year

`write_template()` generates it pre-filled with the current defaults, so an
analyst can copy it, edit the yellow cells and re-run.
"""

from __future__ import annotations

import json
import os
import warnings
from dataclasses import fields, replace
from datetime import date, datetime
from typing import Any, Optional

from .config import ModelConfig
from .development import CommercialProduct, DevelopmentProjections, Product
from .workbook import (FMT_DATE, FMT_INT, FMT_MONEY, FMT_PCT, INPUT_FILL,
                       SECTION_FONT, SUB_FONT, TITLE_FONT, _header, _put,
                       _widths)

#: Range name → ModelConfig attribute, for the fields an analyst edits.
RANGE_TO_ATTR: dict[str, str] = {
    "PID": "district_name", "CITY": "city", "COUNTY": "county",
    "DEVELOPER": "developer", "TITLE4": "scenario_label",
    "SECOND_FINANCING": "second_financing", "REFUND_FINANCING": "refund_financing",
    "COI": "coi", "COI_REFUNDING": "coi_refunding",
    "UWD_SENIOR": "uwd_senior", "UWD_SUB": "uwd_sub",
    "UWD_SENIOR_REFUNDING": "uwd_senior_refunding",
    "COUNTY_COLLECTION_FEE": "county_treasurer_fee",
    "TRUSTEE_FEE": "trustee_fee", "TRUSTEE_FEE_SUB": "trustee_fee_sub",
    "DELIVERY": "delivery", "CAPI": "capi", "CAPI_TERM": "capi_term_months",
    "SURPLUS_ON_OFF": "surplus_on_off",
    "SURPLUS_RELEASE_SIZING": "surplus_release_sizing",
    "SUB_SIZING_THRESHOLD": "sub_sizing_threshold",
    "SUB_PAR_OVERRIDE": "sub_par_override",
    "PREMIUM_CALL_YEARS": "premium_call_years",
    "PREMIUM_CALL_FIRST_PRICE": "premium_call_price",
    "PAR_CALL_YEARS": "par_call_years",
    "FINAL_MAT_YRS": "final_mat_yrs", "FINAL_MAT_SUB_YRS": "final_mat_sub_yrs",
    "FINAL_MAT_YRS_REFUNDING": "final_mat_yrs_refunding",
    "PRIN_MATURITY": "prin_maturity",
    "PRIN_MATURITY_DAY_SENIOR": "prin_maturity_day_senior",
    "PRIN_MATURITY_DAY_SUB": "prin_maturity_day_sub",
    "IG_RATED": "ig_rated",
    "SENIOR_RATE_IG": "senior_rate_ig", "SENIOR_RATE_NR": "senior_rate_nr",
    "SUB_RATE_IG": "sub_rate_ig", "SUB_RATE_NR": "sub_rate_nr",
    "SENIOR_REFUNDING_INTEREST_RATE": "senior_refunding_interest_rate",
    "DSC_SENIOR_LIEN_BONDS": "dsc_senior_lien_bonds",
    "DSC_SUB_LIEN_BONDS": "dsc_sub_lien_bonds",
    "DSC_REFUNDING_BONDS": "dsc_refunding_bonds",
    "FIRST_YEAR": "first_year", "RESID_DELIVERY_YEAR": "resid_delivery_year",
    "INFLATION_RATE": "inflation_rate",
    "INFLATION_RATE_COMM_SALES": "inflation_rate_comm_sales",
    "REASSESS_FREQUENCY": "reassess_frequency",
    "REASSESS_RATE": "reassess_rate_resid",
    "REASSESS_RATE_SUBORDINATE": "reassess_rate_resid_sub",
    "REASSESS_COMM_RATE": "reassess_rate_comm",
    "RESID_TAXABLE_RATIO_PRIOR": "resid_taxable_ratio_prior",
    "RESID_TAXABLE_RATIO": "resid_taxable_ratio",
    "TAX_COLLECT_MILL_PRC": "tax_collect_mill_prc",
    "UNIFORM_FEE_PRC": "uniform_fee_prc",
    "UNIFORM_FEE_AV_THRESHOLD": "uniform_fee_av_threshold",
    "INTEREST_EARN_RATE": "interest_earn_rate",
    "GALLAGHERIZATION": "gallagherization",
    "MILL_LEVY_GOVERNING_DOC": "mill_levy_governing_doc",
    "MILL_LEVY_COMM": "mill_levy_comm",
    "MILL_LEVY_OPS_TARGET": "mill_levy_ops_target",
    "MILL_LEVY_CAP_TOTAL": "mill_levy_cap_total",
    "CONTRIBUTION_RATE": "contribution_rate",
    "ADMIN_COST_BASE": "admin_cost_base",
    "ADMIN_COST_AV_LIMIT": "admin_cost_av_limit",
    "ADMIN_COST_GROWTH": "admin_cost_growth",
    "DISTRICT_COST_START_YEAR": "district_cost_start_year",
    "SYSTEM_DEVELOPMENT_FEE": "system_development_fee",
    "STATE_ASSESSED": "state_assessed",
    "CENTRALLY_ASSESSED_VALUE": "centrally_assessed_value",
    "RESID_NEW_VALUE_ADD": "resid_new_value_add",
    "RESID_NEW_VALUE_ADD_SUB": "resid_new_value_add_sub",
    "COMM_NEW_VALUE_ADD": "comm_new_value_add",
    "COMM_NEW_VALUE_ADD_SUB": "comm_new_value_add_sub",
    "CENTRALLY_ASSESSED_SENIOR": "centrally_assessed_senior",
    "CENTRALLY_ASSESSED_SUB": "centrally_assessed_sub",
    "HYPOTHETICAL_SCENARIO": "hypothetical_scenario",
    "LOT_DELIVERY_SCENARIO": "lot_delivery_scenario",
    "ABSORPTION_SCENARIO": "absorption_scenario",
    "HOME_FIRST_CLOSE_MONTHS": "home_first_close_months",
    "COMMERCIAL_LAG_YEARS": "commercial_lag_years",
    "HOME_LOT_DELIVERY_LEAD_MONTHS": "home_lot_delivery_lead_months",
    "PLATTED_COMM_LOT_VALUE": "platted_comm_lot_value",
    "PLATTED_LOT_VALUE": "platted_lot_value",
    "DEVELOPED_LOT_VALUE": "developed_lot_value",
    "CENTRALLY_ASSESSED_RATIO": "centrally_assessed_ratio",
    "VALUE_LAG_YEARS": "value_lag_years",
}

#: Written into the template but stored on DevelopmentProjections, not ModelConfig.
DEV_START_RANGE = "DEV_START_YEAR"

_DATE_FIELDS = {"delivery", "resid_delivery_year", "inflation_step_start_year"}
_INT_FIELDS = {"capi_term_months", "premium_call_years", "par_call_years",
               "final_mat_yrs", "final_mat_sub_yrs", "final_mat_yrs_refunding",
               "prin_maturity", "prin_maturity_day_senior", "prin_maturity_day_sub",
               "first_year", "commercial_lag_years", "home_first_close_months",
               "home_lot_delivery_lead_months", "value_lag_years",
               "district_cost_start_year", "inflation_step_years"}


def _as_date(v: Any) -> Optional[date]:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        return datetime.fromisoformat(v).date()
    return None


# ── Reading ───────────────────────────────────────────────────────────────────

def load_inputs(path: str) -> tuple[ModelConfig, DevelopmentProjections]:
    """Read an inputs workbook (.xlsx) or a JSON scenario file."""
    if path.lower().endswith(".json"):
        return _load_json(path)
    return _load_excel(path)


def _load_json(path: str) -> tuple[ModelConfig, DevelopmentProjections]:
    with open(path) as f:
        blob = json.load(f)
    overrides: dict[str, Any] = {}
    for key, value in (blob.get("config") or {}).items():
        if key in _DATE_FIELDS:
            value = _as_date(value)
        overrides[key] = value
    cfg = ModelConfig(**overrides)

    dev_blob = blob.get("development") or {}
    products = [
        Product(name=p["name"], product_type=p.get("product_type", "SFD"),
                total_units=int(p.get("total_units", 0)), asp=float(p.get("asp", 0)),
                lot_delivery={int(k): int(v) for k, v in (p.get("lot_delivery") or {}).items()},
                home_closings={int(k): int(v) for k, v in (p.get("home_closings") or {}).items()})
        for p in dev_blob.get("products", [])
    ]
    commercial = [
        CommercialProduct(name=c["name"], total_sf=float(c.get("total_sf", 0)),
                          value_per_sf=float(c.get("value_per_sf", 0)),
                          sf_delivered={int(k): float(v) for k, v in (c.get("sf_delivered") or {}).items()},
                          sf_sold={int(k): float(v) for k, v in (c.get("sf_sold") or {}).items()})
        for c in dev_blob.get("commercial", [])
    ]
    dev = DevelopmentProjections(
        start_year=int(dev_blob.get("start_year", cfg.delivery.year - 4)),
        products=products, commercial=commercial,
        existing_home_market_value=float(dev_blob.get("existing_home_market_value", 0)),
        existing_lot_value=float(dev_blob.get("existing_lot_value", 0)))
    if products:
        cfg = replace(cfg, asp=[p.asp for p in products][:12] + [0.0] * max(0, 12 - len(products)),
                      product_labels=[p.name for p in products][:12] + [""] * max(0, 12 - len(products)))
    return cfg, dev


def _load_excel(path: str) -> tuple[ModelConfig, DevelopmentProjections]:
    import openpyxl
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path, data_only=True)

    named: dict[str, Any] = {}
    ws = wb["Inputs"]
    for row in ws.iter_rows(min_col=2, max_col=4):
        cells = list(row)
        if len(cells) < 3:
            continue
        name, value = cells[2].value, cells[1].value
        if isinstance(name, str) and name.strip():
            named[name.strip()] = value

    overrides: dict[str, Any] = {}
    valid = {f.name for f in fields(ModelConfig)}
    for range_name, attr in RANGE_TO_ATTR.items():
        if range_name not in named or attr not in valid:
            continue
        value = named[range_name]
        if value is None:
            continue
        if attr in _DATE_FIELDS:
            value = _as_date(value)
        elif attr in _INT_FIELDS:
            value = int(value)
        overrides[attr] = value
    cfg = ModelConfig(**overrides)

    products = _read_products(wb["Residential Development"])
    commercial = (_read_commercial(wb["Commercial Development"])
                  if "Commercial Development" in wb.sheetnames else [])
    start = named.get(DEV_START_RANGE)
    start = int(start) if start else min(
        (min(p.lot_delivery) for p in products if p.lot_delivery),
        default=cfg.delivery.year) - 4
    dev = DevelopmentProjections(start_year=start, products=products,
                                 commercial=commercial)
    if products:
        asp = [p.asp for p in products][:12]
        labels = [p.name for p in products][:12]
        cfg = replace(cfg,
                      asp=asp + [0.0] * (12 - len(asp)),
                      product_labels=labels + [""] * (12 - len(labels)))
    return cfg, dev


def _year_columns(ws, header_row: int) -> dict[int, int]:
    out: dict[int, int] = {}
    for col in range(6, ws.max_column + 1):
        v = ws.cell(header_row, col).value
        if isinstance(v, (int, float)) and 1990 < int(v) < 2120:
            out[int(v)] = col
    return out


def _read_products(ws) -> list[Product]:
    """Rows 6+ are lot delivery; the closings block starts after a blank row."""
    years = _year_columns(ws, 5)
    products: dict[str, Product] = {}
    order: list[str] = []
    block = "lot_delivery"
    for r in range(6, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if not name:
            continue
        if isinstance(name, str) and name.strip().upper().startswith("HOME CLOSINGS"):
            block = "home_closings"
            years = _year_columns(ws, r + 1) or years
            continue
        if not isinstance(name, str) or name.strip().lower() in ("product", "total"):
            continue
        key = name.strip()
        if key not in products:
            products[key] = Product(name=key)
            order.append(key)
        p = products[key]
        if block == "lot_delivery":
            p.product_type = ws.cell(r, 3).value or "SFD"
            p.total_units = int(ws.cell(r, 4).value or 0)
            p.asp = float(ws.cell(r, 5).value or 0)
        target = p.lot_delivery if block == "lot_delivery" else p.home_closings
        for year, col in years.items():
            v = ws.cell(r, col).value
            if v:
                target[year] = int(v)
    return [products[k] for k in order if products[k].total_units]


def _read_commercial(ws) -> list[CommercialProduct]:
    years = _year_columns(ws, 5)
    out: list[CommercialProduct] = []
    for r in range(6, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if not isinstance(name, str) or not name.strip():
            continue
        c = CommercialProduct(name=name.strip(),
                              total_sf=float(ws.cell(r, 3).value or 0),
                              value_per_sf=float(ws.cell(r, 4).value or 0))
        for year, col in years.items():
            v = ws.cell(r, col).value
            if v:
                c.sf_sold[year] = float(v)
                c.sf_delivered[year] = float(v)
        if c.total_sf:
            out.append(c)
    return out


# ── Writing the template ──────────────────────────────────────────────────────

_TEMPLATE_SECTIONS: list[tuple[str, list[tuple[str, str, Optional[str]]]]] = [
    ("District", [
        ("Public Infrastructure District Name", "PID", None),
        ("City", "CITY", None),
        ("County", "COUNTY", None),
        ("Developer", "DEVELOPER", None),
        ("Scenario label", "TITLE4", None),
    ]),
    ("Structure", [
        ("Delivery date", "DELIVERY", FMT_DATE),
        ("Principal maturity month (3 = March)", "PRIN_MATURITY", FMT_INT),
        ("Senior principal pay day", "PRIN_MATURITY_DAY_SENIOR", FMT_INT),
        ("Subordinate principal pay day", "PRIN_MATURITY_DAY_SUB", FMT_INT),
        ("Final maturity - senior (years)", "FINAL_MAT_YRS", FMT_INT),
        ("Final maturity - subordinate (years)", "FINAL_MAT_SUB_YRS", FMT_INT),
        ("Premium call (years)", "PREMIUM_CALL_YEARS", FMT_INT),
        ("Premium call price", "PREMIUM_CALL_FIRST_PRICE", None),
        ("Par call (years)", "PAR_CALL_YEARS", FMT_INT),
        ("Capitalized interest", "CAPI", None),
        ("Capitalized interest period (months)", "CAPI_TERM", FMT_INT),
        ("Surplus fund toggle", "SURPLUS_ON_OFF", None),
        ("Release surplus funds?", "SURPLUS_RELEASE_SIZING", None),
        ("Subordinate par override (blank = size it)", "SUB_PAR_OVERRIDE", FMT_MONEY),
        ("Refunding", "REFUND_FINANCING", None),
    ]),
    ("Pricing", [
        ("Investment grade rating", "IG_RATED", None),
        ("Senior rate - investment grade", "SENIOR_RATE_IG", FMT_PCT),
        ("Senior rate - non-rated", "SENIOR_RATE_NR", FMT_PCT),
        ("Subordinate rate - investment grade", "SUB_RATE_IG", FMT_PCT),
        ("Subordinate rate - non-rated", "SUB_RATE_NR", FMT_PCT),
        ("Senior coverage requirement", "DSC_SENIOR_LIEN_BONDS", None),
        ("Subordinate coverage requirement", "DSC_SUB_LIEN_BONDS", None),
    ]),
    ("Fees", [
        ("Costs of issuance", "COI", FMT_MONEY),
        ("Underwriters' discount - senior", "UWD_SENIOR", FMT_PCT),
        ("Underwriters' discount - subordinate", "UWD_SUB", FMT_PCT),
        ("County collection cost", "COUNTY_COLLECTION_FEE", FMT_PCT),
        ("Trustee fee - senior", "TRUSTEE_FEE", FMT_MONEY),
        ("Trustee fee - subordinate", "TRUSTEE_FEE_SUB", FMT_MONEY),
        ("Annual district administration", "ADMIN_COST_BASE", FMT_MONEY),
        ("Administration cost growth", "ADMIN_COST_GROWTH", FMT_PCT),
        ("First year district costs are charged", "DISTRICT_COST_START_YEAR", FMT_INT),
    ]),
    ("Utah property tax", [
        ("Debt service mill levy (governing document)", "MILL_LEVY_GOVERNING_DOC", '0.000'),
        ("Primary residential taxable ratio (1 - 45% exemption)", "RESID_TAXABLE_RATIO", FMT_PCT),
        ("Developed lot taxable ratio", "DEVELOPED_LOT_VALUE", FMT_PCT),
        ("Finished lot value (% of ASP)", "PLATTED_LOT_VALUE", FMT_PCT),
        ("Property tax collection rate", "TAX_COLLECT_MILL_PRC", FMT_PCT),
        ("Personal property uniform fee %", "UNIFORM_FEE_PRC", FMT_PCT),
        ("Reassessment frequency (Annual | Biennial)", "REASSESS_FREQUENCY", None),
        ("Reassessment rate - existing homes", "REASSESS_RATE", FMT_PCT),
        ("Reassessment rate - new homes", "REASSESS_RATE_SUBORDINATE", FMT_PCT),
        ("Value lag (years)", "VALUE_LAG_YEARS", FMT_INT),
        ("State assessed valuation", "STATE_ASSESSED", FMT_MONEY),
        ("Centrally assessed valuation", "CENTRALLY_ASSESSED_VALUE", FMT_MONEY),
    ]),
    ("Development", [
        ("Home price inflation", "INFLATION_RATE", FMT_PCT),
        ("Residential delivery base year", "RESID_DELIVERY_YEAR", FMT_DATE),
        ("First year for Summary", "FIRST_YEAR", '0'),
        ("Development schedule start year", "DEV_START_YEAR", '0'),
        ("Hypothetical sizing scenario", "HYPOTHETICAL_SCENARIO", None),
        ("Lot delivery scenario factor", "LOT_DELIVERY_SCENARIO", FMT_PCT),
        ("Absorption scenario factor", "ABSORPTION_SCENARIO", FMT_PCT),
        ("Interest earnings rate", "INTEREST_EARN_RATE", FMT_PCT),
        ("System development fee (per unit)", "SYSTEM_DEVELOPMENT_FEE", FMT_MONEY),
    ]),
]


def write_template(path: str, cfg: Optional[ModelConfig] = None,
                   dev: Optional[DevelopmentProjections] = None) -> str:
    """Write a pre-filled inputs workbook."""
    from openpyxl import Workbook

    from .development import viridian_farm_projections

    cfg = cfg or ModelConfig()
    dev = dev or viridian_farm_projections()

    wb = Workbook()
    ws = wb.active
    ws.title = "Inputs"
    _widths(ws, {"A": 3, "B": 52, "C": 22, "D": 34})
    _put(ws, 1, 2, "Utah PID Model — Inputs", font=TITLE_FONT)
    _put(ws, 2, 2, "Edit the shaded cells, then run:  python main.py "
                   "--inputs <this file>",
         font=SUB_FONT)
    for col, text in ((2, "Input"), (3, "Value"), (4, "Range Name")):
        _header(ws, 4, col, text)

    r = 5
    for section, items in _TEMPLATE_SECTIONS:
        r += 1
        _put(ws, r, 2, section, font=SECTION_FONT)
        r += 1
        for label, range_name, fmt in items:
            if range_name == DEV_START_RANGE:
                value = dev.start_year
            else:
                attr = RANGE_TO_ATTR[range_name]
                value = getattr(cfg, attr, None)
            _put(ws, r, 2, label)
            _put(ws, r, 3, value, fmt=fmt, fill=INPUT_FILL, align="center")
            _put(ws, r, 4, range_name,
                 font=__import__("openpyxl").styles.Font(name="Consolas", size=9,
                                                         color="808080"))
            r += 1

    _write_dev_template(wb, dev)
    _write_comm_template(wb, dev)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


def _write_dev_template(wb, dev: DevelopmentProjections) -> None:
    ws = wb.create_sheet("Residential Development")
    years = sorted({y for p in dev.products
                    for y in list(p.lot_delivery) + list(p.home_closings)}) or [2025]
    years = list(range(min(years), max(years) + 1))
    _widths(ws, {"A": 3, "B": 30, "C": 10, "D": 10, "E": 14})
    _put(ws, 1, 2, "Residential Development", font=TITLE_FONT)
    _put(ws, 2, 2, "One row per product. Lot delivery above, home closings below.",
         font=SUB_FONT)

    def head(row: int) -> None:
        for col, text in ((2, "Product"), (3, "Type"), (4, "Units"), (5, "Base ASP")):
            _header(ws, row, col, text)
        for i, y in enumerate(years):
            _header(ws, row, 6 + i, y)
            ws.column_dimensions[__import__("openpyxl").utils
                                 .get_column_letter(6 + i)].width = 9

    _put(ws, 4, 2, "LOT DELIVERY", font=SECTION_FONT)
    head(5)
    for i, p in enumerate(dev.products):
        r = 6 + i
        _put(ws, r, 2, p.name)
        _put(ws, r, 3, p.product_type, align="center")
        _put(ws, r, 4, p.total_units, fmt=FMT_INT, fill=INPUT_FILL)
        _put(ws, r, 5, p.asp, fmt=FMT_MONEY, fill=INPUT_FILL)
        for j, y in enumerate(years):
            _put(ws, r, 6 + j, p.lot_delivery.get(y) or None, fmt=FMT_INT,
                 fill=INPUT_FILL, align="center")

    start = 6 + len(dev.products) + 2
    _put(ws, start, 2, "HOME CLOSINGS", font=SECTION_FONT)
    head(start + 1)
    for i, p in enumerate(dev.products):
        r = start + 2 + i
        _put(ws, r, 2, p.name)
        for j, y in enumerate(years):
            _put(ws, r, 6 + j, p.home_closings.get(y) or None, fmt=FMT_INT,
                 fill=INPUT_FILL, align="center")


def _write_comm_template(wb, dev: DevelopmentProjections) -> None:
    ws = wb.create_sheet("Commercial Development")
    _widths(ws, {"A": 3, "B": 30, "C": 14, "D": 14})
    _put(ws, 1, 2, "Commercial Development", font=TITLE_FONT)
    _put(ws, 2, 2, "Optional. One row per commercial parcel; square feet by year.",
         font=SUB_FONT)
    for col, text in ((2, "Parcel"), (3, "Total Sq. Ft."), (4, "Value / Sq. Ft.")):
        _header(ws, 5, col, text)
    years = list(range(2025, 2041))
    for i, y in enumerate(years):
        _header(ws, 5, 6 + i, y)
    for i, c in enumerate(dev.commercial):
        r = 6 + i
        _put(ws, r, 2, c.name)
        _put(ws, r, 3, c.total_sf, fmt=FMT_INT, fill=INPUT_FILL)
        _put(ws, r, 4, c.value_per_sf, fmt=FMT_MONEY, fill=INPUT_FILL)
        for j, y in enumerate(years):
            _put(ws, r, 6 + j, c.sf_sold.get(y) or None, fmt=FMT_INT, fill=INPUT_FILL)
