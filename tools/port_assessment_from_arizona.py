#!/usr/bin/env python3
"""
port_assessment_from_arizona.py
Generate ut_pid_model/assessment_bonds.py from the Arizona CFD model.

The special-assessment engine is the same arithmetic in both states — a lien
per lot, level installments, a payoff at conversion, a value-to-lien test — so
Utah does not reimplement it.  This script ports Arizona's
``cfd_model/assessment_bonds.py`` the same way ``port_from_colorado.py`` ports
the ad valorem model: identifier renames, then label substitutions, then
ordered patches.  Nothing is written unless every patch applies, so when
Arizona moves, the port fails loudly instead of producing a Utah model that
quietly still says "A.R.S. § 48-721".

    python tools/port_assessment_from_arizona.py --source ../AZ-CFD-Model
    python tools/port_assessment_from_arizona.py --source ../AZ-CFD-Model --check

What Utah changes, and why:

* **Statutory framing.**  Arizona assesses under A.R.S. § 48-721; Utah assesses
  under the Assessment Area Act (UCA Title 11, Chapter 42), which a public
  infrastructure district may use alongside its ad valorem levy.
* **The development feed.**  Arizona reads an assessed-valuation model's
  ``sections``; Utah reads ``DeveloperProjections.products`` directly — the same
  lots, lot values and conversions, one layer closer to the source.
* **Costs of issuance.**  Arizona charges a municipal advisor fee, a bond
  discount and a COI percentage, all percentages of par.  Utah's Inputs page
  carries an underwriter's discount percentage and a DOLLAR costs-of-issuance
  figure, so the dollar COI is expressed as its share of the senior par and
  applied to each series at that rate.
* **The reserve.**  Utah's Inputs page funds the assessment DSRF as a
  percentage (DSRF_SA_FUND_PRC) of the three-prong amount, and switches it off
  entirely with DSRF_SA = No.
"""
from __future__ import annotations

import argparse
import difflib
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent

#: The one Arizona module Utah ports wholesale.
SOURCE_MODULE = "assessment_bonds.py"


# ── Pass 1: identifier renames (whole word, longest first) ────────────────────

IDENTIFIER_RENAMES = [
    ("AssessedValuationModel", "SummaryModel"),
    ("assessed_valuation", "summary"),
    ("CFDConfig", "ModelConfig"),
    ("cfd_model", "ut_pid_model"),
    ("av_model", "summary_model"),
    ("avm", "summary_model"),
    # Utah's Inputs page names the senior principal pay day; Arizona has one day
    # for the whole structure.
    ("prin_day", "prin_maturity_day_senior"),
]


# ── Pass 2: label substitutions (plain text, anywhere) ────────────────────────

LABEL_SUBS = [
    ("Arizona CFDs", "Utah public infrastructure districts"),
    ("A.R.S. § 48-721 et seq.", "UCA Title 11, Chapter 42 — the Assessment Area Act"),
    ("30 for a SAID, 25 for a CFD", "the Final Maturity - Senior (yrs) cell"),
    ("the $5.00 debt-service tax rate", "the district's mill levy"),
    ("GO bonds", "mill levy bonds"),
    ("GO bond", "mill levy bond"),
    ("GO series", "senior lien series"),
    ("the GO authorization", "the § 17D-4-303 fifteen-mill cap"),
]


def _p(old: str, new: str) -> None:
    PATCHES.append((old, new))


PATCHES: list[tuple[str, str]] = []


# ── The module docstring states Utah's framework ─────────────────────────────

_p('''"""
Special assessment bond sizing (Inputs!ASSESSMENT_BONDS = Yes).

Utah public infrastructure districts may levy special assessments on the lots within the district and
issue assessment bonds secured by those liens (UCA Title 11, Chapter 42 — the Assessment Area Act), in
addition to — and separate from — the ad valorem GO program sized by
``BondSizingAnalysis``.''',
   '''"""
Special assessment bond sizing (Inputs!ASSESSMENT_BONDS = Yes).

A Utah public infrastructure district may levy assessments on the benefited
lots under the Assessment Area Act (UCA Title 11, Chapter 42) and issue
assessment bonds secured by those liens — a second, separate security from the
ad valorem mill levy that the senior and subordinate liens are sized against.
The two programs do not compete: an assessment is a lien on the lot, not a
charge against the levy, so assessment par does not consume room under the
§ 17D-4-303 fifteen-mill cap and assessment installments are not netted from
the revenue available to the mill levy bonds.''')

_p('''Assessment bonds are secured by the liens, not the District's taxing power, so
their par does NOT count against the § 17D-4-303 fifteen-mill cap or the direct-debt ratio
and their installments are outside the district's mill levy.
"""''',
   '''Assessment bonds are secured by the liens, not the District's taxing power, so
their par does NOT count against the § 17D-4-303 fifteen-mill cap and their
installments sit outside the mill levy entirely.  What the homeowner pays is
the sum of the two: the mill levy on the taxable value of the home plus the
annual installment on the assessment still attached to the lot.
"""''')


# ── Imports: Utah has no GO date helpers to borrow, so carry them here ───────
# Arizona's `first_interest_date` / `first_principal_date` live in its
# debt_service module.  Utah's debt_service builds whole schedules rather than
# single dates, so the two helpers come across with the engine that uses them
# rather than being bolted onto a module that has no other use for them.

_p('''from .config import ModelConfig
from .summary import SummaryModel
from .debt_service import first_interest_date as _go_first_interest, first_principal_date as _go_first_principal
''',
   '''from .config import ModelConfig
from .summary import SummaryModel


def _md(year: int, month: int, day: int = 1) -> date:
    """A month/day in a year, clamped to the length of the month."""
    return date(year, month, min(day, calendar.monthrange(year, month)[1]))


def _add_months(d: date, months: int) -> date:
    return _md(*_ym_plus(d.year, d.month, months), d.day)


def _ym_plus(year: int, month: int, months: int) -> tuple[int, int]:
    total = (year * 12 + (month - 1)) + months
    return total // 12, total % 12 + 1


def _go_first_interest(dated: date, prin_month: int, int_month: int, day: int = 1) -> date:
    """First coupon date strictly after the dated date."""
    lo = _md(dated.year, min(prin_month, int_month), day)
    hi = _md(dated.year, max(prin_month, int_month), day)
    if lo <= dated:
        return hi if hi > dated else _md(dated.year + 1, min(prin_month, int_month), day)
    return lo


def _go_first_principal(dated: date, capi_years: float, prin_month: int,
                        int_month: int, first_int: date, day: int = 1) -> date:
    """
    First principal date — the first principal date strictly AFTER the
    capitalized-interest period (dated + ``capi_years``).

    Nothing amortizes while interest is being capitalized, which is the same
    rule the mill levy bonds follow: a series dated 9/30/2026 with three CAPI
    years capitalizes every coupon through 9/30/2029 and pays its first
    principal in 2030.
    """
    end_capi = _add_months(dated, int(round(12 * capi_years)))
    if end_capi >= _md(end_capi.year, prin_month, day):
        candidate = _md(end_capi.year + 1, prin_month, day)
    else:
        candidate = _md(end_capi.year, prin_month, day)
    return max(candidate, first_int)
''')


# ── Costs of issuance: Utah's Inputs page, not Arizona's ────────────────────
# Arizona charges three percentages of par (municipal advisor fee, bond
# discount, COI %).  Utah's Inputs page has an underwriter's discount PERCENT
# and a costs-of-issuance DOLLAR amount, so the series carries the underwriter's
# discount at its entered rate and a share of the dollar COI at the rate that
# figure represents on the senior par — "the same assumptions as the mill levy
# bonds", expressed in the units Utah enters them in.

_p('''    muniadv_fee: float               # same COI assumptions as the mill levy bonds (Inputs page)
    bond_discount: float
    coi_pct_fee: float''',
   '''    muniadv_fee: float               # unused in Utah (no municipal advisor fee cell); kept at 0.0
    bond_discount: float             # UWD_SUB x par — the underwriter's discount
    coi_pct_fee: float               # COI as a share of senior par, applied to this series''')

_p('''                muniadv_fee=par * cfg.muniadv_fee, bond_discount=par * cfg.bond_discount,
                coi_pct_fee=par * cfg.coi_pct, reserve_fund=_dsrf["reserve"] * _hair,''',
   '''                muniadv_fee=0.0, bond_discount=par * _uwd_rate(cfg),
                coi_pct_fee=par * _coi_rate(cfg), reserve_fund=_dsrf["reserve"] * _hair,''')

_p('''def _ratio(a: float, b: float, default: float) -> float:''',
   '''def _uwd_rate(cfg) -> float:
    """
    Underwriter's discount on an assessment series, as a share of par.

    Assessment bonds are sold on the lots' security rather than the district's
    levy, so they take the SUBORDINATE lien's discount (UWD_SUB), not the
    senior's — the senior rate understates what a lien-secured developer bond
    costs to place.
    """
    return float(getattr(cfg, "uwd_sub", 0.0) or 0.0)


def _coi_rate(cfg) -> float:
    """
    Costs of issuance on an assessment series, as a share of par.

    Utah enters COI as a DOLLAR amount for the mill levy financing, which cannot
    be charged whole to each assessment series without multiplying it by the
    number of series.  "The same assumptions as the mill levy bonds" therefore
    means the same RATE: the dollar COI as a share of the senior par it was
    quoted against.  With no senior par to measure against, nothing is charged.
    """
    coi = float(getattr(cfg, "coi", 0.0) or 0.0)
    basis = float(getattr(cfg, "_senior_par_for_coi", 0.0) or 0.0)
    return (coi / basis) if (coi > 0 and basis > 0) else 0.0


def _ratio(a: float, b: float, default: float) -> float:''')


# ── The development feed: Utah reads the products, not an AV model's sections ─
# Arizona's assessed-valuation model exposes `sections` keyed by roll date with
# per-product lot units, lot value and closings.  Utah's DeveloperProjections
# already holds exactly that, per product, keyed by year — so read it straight
# rather than build an intermediate.

_p('''        try:
            summary_model._ensure_built()
            products = list(getattr(summary_model.dev, "products", []) or [])
        except Exception:
            products = []''',
   '''        try:
            products = [p.name for p in getattr(summary_model.dev, "products", []) or []]
        except Exception:
            products = []''')

_p('''    def run(self) -> "AssessmentBondAnalysis":
        cfg = self.cfg
        self.summary_model._ensure_built()
        dev = self.summary_model.dev
        sections = getattr(dev, "sections", {}) or {}
        lots_by_year: dict[int, float] = {}
        lotval_by_year: dict[int, float] = {}
        homes_by_year: dict[int, float] = {}
        lots_p: dict[int, dict] = {}
        lotval_p: dict[int, dict] = {}
        homes_p: dict[int, dict] = {}
        for d in sorted(sections):
            s = sections[d]
            y = d.year
            for p, n in s["lot_units"].items():
                if n <= 0:
                    continue
                lots_by_year[y] = lots_by_year.get(y, 0.0) + n
                lotval_by_year[y] = lotval_by_year.get(y, 0.0) + float(s["lot_value"].get(p, 0.0))
                lots_p.setdefault(y, {})[p] = lots_p.get(y, {}).get(p, 0.0) + n
                lotval_p.setdefault(y, {})[p] = lotval_p.get(y, {}).get(p, 0.0) + float(s["lot_value"].get(p, 0.0))
            for p, n in s["closing_units"].items():
                if n > 0:
                    homes_by_year[y] = homes_by_year.get(y, 0.0) + n
                    homes_p.setdefault(y, {})[p] = homes_p.get(y, {}).get(p, 0.0) + n''',
   '''    def run(self) -> "AssessmentBondAnalysis":
        cfg = self.cfg
        dev = self.summary_model.dev
        (lots_by_year, lotval_by_year, homes_by_year,
         lots_p, lotval_p, homes_p) = _development_feed(dev, cfg)''')

_p('''def _int_date(year: int, cfg: ModelConfig) -> date:''',
   '''def _development_feed(dev, cfg):
    """
    Lots delivered, their finished-lot value and homes closed — by year, and by
    year and product — read straight off ``DeveloperProjections``.

    The lot value is the same figure the ad valorem model puts on a delivered
    lot: the product's ASP inflated to the delivery year, times the platted-lot
    percentage.  It is a MARKET value, which is what a value-to-lien test wants;
    the 45% residential exemption reduces what the lot is TAXED on, not what it
    is worth to a lienholder.
    """
    lots_by_year: dict[int, float] = {}
    lotval_by_year: dict[int, float] = {}
    homes_by_year: dict[int, float] = {}
    lots_p: dict[int, dict] = {}
    lotval_p: dict[int, dict] = {}
    homes_p: dict[int, dict] = {}
    infl = float(getattr(cfg, "inflation_rate", 0.0) or 0.0)
    infl_start = int(getattr(cfg, "inflation_start_year", 0) or 0)
    platted = float(getattr(cfg, "platted_lot_value", 0.10) or 0.10)
    for prod in (getattr(dev, "products", []) or []):
        for y, n in (prod.lot_deliveries or {}).items():
            if not n:
                continue
            y = int(y)
            val = float(n) * prod.lot_value(y, infl, platted, infl_start)
            lots_by_year[y] = lots_by_year.get(y, 0.0) + float(n)
            lotval_by_year[y] = lotval_by_year.get(y, 0.0) + val
            by_u = lots_p.setdefault(y, {})
            by_v = lotval_p.setdefault(y, {})
            by_u[prod.name] = by_u.get(prod.name, 0.0) + float(n)
            by_v[prod.name] = by_v.get(prod.name, 0.0) + val
        for y, n in (prod.home_closings or {}).items():
            if not n:
                continue
            y = int(y)
            homes_by_year[y] = homes_by_year.get(y, 0.0) + float(n)
            by_h = homes_p.setdefault(y, {})
            by_h[prod.name] = by_h.get(prod.name, 0.0) + float(n)
    return lots_by_year, lotval_by_year, homes_by_year, lots_p, lotval_p, homes_p


def _int_date(year: int, cfg: ModelConfig) -> date:''')

_p('''        dev = self.summary_model.dev
        prices = getattr(dev, "home_prices", {}) or {}''',
   '''        dev = self.summary_model.dev
        prices = {p.name: p.asp_base for p in (getattr(dev, "products", []) or [])}''')

# Arizona walks the AV model's roll sections for closings and closing value;
# Utah reads the products' own closing schedule, and values a closing at the
# product's ASP in the year it closes.
_p('''        units = {}
        values = {}          # Σ closing value by product — fallback home value when no base ASP
        for sec in (getattr(dev, "sections", {}) or {}).values():
            for prod, n in sec["closing_units"].items():
                units[prod] = units.get(prod, 0.0) + n
                values[prod] = values.get(prod, 0.0) + float(sec.get("closing_value", {}).get(prod, 0.0) or 0.0)
        rows = []
        w_units = w_asp = 0.0
        for prod in getattr(dev, "products", []) or list(set(prices) | set(units)):''',
   '''        units = {}
        values = {}          # Σ closing value by product — fallback home value when no base ASP
        _infl = float(getattr(cfg, "inflation_rate", 0.0) or 0.0)
        _infl_start = int(getattr(cfg, "inflation_start_year", 0) or 0)
        for p in (getattr(dev, "products", []) or []):
            for y, n in (p.home_closings or {}).items():
                if not n:
                    continue
                units[p.name] = units.get(p.name, 0.0) + float(n)
                values[p.name] = values.get(p.name, 0.0) + float(n) * p.asp(
                    int(y), _infl, _infl_start)
        rows = []
        w_units = w_asp = 0.0
        for prod in [p.name for p in (getattr(dev, "products", []) or [])] or list(set(prices) | set(units)):''')


# ── The assessment term follows Utah's Final Maturity cell ──────────────────

_p('''Level annual installments at the Special Assessment Interest Rate
(Inputs!BONDRATE_SPECIAL) over the program's Final Maturity (FINAL_MAT_YRS —
the Final Maturity - Senior (yrs) cell), beginning the year after each series is dated;''',
   '''Level annual installments at the Special Assessment Interest Rate
(Inputs!BONDRATE_SPECIAL) over the program's Final Maturity (Inputs!FINAL_MAT_YRS,
the same term the senior lien runs), beginning the year after each series is dated;''')


# ── Port machinery ───────────────────────────────────────────────────────────

def rename_identifiers(text: str) -> str:
    for old, new in IDENTIFIER_RENAMES:
        text = re.sub(rf"\b{re.escape(old)}\b", new, text)
    return text


def substitute_labels(text: str) -> str:
    for old, new in LABEL_SUBS:
        text = text.replace(old, new)
    return text


def apply_patches(text: str, failures: list[str]) -> str:
    for i, (old, new) in enumerate(PATCHES):
        if old not in text:
            failures.append(f"patch #{i} did not match:\n    "
                            + old.strip().splitlines()[0][:100])
            continue
        text = text.replace(old, new, 1)
    return text


def port(src: Path, failures: list[str]) -> str:
    text = src.read_text()
    text = rename_identifiers(text)
    text = substitute_labels(text)
    return apply_patches(text, failures)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--source", required=True,
                    help="path to the AZ-CFD-Model checkout")
    ap.add_argument("--check", action="store_true",
                    help="report differences without writing")
    args = ap.parse_args(argv)

    src = Path(args.source).expanduser().resolve() / "cfd_model" / SOURCE_MODULE
    if not src.is_file():
        print(f"error: {src} not found", file=sys.stderr)
        return 2

    failures: list[str] = []
    ported = port(src, failures)
    if failures:
        print("PATCHES THAT DID NOT APPLY — Arizona has moved, re-read these:",
              file=sys.stderr)
        for f in failures:
            print(f"  • {f}", file=sys.stderr)
        print("\nNothing was written.", file=sys.stderr)
        return 1

    dst = REPO / "ut_pid_model" / SOURCE_MODULE
    if dst.exists() and dst.read_text() == ported:
        print("already up to date with Arizona")
        return 0
    if args.check:
        before = dst.read_text().splitlines() if dst.exists() else []
        print("\n".join(list(difflib.unified_diff(
            before, ported.splitlines(), fromfile=f"a/{SOURCE_MODULE}",
            tofile=f"b/{SOURCE_MODULE}", lineterm="", n=2))[:80]))
        print(f"would update ut_pid_model/{SOURCE_MODULE}")
        return 0
    dst.write_text(ported)
    print(f"updated ut_pid_model/{SOURCE_MODULE}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
