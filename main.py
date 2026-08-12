#!/usr/bin/env python3
"""
Utah PID financing model — command line entry point.

    python main.py                              run the built-in Viridian Farm deal
    python main.py --inputs inputs.xlsx         run from an inputs workbook
    python main.py --template inputs.xlsx       write a blank inputs workbook
    python main.py --mills 4.0 --coverage 1.25  override single assumptions
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import replace
from datetime import date, datetime

from ut_pid_model import ModelConfig, run_model, viridian_farm_projections
from ut_pid_model.inputs import load_inputs, write_template
from ut_pid_model.memo import build_memo
from ut_pid_model.workbook import build_workbook


def _parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="ut-pid-model",
        description="Utah Public Infrastructure District financing model "
                    "(UCA 17D-4). Produces an Excel workbook and a memo.")
    p.add_argument("--inputs", metavar="FILE",
                   help="inputs workbook (.xlsx) or scenario file (.json)")
    p.add_argument("--template", metavar="FILE",
                   help="write a blank inputs workbook to FILE and exit")
    p.add_argument("--out", default="outputs", metavar="DIR",
                   help="output directory (default: outputs)")
    p.add_argument("--name", metavar="STEM",
                   help="output file stem (default: derived from the district name)")
    p.add_argument("--mills", type=float, help="debt service mill levy override")
    p.add_argument("--coverage", type=float, help="senior coverage requirement override")
    p.add_argument("--senior-rate", type=float, help="senior interest rate, e.g. 0.05875")
    p.add_argument("--sub-rate", type=float, help="subordinate interest rate")
    p.add_argument("--delivery", help="delivery date, YYYY-MM-DD")
    p.add_argument("--sub-par", type=float,
                   help="force a subordinate par instead of sizing it")
    p.add_argument("--absorption", type=float,
                   help="absorption scenario factor, e.g. 0.75 for a 25%% slowdown")
    p.add_argument("--no-memo", action="store_true", help="skip the memo")
    p.add_argument("--no-excel", action="store_true", help="skip the workbook")
    p.add_argument("--quiet", action="store_true")
    return p.parse_args(argv)


def _apply_overrides(cfg: ModelConfig, args: argparse.Namespace) -> ModelConfig:
    changes: dict = {}
    if args.mills is not None:
        changes["mill_levy_governing_doc"] = args.mills
    if args.coverage is not None:
        changes["dsc_senior_lien_bonds"] = args.coverage
    if args.senior_rate is not None:
        changes["senior_rate_nr"] = args.senior_rate
        changes["senior_rate_ig"] = args.senior_rate
    if args.sub_rate is not None:
        changes["sub_rate_nr"] = args.sub_rate
        changes["sub_rate_ig"] = args.sub_rate
    if args.delivery:
        changes["delivery"] = datetime.strptime(args.delivery, "%Y-%m-%d").date()
    if args.sub_par is not None:
        changes["sub_par_override"] = args.sub_par
    if args.absorption is not None:
        changes["hypothetical_scenario"] = "Yes"
        changes["absorption_scenario"] = args.absorption
        changes["lot_delivery_scenario"] = args.absorption
    return replace(cfg, **changes) if changes else cfg


def _stem(cfg: ModelConfig, args: argparse.Namespace) -> str:
    if args.name:
        return args.name
    slug = "".join(ch if ch.isalnum() else "_" for ch in cfg.district_name)
    slug = "_".join(filter(None, slug.split("_")))[:60]
    return f"{date.today():%Y.%m.%d}_{slug}_{cfg.mill_levy_ds_target:.3f}mills"


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv if argv is not None else sys.argv[1:])

    if args.template:
        path = write_template(args.template)
        print(f"Inputs template written to {path}")
        return 0

    if args.inputs:
        cfg, dev = load_inputs(args.inputs)
    else:
        cfg, dev = ModelConfig(), viridian_farm_projections()
    cfg = _apply_overrides(cfg, args)

    res = run_model(cfg, dev)
    os.makedirs(args.out, exist_ok=True)
    stem = _stem(cfg, args)

    written: list[str] = []
    if not args.no_excel:
        written.append(build_workbook(res, os.path.join(args.out, stem + ".xlsx")))
    if not args.no_memo:
        written += build_memo(res, os.path.join(args.out, stem + "_memo"))

    if not args.quiet:
        print(f"{cfg.district_name}")
        print(f"  {cfg.city} City, {cfg.county} County, Utah — "
              f"{cfg.mill_levy_ds_target:.3f} mills for debt service")
        print(f"  Senior lien   {cfg.senior_bonds_series}  "
              f"${res.senior_par:>12,.0f}  at {cfg.senior_interest_rate:.3%}  "
              f"({cfg.dsc_senior_lien_bonds:.2f}x)")
        print(f"  Subordinate   {cfg.sub_bonds_series}  "
              f"${res.sub_par:>12,.0f}  at {cfg.sub_interest_rate:.3%}  "
              f"({cfg.dsc_sub_lien_bonds:.2f}x)")
        print(f"  Reimbursement                ${res.total_reimbursement:>12,.0f}  "
              f"(${res.reimbursement_per_lot:,.0f} per unit across "
              f"{dev.total_units():,} units)")
        print(f"  Repayment ratio {res.repayment_ratio:.2f}x  |  "
              f"All-in TIC {res.senior_stats.all_in_tic:.3%}  |  "
              f"final maturity {res.senior_stats.final_maturity}")
        if not res.converged:
            print("  ! sizing did not fully converge — review the surplus fund inputs")
        for w in res.warnings:
            print(f"  ! {w}")
        print()
        for path in written:
            print(f"  wrote {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
