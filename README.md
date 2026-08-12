# Utah PID Model

A financing model for Utah Public Infrastructure Districts, built on the Tierra
Financial Advisors Colorado metropolitan district template. Same tabs, same
layout, same sizing mechanics — with Utah's property tax framework substituted
for Colorado's.

It produces two deliverables from one run:

- **an Excel workbook** with the Colorado model's sixteen tabs, in order; and
- **a financing memorandum** in Markdown and Word.

## Quick start

```bash
pip install -r requirements.txt

python main.py                                 # the built-in Viridian Farm deal
python main.py --template inputs.xlsx          # write a blank inputs workbook
python main.py --inputs inputs.xlsx            # run from that workbook
python main.py --mills 4.0 --coverage 1.25     # override single assumptions
python main.py --absorption 0.75               # 25% absorption slowdown
```

Output lands in `outputs/`:

```
2026.08.12_Viridian_Farm_..._3.000mills.xlsx        the workbook
2026.08.12_Viridian_Farm_..._3.000mills_memo.md     the memo (Markdown)
2026.08.12_Viridian_Farm_..._3.000mills_memo.docx   the memo (Word)
```

## Workbook tabs

Identical set and order to the Colorado model:

| # | Tab | What it holds |
|---|---|---|
| 1 | Inputs - First | Every assumption, at the template's row numbers, with named ranges |
| 2 | Capital Costs | Eligible cost build-up to the bond par amount |
| 3 | Costs of Issuance | Line-item cost of issuance |
| 4 | Sources and Uses - First | Sources, uses, key assumptions, bond statistics, taxing authority |
| 5 | Sources and Uses - Refunding | The refunding series (laid out even when toggled off) |
| 6 | Summary | Taxable value → pledged revenue → debt service → surplus → subordinate waterfall |
| 7 | Residential Development | Lot delivery, absorption, pricing and taxable value creation, by product |
| 8 | Comm Development | Commercial square footage and value |
| 9 | Ops Rev & Exp Projection | Operations levy projection |
| 10 | Senior Lien DS - First | Senior sizing and debt service, semi-annual |
| 11 | Sub Lien DS - First (Annual) | Subordinate schedule, annual |
| 12 | Sub Lien DS - First (SA) | Subordinate schedule, semi-annual |
| 13 | Senior Lien DS - Refunding | Refunding debt service |
| 14 | CAPI Fund - First | Capitalized interest fund draw-down |
| 15 | Scratch--->>> | Working divider |
| 16 | DBC Output | Flat principal / debt service export |

Named ranges are written on `Inputs - First` at the template's cells. Where Utah
re-labels a Colorado concept, both names point at the same cell
(`RESID_TAXABLE_RATIO` and `TABOR_CURRENT`, `MILL_LEVY_GOVERNING_DOC` and
`MILL_LEVY_SERVICE_PLAN`, and so on) so the two workbooks stay diff-able.

## Utah vs. Colorado

| Item | Colorado metro district | Utah PID |
|---|---|---|
| Enabling act | Title 32, Art. 1, C.R.S. | Title 17D, Ch. 4, Utah Code |
| Levy cap | Service plan cap, gallagherized | Least of **15.000 mills** (§ 17D-4-303), the governing document cap, and the indenture cap — no adjustment mechanism |
| Residential assessment | Gallagher rate (6.7% in the reference model) | **55% of fair market value** — the 45% primary residential exemption, § 59-2-103 |
| Developer lot inventory | 29% vacant-land rate | 55% — the exemption reaches unoccupied property the assessor finds will become a primary residence (Utah Admin. Code R884-24P-52) |
| Reassessment | Biennial | **Annual** (§ 59-2-303.1) → a one-year lag from value creation to the tax roll, against Colorado's two |
| Tax due date | Half 28 Feb / half 15 Jun | **Single payment, 30 November** |
| Principal payment | 1 December | **1 March** |
| Vehicle tax | Specific ownership tax, ~6–8% of the levy | Personal property uniform fee (§ 59-2-405), distributed pro rata |
| County collection cost | Treasurer's fee, ~1.5% of the distribution | Recovered by a separate statewide levy (§ 59-2-1602) — no haircut on the district |
| Raising the rate | TABOR election | Truth in Taxation hearing above the certified tax rate; the required mill levy is exempt while within the caps |
| Agricultural land | Agricultural classification | Greenbelt Reduction (§ 59-2-503) with up to five years of rollback tax on withdrawal |
| Operations | Separate O&M mill levy | Typically none — administration is charged against pledged revenue |

`docs/utah-vs-colorado.md` walks through each of these and where it lands in the
code.

## How the model runs

```
Residential / Comm Development  →  Summary taxable value
                                →  Summary pledged revenue
                                →  Senior Lien DS sizing
                                →  Surplus fund / CAPI
                                →  Sub Lien DS sizing
                                →  Sources and Uses, Ops, DBC Output
```

**Taxable value.** Finished lots are carried at a share of the eventual home
price until a home closes on them; closed homes enter at their inflated sale
price. Both are taxed at 55% of fair market value, one year after the value is
created. Values already on the roll grow at the reassessment rate.

**Senior sizing.** Principal is solved backwards from the final maturity. At
each principal date the year's net pledged revenue is divided by the coverage
requirement, the coupon that later maturities already generate is removed, and
the balance becomes principal in $5,000 denominations.

**The circular bit.** The surplus fund deposit depends on debt service, which
depends on the surplus fund. The workbook resolves this with Excel's iterative
calculation; here it is an explicit fixed-point loop that reports whether it
converged (`Results.converged`) rather than leaving a half-settled spreadsheet.

**Subordinate sizing.** The subordinate bonds are cashflow bonds: they take
whatever the senior lien releases, unpaid interest accrues, and principal is
retired as headroom allows. The model sizes the largest par the residual
cashflow retires in full — principal *and* accrued interest — by the final
maturity. Set `sub_par_override` (or `--sub-par`) to force a round number
instead.

## Tie-out

The Python model reproduces the reference workbook in this repository
(`Financial Analysis - Viridian Farms PID (3 MILLS) - Salem_Pricing Day
(Sept 17 2024).xlsm`) exactly on taxable value and pledged revenue, and to the
dollar on senior par:

| | Model | Reference workbook | Priced deal |
|---|---|---|---|
| Senior par | $5,690,000 | $5,690,000 | $5,645,000 |
| Total taxable value, 2030 | $195,362,877 | $195,362,877 | — |
| Net pledged revenue, 2030 | $512,933 | $512,933 | — |
| Final maturity | 3/1/2054 | 3/1/2054 | 3/1/2054 |

Those figures are pinned in `tests/test_model.py`, so a change that moves them
fails the suite.

Two places the model deliberately departs from the reference workbook, both
documented in the memo:

1. **Subordinate par is sized, not hand-entered.** The workbook carries a typed
   $1,000,000, which leaves accrued interest outstanding at maturity. The model
   sizes $801,000, the largest amount fully retired.
2. **The surplus fund is solved to convergence.** The reference workbook's own
   cells disagree with each other ($545,055 on the Inputs tab, $393,950 on
   Sources and Uses) because its iterative calculation had not settled.

## Layout

```
main.py                     CLI
ut_pid_model/
  config.py                 every input, at the template's row numbers
  development.py            absorption schedules and the value blocks
  engine.py                 taxable value, revenue, sizing, waterfalls
  workbook.py               the Excel writer
  memo.py                   the memorandum
  inputs.py                 read an inputs workbook; write the blank template
  xlfin.py                  EDATE, YEARFRAC, DAYS360, PRICE, TIC solver
tests/test_model.py         regression tests against the reference workbook
docs/utah-vs-colorado.md    the statutory walk-through
```

## Reference documents

- `UTViridianFarm01a-FIN.pdf` — Viridian Farm PID No. 1 limited offering
  memorandum, September 2024
- `Financial Analysis - Viridian Farms PID (3 MILLS) - Salem_Pricing Day
  (Sept 17 2024).xlsm` — the pricing-day workbook

## Tests

```bash
python -m pytest tests -q
```
