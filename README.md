# Utah Public Infrastructure District Financial Model

The Tierra Colorado metro-district model (`co_metro_model`), rebuilt for Utah
public infrastructure districts organised under the **Public Infrastructure
District Act, Title 17D, Chapter 4, Utah Code**. Same module layout, same
engines, same deliverables — with Utah's property tax framework in place of
Colorado's.

The model sizes a layered **senior / subordinate** financing, computes the
developer **reimbursement**, and **refinances the senior new-money bonds** once
they are callable to generate additional reimbursement ("new money").

Validated against the **Viridian Farm Public Infrastructure District No. 1**
financing (Salem City, Utah County) priced 17 September 2024 — the limited
offering memorandum and pricing-day workbook are both in this repository.

## Why Utah is different

| Mechanic | Colorado (`co_metro_model`) | Utah (this model) |
|---|---|---|
| Enabling act | Title 32, Art. 1, C.R.S. | Title 17D, Ch. 4, Utah Code |
| Residential taxation | Market value × TABOR ratio (~6.7%) | **55% of fair market value** — the 45% primary residential exemption (§ 59-2-103) |
| Builder lot inventory | 29% vacant-land rate, phasing to 25% under SB24-233 | **55%** — the exemption reaches unoccupied property the assessor finds will be a primary residence (Utah Admin. Code R884-24P-52) |
| Reassessment | Biennial, odd-year reappraisal, value held flat between | **Annual** (§ 59-2-303.1) — the base steps up every year |
| Levy cap | Service Plan cap, "Gallagherized" as the residential ratio falls | Least of **15.000 mills** (§ 17D-4-303), the governing document, and the indentures — fixed rates, no adjustment mechanism |
| Rate increases | TABOR election | Truth in Taxation hearing above the certified tax rate; the required mill levy is **exempt** while within the caps |
| Taxes due | Half 28 Feb / half 15 Jun | **Single payment, 30 November** |
| Principal date | 1 December | **1 March** |
| Vehicle revenue | Specific ownership tax (~6–8% of the levy) | Personal property **uniform fee** (§ 59-2-405), distributed pro rata |
| County collection | Treasurer's fee off the top (~1.5%) | Separate statewide levy on property (§ 59-2-1602) — **no haircut** |
| Agricultural land | Agricultural classification | **Greenbelt Reduction** (§ 59-2-503), up to five years of rollback tax on withdrawal |

The headline consequence: because Utah taxes homes on 55% of value rather than
6.7%, a **3-mill Utah levy raises more than a 60-mill Colorado levy** on the
same houses. `docs/utah-vs-colorado.md` walks through each difference and points
at where it lands in the code.

### Utah taxable-value lag

Value is set as of **1 January**, appears on that year's roll, and the taxes are
due **30 November of the same year** — funding the following **1 March** debt
service payment. So value created during calendar year *V* lands on the roll for
*V+1*, is collected in *V+1*, and pays debt service in *V+2*: a **two-year lag**
(`ModelConfig.av_lag_years = 2`), applied in `summary.py`.

Colorado reaches the same two-year lag by a different route — a June-30 level of
value the year before a biennial reappraisal, collected the year *after* the
roll. Same number of years; different mechanism, and different behaviour in
between, because Utah has no two-year hold.

### Levy caps

Three caps bind a Utah PID; the most restrictive controls, and `mill_levy_cap`
returns it:

| Source | Viridian Farm |
|---|---|
| § 17D-4-303 — 0.015 per dollar of taxable value | 15.000 mills |
| Governing document | 5.000 mills |
| Indentures | **3.000 mills** ← controls |

None of them float with the residential exemption, so there is no
"Gallagherization" step. `ModelConfig.validate()` flags a levy above the
statutory cap and a target above the controlling cap.

### Subordinate lien — fully modeled cash-flow structure

Unchanged from Colorado (`subordinate.py`): the senior surplus / debt-service
reserve fund accumulates the senior residual toward a target (½ × max senior
annual debt service), spills above-target cash to the sub lien, and releases the
reserve at senior maturity; the subordinate cash-flow bond accrues interest that
compounds while unpaid, and is repaid interest-first then principal. The one
Utah change is what feeds it — pledged revenue net of district O&M,
which a Utah PID pays out of the debt service levy rather than a separate
operations levy.

## Staying in sync with Colorado

This model is a **port** of `co_metro_model`, not a fork. `tools/port_from_colorado.py`
regenerates the package from a Colorado checkout in one command:

```bash
python tools/port_from_colorado.py --source ../CO-Metro-District-Model --check   # dry run
python tools/port_from_colorado.py --source ../CO-Metro-District-Model           # apply
python -m pytest tests -q                                                        # re-verify
```

The port runs three passes — identifier renames, display-label substitutions,
and semantic patches where Utah law changes the arithmetic or the prose. **Every
patch is checked**: if Colorado refactors the code a patch anchors to, the port
fails and writes nothing, rather than quietly emitting a Colorado-flavoured Utah
model. `tests/test_model.py::test_port_is_up_to_date_with_colorado` fails when
the Colorado checkout has moved ahead.

**Currently synced to Colorado `a5d8c31` plus a cherry-pick of `16b4edd`**
(the Debt Structure rebuild). Colorado has commits between and after those two —
a development volume stress, costs-of-issuance proration, a memo absorption
reorder, deliverable-filename fixes and a call-date alignment — that are
deliberately *not* ported yet. The port source is reproduced with:

```bash
git -C ../CO-Metro-District-Model checkout -B ut-debt-structure a5d8c31
git -C ../CO-Metro-District-Model cherry-pick 16b4edd     # .py files merge clean
```

Colorado's optional
**Series C** third lien ports across as code but is carried **inert**: the
`SIZE_SERIES_C` toggle stays `"No"` and its rows are kept off the Utah Inputs
template. Series C is sized against a *separate* assessment that reassesses the
created value at its own rate; whether a Utah PID may levy against a second
assessment on the same property is a statutory question this model has not
worked through, and an untested toggle in a client-facing template is worse than
no toggle. Turning it on is a one-line change once that authority is settled.

Eight upstream fixes ride in the patch set and are worth pushing back to
`co_metro_model` (an eighth — flooring the solved subordinate par instead of
rounding it — was adopted upstream in `ba72e6b`, so the patch is retired):

* **The capitalized-interest end date snaps to the principal month only.**
  `_snap_to_payment_date` can never land on the other coupon, so 36 months from
  9/30/2026 — which is 9/30/2029 — was reported as 3/1/2029, i.e. 29 months.
  The port snaps to the most recent *coupon* date at the bond's frequency
  (9/1/2029 semiannual, 3/1/2029 annual) and decides capitalization by date
  rather than by year.
* **Interest is charged as a full half-year, including the first coupon.**
  Interest accrues from the DATED date, so a bond not dated on a coupon date
  owes a stub first coupon — the reference workbook's first senior coupon is
  $143,929.34 against $167,143.75 thereafter, 155/180 of a period. Upstream
  charged the full amount.
* **A year's two payments are emitted mid-year coupon first, then the principal
  date.** That is date order in Colorado (June, then December) but reversed in
  Utah, where principal falls in March and the other coupon in September — so
  the September coupon was charged on a balance March had already paid down.
  The port builds the dates in order and carries a running balance through.
* **The revenue wrap sizes at the flat rate, not the entered coupons.**
  `size_for_par` charges `balance * rate` while `_apply_coupon_scale` afterwards
  restates the real interest at the per-maturity coupons, so a deal is sized as
  though it pays the Inputs-page rate and then actually pays the entered one.
  Entering a 6.250% scale against a 5.875% Inputs rate dropped achieved coverage
  to 1.29x against a 1.30x target (minimum 0.76x against 0.82x). The port
  charges the wrap the weighted-average coupon of the maturities still
  outstanding, solved as a fixed point from the flat-rate schedule.
* **Pricing uses the flat sizing rate as the coupon.** `BondTranche.price_for`
  passes `self.rate` — the Inputs-page rate the structure is *sized* with — to
  the pricing engine, so the Coupon column on the Debt Structure tab never
  reaches the price. A 3/1/2056 term at 6.250% / 6.625% off a 5.875% flat rate
  prices at 90.334 instead of 95.167: 4.8 points of OID, which is a Source of
  Funds and moves the reimbursement. The port prices at `coupon_for(...)`, and
  makes `coupon_for` respect term-bond membership so a later term does not
  inherit an earlier term's coupon.
* **The builder lot inventory build keys off home closings, not lot
  deliveries.** Upstream feeds the "Value of New Lots" column the value of lots
  *converting to homes* that year, so a builder holding delivered lots shows
  $0 of inventory until homes start closing. A lot platted during year y−1 is
  on the roll set 1 January of year y and is assessed whether or not a house
  stands on it. The port feeds the column the prior year's deliveries, which
  makes the cumulative telescope to lots delivered less homes closed — the
  inventory actually held. It looks correct upstream only because the Viridian
  No. 1 closings equal the prior year's deliveries exactly.
* `SummaryModel.build` loops collection years to `dev.last_year` (2067) while
  the value builds stop at the senior final maturity, so the Summary tail shows
  a decade of zero taxable value and fee-only negative revenue. The port bounds
  the loop to the builds. No effect on sizing.
* `SubordinateLien` measures the first (stub) coupon period only when the
  payment date falls in the **same calendar year** as the dated date. That holds
  in Colorado (dated 1 December, sub pays 15 December) but not in Utah, where a
  September delivery's first sub payment is the following 15 March — upstream
  charges that period a full year of interest on bonds dated less than six
  months earlier. The port anchors the stub to the dated date itself, which
  moves the sized sub par materially — on the reference deal it is worth about
  $29,000 of subordinate par.

## Layout

```
ut_pid_model/
  config.py          # ModelConfig — every input, plus the Utah statutory tables
  development.py     # ProductLine + aggregate absorption, taxable-value drivers
  summary.py         # taxable value + pledged revenue engine (the "Summary" sheet)
  debt_service.py    # senior-lien revenue-wrap sizing + amortization (CAPI, DSRF)
  subordinate.py     # subordinate cash-flow lien waterfall
  sources_uses.py    # Sources & Uses + developer reimbursement
  refunding.py       # refinancing of the senior new-money bonds ("new money")
  pricing.py         # price / yield-to-worst, premium & OID
  inputs.py          # Excel inputs page (write template / load config from Excel)
  report.py          # formatted multi-sheet Excel output (model view)
  residential_report.py  # the Development Projections tab
  forecast_report.py # CPA-style forecast exhibits (Exhibit A/B/C, each -1..7)
  scenarios.py       # base + development stress cases (slowed absorption)
  memo.py            # Tierra-style reimbursement memo (HTML)
main.py                        # console run + exports
build_notebook.py              # regenerates the notebook
tools/port_from_colorado.py    # re-port from co_metro_model
ut_pid_model_notebook.ipynb    # narrated walkthrough with charts & scenarios
tests/test_model.py            # statute + reference-deal regression tests
docs/utah-vs-colorado.md       # the statutory walk-through
```

## Running

```bash
pip install -r requirements.txt

python main.py                                  # console summary + output/
jupyter notebook ut_pid_model_notebook.ipynb
```

## Driving the model from Excel (inputs page)

`write_inputs_workbook()` produces an editable **Inputs** workbook — an *Inputs*
sheet (Title | Value | Range Name | Notes, editable **yellow** value cells), a
*Development Inputs* sheet (product definitions plus lot delivery and home
closing grids), a *Debt Structure* sheet (per-maturity par, coupon, yield and
Serial/Term structuring — see below), and a **Utah Property Tax Reference** sheet carrying the
residential exemption history, the editable builder-inventory ratio table, and
the Utah property tax calendar.

```python
from ut_pid_model import write_inputs_workbook, load_inputs_workbook, SummaryModel
write_inputs_workbook(output_path="my_inputs.xlsx")   # blank, pre-populated
cfg, dev = load_inputs_workbook("my_inputs.xlsx")     # read your edits back
sm = SummaryModel(cfg, dev.build(cfg)).build()
```

```bash
python -c "import main; main.run_model(inputs_path='my_inputs.xlsx')"
```

### The Debt Structure tab

One row per maturity, six columns — **Maturity Year | Par Amount | Coupon |
Yield | Type (Serial/Term) | Price** — for the senior bonds and again for the
refunding:

* **Leave it blank** for preliminary work. Sizing runs off the single flat rate
  on the Inputs page, exactly as before.
* **Serial maturities** — fill Coupon and Yield on each row and mark it
  `Serial`.
* **Term bonds** — fill the row of the term's *final* maturity and mark it
  `Term`. Everything up to that maturity amortizes into it at that coupon and
  yield. A single Term row therefore makes the whole structure one amortizing
  term bond. Consecutive Term rows define successive terms.
* **Price** is optional and sits in the last column, after Type. Leave it blank
  and the model computes the price from the coupon and yield; enter the
  underwriter's price (e.g. `95.148`) and **that price governs the OID
  outright**. Use it whenever you have a pricing run in hand — it removes any
  first-period convention difference between this model and the underwriter's
  program. Entering DBC's 95.148 on Viridian Farm PID No. 2 reproduces its
  −$348,373.60 of OID to the cent, maturity by maturity.
* One filled Coupon row sets the coupon for **every** maturity (the lookup
  carries to the nearest entered row), and a Coupon with no Yield prices off the
  Inputs rate — fill Yield alongside Coupon.
* **Par Amount** is optional and overrides the revenue-wrap amortization with a
  manual principal schedule. Leave it blank to let the model size the
  amortization; fill it to dictate one.

The coupons you enter drive the sizing as well as the interest and the price:
the revenue wrap charges the weighted-average coupon of the maturities still
outstanding, so the deal is sized against the debt service it will actually pay
and lands on your coverage target. **With the sheet filled, the Inputs-page
senior interest rate no longer affects the par** — it is only the fallback for a
blank sheet. On the reference deal a 6.250% scale sizes to $5,430,000 whether
the Inputs rate says 4.000% or 8.000%, against $5,665,000 blank at 5.875%.

Prices are computed on **Excel's `PRICE()`** — the 30/360 clean price, at the
frequency set by `INTEREST_FREQUENCY` — and then **truncated to three decimals**,
the convention DBC prints and computes its OID from. So the printed price and
the premium/OID always agree, and a price you enter to three decimals is used
verbatim. A bond reoffered *at* its coupon is quoted at par.

On Viridian Farm PID No. 2 this reproduces the underwriter's run exactly:
`PRICE(9/30/2026, 3/1/2056, 6.250%, 6.625%, 100, 1)` = 95.14825496 → **95.148**,
OID **−$348,373.60**, first coupon **$188,226** (151/360 from the dated date),
total debt service **$17,139,788**.

Prices run through the existing price-to-worst engine, so a yield above the
coupon prices at a discount — a single `2054 · 6.00% · 6.36% · Term` row on the
reference deal sizes to $5,685,000 par with a $365,154 discount, against
$5,665,000 at the flat rate.

The columns shifted when Par Amount was added, so a sheet saved against the old
four-column layout would read its coupons as par amounts. Loading one is
**refused with an explanation** rather than silently sizing a $0.06 bond —
regenerate the sheet with `write_inputs_workbook()` and re-enter the scale.

### District costs on the Inputs page

**District O&M is one line, not two.** Administration (accounting, audit, legal,
assessor, continuing disclosure) and operations (landscaping, parks and trails,
snow removal, lighting) are the same budget for a Utah PID, which rarely carries
an operations levy to fund either — so they share a single row:

| Row | Range name | Default |
|---|---|---:|
| Starting O&M Expense | `OM_EXPENSE` | $53,060 |
| O&M Expense Growth Rate | `OM_GROWTH_RATE` | 2.0% |
| O&M Expense Taxable Value Limit | `OM_EXPENSE_AV_LIMIT` | $0 (no limit) |
| First Year District Costs Are Charged | `DISTRICT_COST_START_YEAR` | blank ⇒ delivery + 2 |

The base starts in `DISTRICT_COST_START_YEAR` (blank ⇒ two years after closing,
the first year with a full year of collections) and inflates from there. It is
netted from the revenue available to the **senior and the subordinate lien** —
money the district actually spends is available to neither bond — and shows as
its own **− District O&M** column on Summary - Detail and against operations
revenue on the **O&M Revenue** tab.

**`INTEREST_FREQUENCY`** (Bond Structure) is a Semiannual/Annual dropdown, and
it drives the debt service schedules, the pricing and the bond statistics
together:

| | Semiannual (default) | Annual |
|---|---|---|
| Coupons a year | 2 — principal month and six months later | 1 — on the principal date |
| Accrual | half a year each, 30/360 | the full year |
| Price | Excel `PRICE()` frequency 2 | Excel `PRICE()` frequency 1 |
| TIC | semiannual compounding | annual compounding |

Interest accrues from the **dated date** either way, so the first coupon is a
stub whenever the bonds are not dated on a coupon date. On the reference deal
dated 9/26/2024: semiannual pays $145,321 then $168,759, annual pays $143,297
then $332,819 — both 155/360 of a year to start.

Two more rows worth knowing about: **`PROJECTION_YEARS`** (Tax & Valuation,
default 40) sets how far the development, taxable-value, revenue and summary
tabs project from delivery — it floors at the longest bond maturity, so it can
never truncate the revenue a bond sizes against — and **`DEVELOPER_CONTRIBUTION`**
(default $0) adds developer cash as a source of funds, applied to the senior
lien, the subordinate lien, or proportionally by par.

**O&M defaults to zero** — an operating budget is a district-specific number and
the model will not assume one. Enter a starting expense and it flows through
sizing: it is netted from the revenue available to the **senior and the
subordinate lien**, appears as its own column on **Summary - Detail**, and shows
against operations revenue on the **O&M Revenue** tab as a surplus/(deficit). On
the reference deal $40,000 of starting O&M at 3.5% costs about $690,000 of
senior par, $152,000 of subordinate par, and $631,000 of reimbursement.

Deliverables land in a **`reimbursement analysis`** folder beside the inputs
workbook:

Dates read **`m/d/yyyy`** throughout the Excel output — every date-valued cell
in all three workbooks, the title band on each tab, and the prose subtitles on
the CAPI and Call Schedule sheets. **File names are the one exception**: a slash
is a path separator and would be stripped out, leaving `9242026`, so saved files
carry the same date as `m.d.yyyy` (`9.24.2026 - Reimbursement Analysis - …`),
which is the substitution Colorado settled on in `fd4a687`. The memo keeps
long-form prose dates, as a letter should.

* `<m.d.yyyy> - Reimbursement Analysis - <district> - <N> Lots - Tierra Financial
  Advisors.xlsx` — model view. Tabs: **Summary - Light**,
  **Summary - Detail**, **Builder Lot Inventory Value**, **Residential Value**,
  **Development Projections**, Sources & Uses – First, Senior Lien DS – First,
  Subordinate Lien, Senior Surplus Fund, CAPI Fund – First, O&M Revenue,
  Sources & Uses – Refunding, Senior Lien DS – Refunding,
  **Subordinate Lien – Refunding**, Senior Lien Coverage, Call Schedule,
  **Notes**.

  The two **value tabs** are the single source of taxable value: the lot
  inventory build (value of new lots → less lots rolled into homes → net with
  lag → adjustments → cumulative → taxable ratio → taxable value) and the
  residential build (beginning market value → value added to the rolls →
  annual reassessment → adjustments → gross market value → taxable value).
  The Summary tabs and the bond sizing read from them, so the presentation and
  the arithmetic cannot disagree.
* `<m.d.yyyy> - Forecast Exhibits - <district> - <N> Lots - Tierra Financial
  Advisors.xlsx` — CPA-style forecast exhibits for the base case and two
  development stress scenarios. The exhibits are their own deliverable, not an
  appendix to the reimbursement analysis, so the label sits where the model
  view says `Reimbursement Analysis` — right after the date — rather than
  trailing the name, and the reimbursement wording appears nowhere in the file:
  * **Exhibit A** — Base Case (100% of forecast absorption pace)
  * **Exhibit B** — Alternative Scenario (80% pace)
  * **Exhibit C** — Alternative Scenario (45% pace)

  Each set has the master debt-service-fund cash flow plus sub-exhibits 1–7
  (taxable values & net tax revenue, builder lot inventory and residential
  taxable value, residential market value, senior and subordinate debt service,
  and sources & uses).
* `… - Memo.html` — the reimbursement memo, populated from the model:
  the absorption table, the Utah assumption bullets (levy cap, the 45%
  residential exemption on homes and builder inventory, the annual roll, the
  30 November / 1 March calendar), the bond program, and a Uses of Funds table
  split **per series** so the senior and subordinate shares of the
  reimbursement are visible side by side. It opens in Word.

## How the senior bonds are sized (revenue wrap)

Unchanged from Colorado: principal is sized year-by-year so net annual debt
service is exactly covered by pledged net revenue at the target coverage ratio.

```
P_t = FLOOR( ( NetRevenue_t / coverage + DSRF_earnings
               - balance_t · rate ) / 5000 ) · 5000
```

with the par solved by bisection so the schedule self-consistently amortizes.
Capitalized-interest years carry zero principal; the final maturity absorbs the
released debt-service-reserve fund as a balloon.

## Validation against the Viridian Farm workbook

| Item | Model | Pricing-day workbook | Priced deal |
|---|---:|---:|---:|
| Senior new-money par | $5,745,000 | $5,690,000 | $5,645,000 |
| Total taxable value, 2031 roll | $197,316,506 | $197,316,506 | — |
| Net pledged revenue, 2028 roll | $375,746 | $375,625 | — |
| Senior final maturity | 3/1/2054 | 3/1/2054 | 3/1/2054 |
| DSRF | $475,000 | $545,055 | — |
| Subordinate par | $1,139,000 (sized) | $1,000,000 (typed) | $1,000,000 |

Taxable value ties **to the dollar** from the 2031 roll onward. Through
build-out the model runs within 0.2% because lot inventory is carried at the
**inflated** ASP (the current Colorado methodology) where the 2024 workbook used
a flat base ASP. The first roll (2024) sits ~18% under the workbook — the value
build recognises the opening lot inventory differently — but it backs $388 of
revenue in a year with no debt service, so it is tracked rather than chased. The
DSRF differs because the 3-prong test here rounds to $5,000, is measured on
**net** annual debt service, and excludes the final maturity year from the
max-DS prong. The subordinate par
differs because the model **sizes** the largest par the residual cashflow
retires in full, where the workbook carries a hand-typed round number.

Those figures are pinned in `tests/test_model.py`, so a change that moves them
fails the suite.

## Two things worth reading in the output

1. **Coverage dips below 1.00x in the first year after capitalized interest
   runs out** (0.87x in 2028 on the base case). The reference workbook has the
   same shape — the sizer only coverage-tests years that carry principal, so
   interest-only years are not tested. The shortfall is met from the debt
   service reserve. Lengthening the capitalized interest period past 36 months
   closes it.

   The CAPI fund covers the term, not a whole number of coupons: 36 months from
   a 9/26/2024 dated date runs to 9/26/2027, so the coupon that straddles that
   date is split — the fund pays the part accrued on or before it and the
   district pays the rest. The deposit is therefore exactly `par × coupon × 3`,
   which is what "36 months of capitalized interest" is quoted to mean. Funding
   only the coupons falling wholly inside the period would carry the district
   for 35 months here, and for just 29 on an annual-pay bond dated 9/30/2026.
2. **Most of the refunding's "new money" is a new subordinate lien, not an
   interest saving.** Of $1,715,067 of new money, roughly **$1,240,000** is the
   refunding subordinate lien ($1,259,000 of par, sized against the residual
   surplus the defeased new-money sub gives back); the senior refunding itself
   contributes about **$475,000**. Utah's fixed levy caps leave far less
   refunding headroom than Colorado's 50-plus mills, so the headline is largely
   the district re-levering the same residual surplus at 8.125% out to 2059 —
   worth saying out loud before anyone reads it as rate savings. Total developer
   reimbursement is **$6,584,681** ($4,869,614 first financing + $1,715,067).

## Tests

```bash
python -m pytest tests -q
```

## Reference documents

* `UTViridianFarm01a-FIN.pdf` — Viridian Farm PID No. 1 limited offering
  memorandum, September 2024
* `Financial Analysis - Viridian Farms PID (3 MILLS) - Salem_Pricing Day
  (Sept 17 2024).xlsm` — the pricing-day workbook
* `ut_pid_model_inputs - Viridian Farm.xlsx` — the inputs workbook for the
  reference deal
