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
Utah change is what feeds it — pledged revenue net of district administration,
which a Utah PID pays out of the debt service levy rather than a separate
operations levy.

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
closing grids), a *Debt Structure* sheet (per-maturity coupons, yields and term
bonds for pricing day), and a **Utah Property Tax Reference** sheet carrying the
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

Deliverables land in a **`reimbursement analysis`** folder beside the inputs
workbook:

* `ut_pid_model_output.xlsx` — model view. Tabs: **Summary - Light**,
  **Summary - Detail**, **Development Projections**, Sources & Uses – First,
  Senior Lien DS – First, Subordinate Lien, Senior Surplus Fund, CAPI Fund –
  First, O&M Revenue, Sources & Uses – Refunding, Senior Lien DS – Refunding,
  Senior Lien Coverage, Call Schedule, **Notes**.
* `ut_pid_forecast_exhibits.xlsx` — CPA-style forecast exhibits for the base case
  and two development stress scenarios:
  * **Exhibit A** — Base Case (100% of forecast absorption pace)
  * **Exhibit B** — Alternative Scenario (80% pace)
  * **Exhibit C** — Alternative Scenario (45% pace)

  Each set has the master debt-service-fund cash flow plus sub-exhibits 1–7
  (taxable values & net tax revenue, builder lot inventory and residential
  taxable value, residential market value, senior and subordinate debt service,
  and sources & uses).
* `ut_pid_model_memo.html` — the reimbursement memo, populated from the model.

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
| Senior new-money par | $5,670,000 | $5,690,000 | $5,645,000 |
| Total taxable value, 2031 roll | $197,316,506 | $197,316,506 | — |
| Net pledged revenue, 2028 roll | $375,746 | $375,625 | — |
| Senior final maturity | 3/1/2054 | 3/1/2054 | 3/1/2054 |
| DSRF | $490,000 | $545,055 | — |
| Subordinate par | $1,606,000 (sized) | $1,000,000 (typed) | $1,000,000 |

Taxable value ties **to the dollar** from the 2031 roll onward. Through
build-out the model runs 0.2–0.4% high because lot inventory is carried at the
**inflated** ASP (the current Colorado methodology) where the 2024 workbook used
a flat base ASP. The DSRF differs because the 3-prong test here rounds to $5,000
and excludes the final maturity year from the max-DS prong. The subordinate par
differs because the model **sizes** the largest par the residual cashflow
retires in full, where the workbook carries a hand-typed round number.

Those figures are pinned in `tests/test_model.py`, so a change that moves them
fails the suite.

## Two things worth reading in the output

1. **Coverage dips below 1.00x in the first year after capitalized interest
   runs out** (0.82x in 2028 on the base case). The reference workbook has the
   same shape — the sizer only coverage-tests years that carry principal, so
   interest-only years are not tested. The shortfall is met from the debt
   service reserve. Lengthening the capitalized interest period past 36 months
   closes it.
2. **The refunding generates no new money at 3 mills.** Refunding par plus the
   released reserve and surplus on hand just covers defeasing both liens and the
   transaction costs. Utah's tight levy caps leave far less refunding headroom
   than Colorado's 50-plus mills — the capability is modeled and reported
   honestly rather than assumed to pay.

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
