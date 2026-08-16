"""Generate ut_pid_model_notebook.ipynb (run once)."""
import nbformat as nbf

nb = nbf.v4.new_notebook()
cells = []

def md(s): cells.append(nbf.v4.new_markdown_cell(s))
def code(s): cells.append(nbf.v4.new_code_cell(s))

md("""# Utah Public Infrastructure District Financial Model

*District name, county, and all assumptions are entered on the Excel Inputs page.*

A Python conversion of the Tierra Utah PID financial-analysis
workbook.

This notebook walks through the full financing:

1. **Inputs / assumptions** (`ModelConfig`)
2. **Development & taxable value** — lot deliveries, home closings, Residential Exemption
   taxable ratio, reassessment
3. **Pledged revenue** — mill levy + specific-ownership tax, net of fees
4. **Senior new-money bonds** — sized by *revenue-wrap* at a target coverage
5. **Sources & Uses + reimbursement** to the developer
6. **Subordinate cash-flow lien**
7. **Refunding** — *refinancing the senior new-money bonds* to generate
   additional reimbursement ("new money")

> Utah PIDs assess property at a primary residential ratio
> (6.7%), reassessed on a two-year cycle, and debt is pledged a mill levy
> (capped by the governing document) plus specific-ownership tax, layered into senior
> and subordinate liens.""")

code("""import pandas as pd
from ut_pid_model import (
    ModelConfig, DeveloperProjections, SummaryModel, SeniorLienSizer,
    SubordinateLien, RefundingAnalysis, first_financing_sources_uses,
    schedule_dataframe, build_excel_report,
)
pd.options.display.float_format = lambda x: f"{x:,.0f}" """)

md("""## 0. Excel inputs page

Drive the whole model from an Excel **Inputs** workbook. Set `INPUTS_PATH` to
your file and copy your filled-in inputs page there (or run this cell once to
generate a blank, pre-populated template at that path, then edit the **yellow
Value cells** on the *Inputs* sheet and the *Development Inputs* sheet).

Everything downstream (taxable value, debt service, exhibits, stress cases)
uses the `cfg` and `dev` loaded here.""")
code("""import os
from ut_pid_model import write_inputs_workbook, load_inputs_workbook

# >>> EDIT THIS PATH: copy your inputs workbook here (or let it be generated) <<<
INPUTS_PATH = r"ut_pid_model_inputs - testing v5.xlsx"
# Deliverables land in a "reimbursement analysis" folder sitting at the SAME tier
# as the inputs workbook (next to INPUTS_PATH), not under the code folder.
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(INPUTS_PATH)),
                          "reimbursement analysis")
os.makedirs(OUTPUT_DIR, exist_ok=True)

if not os.path.exists(INPUTS_PATH):
    write_inputs_workbook(output_path=INPUTS_PATH)
    print(f"Generated a blank inputs template at: {INPUTS_PATH}")
    print("Edit the yellow Value cells, save, and re-run this cell.")

cfg, dev = load_inputs_workbook(INPUTS_PATH)   # config + development from Excel
if cfg.absorption_pace_factor != 1.0:          # home-sales pacing stress (inputs page)
    dev = dev.stressed(cfg.absorption_pace_factor)
    print(f"Home-sales pace stressed to {cfg.absorption_pace_factor:.0%} of forecast.")
dev = dev.build(cfg)
print(f"Loaded inputs from: {INPUTS_PATH}")""")

md("## 1. Inputs / assumptions (loaded from the Excel page above)")
code("""print(f"District:           {cfg.pid_name}")
print(f"Mill levy (DS):     {cfg.mill_levy_ds_target:.0f}")
print(f"residential taxable ratio:        {cfg.resid_taxable_ratio:.2%}")
print(f"Biennial reassess:  {cfg.reassess_rate:.1%}")
print(f"Senior rate:        {cfg.senior_interest_rate:.2%}  ({cfg.dsc_senior:.2f}x coverage)")
print(f"Refunding rate:     {cfg.senior_refunding_interest_rate:.2%}  ({cfg.dsc_refunding:.2f}x coverage)")
print(f"Sub rate:           {cfg.sub_interest_rate:.2%}")
print(f"Delivery:           {cfg.delivery}  |  Refunding delivery: {cfg.delivery_refunding}")""")

md("""## 2. Development & taxable value

**Colorado assessed-value lag.** Colorado reassesses real property on a
**two-year cycle** (in odd-numbered re-valuation years; even years are
intervening years where value carries over). Value is set as of the **June-30
appraisal date in the year before** the reappraisal year, held flat across the
two-year cycle, and the resulting taxes are collected the *following* year. The
net effect is that the taxable value backing a given collection year's
mill-levy revenue reflects market value from roughly **two years earlier**
(`cfg.av_lag_years = 2`), and steps up only every other year. By default the
model applies the reassessment step-up on **odd years** (configurable
via `cfg.reassess_on_even_years`). The model applies this lag in `summary.py`.

> Sources:
> [Adams County Assessor — Property Assessment Process](https://adamscountyco.gov/our-county/elected-officials/assessor/property-assessment-process/),
> [Larimer County — Understanding Property Values](https://www.larimer.gov/assessor/understanding-property-values),
> [Eagle County — Assessment Process](https://www.eaglecounty.us/departments___services/assessor/assessment_process.php).""")
code("""# 'dev' was loaded from the Excel Development Inputs sheet in section 0
print(f"Total lots: {sum(round(c) for c in dev.home_closings.values())}")
print("Home closings:", {y: round(c) for y, c in dev.home_closings.items()})
sm = SummaryModel(cfg, dev).build()
df_sum = sm.to_dataframe()
df_sum.head(15)""")

md("""### Per-product / per-builder development inputs

The development can be modeled as a single aggregate stream **or** as up to eight
product lines (e.g. one builder/one product, or four builders with two products
each). These are entered on the **Development Inputs** sheet (PRODUCT SETUP plus
the Lot Delivery and Home Closings grids) and loaded into `dev` above — nothing
is hardcoded here. Each `ProductLine` carries its own lot deliveries, home
closings, lot value (= ASP × the platted-lot value %), and ASP; the aggregate
drivers and the assessed-value engine roll up from them, and the Summary tab
shows one column per product.""")
code("""# Product lines come straight from the Development Inputs sheet (loaded into 'dev').
prods = dev.product_lines()
mode = "per-product" if dev.is_per_product else "aggregate (single stream)"
print(f"Input mode: {mode}  |  {len(prods)} line(s)  |  total units {dev.total_lots}")
for p in prods:
    print(f"  {p.name:<22} units={p.total_units:>4}  base ASP ${p.asp_base:>10,.0f} "
          f"({p.asp_base_year})")
print("\\nRolled-up home closings:", {y: round(c) for y, c in dev.home_closings.items()})
pd.DataFrame([{
    "product": p.name,
    "lot_value_per_lot": round(p.lot_value(cfg.first_year, cfg.inflation_rate,
                                           cfg.platted_lot_value, cfg.inflation_start_year)),
    "lots_delivered": int(round(sum(p.lot_deliveries.values()))),
    "homes_closed": int(round(sum(p.home_closings.values()))),
} for p in prods])""")

md("""### Statutory level-of-value option (SB 24-233 time-varying rates)

By default the model layers in new development value each year on a 2-year lag
(the practitioner-workbook convention it is validated against). It can also run
the strict **Colorado statutory** treatment: the taxable value is **held flat
across the two-year reassessment cycle**, and the residential / lot-inventory
**taxable ratios vary by year** per SB 24-233 / HB 24B-1001. These are
controlled from the inputs page (`HOLD_VALUE_FLAT`, the residential taxable ratio, and the
assessment-rate schedules), so the table below reflects the **loaded `cfg`** —
the value-set year and the taxable ratios actually in effect.""")
code("""# Reflects the loaded config: AV source year (lag / biennial hold) and the
# taxable ratios in force for each collection year.
y0 = cfg.first_collection_year
pd.DataFrame([{
    "collection_year": y,
    "value_set_as_of": cfg.av_source_year(y),
    "residential_rate": cfg.residential_assessment_rate(y),
    "vacant_land_rate": cfg.lot_inventory_taxable_rate(y),
    "total_AV": round(sm.total_av(y)),
} for y in range(y0, y0 + 8)])""")

md("""### Oil & gas and commercial value in the taxable AV

Oil & gas producing property and commercial property can be added to the total
taxable taxable value, each taxed at its own taxable ratio (centrally assessed at the
statutory ~87.5%, commercial at its taxable ratio and the commercial mill
levy). These come from the inputs page (the *Centrally Assessed / Commercial* section and
the Development Inputs commercial schedule). When there are no such values, those
columns stay zero and are **hidden** in the Excel output. The cell below reads
the **loaded `cfg`/`sm`** — no hardcoded values.""")
code("""print(f"Oil & gas in taxed AV: {cfg.centrally_assessed}   commercial in taxed AV: {cfg.comm_new_value_add}")
print(f"Oil & gas taxable value: ${cfg.centrally_assessed_av:,.0f}  "
      f"(= ${cfg.centrally_assessed_value:,.0f} x {cfg.centrally_assessed_ratio:.1%} + "
      f"${cfg.centrally_assessed_equipment:,.0f} x {cfg.lot_inventory_taxable_ratio:.0%})")
print(f"Commercial mill levy: {cfg.commercial_mill_levy:.1f}   "
      f"commercial taxable ratio: {cfg.commercial_assessment_ratio:.1%}")
sm.to_dataframe()[["collection_year","residential_av","lot_av","centrally_assessed_av",
                   "commercial_av","total_av","mill_revenue"]].head(8)""")

md("## 3. Pledged revenue (mill levy + specific-ownership tax, net of fees)")
code("""df_sum[["collection_year","mill_revenue","net_senior_revenue","net_sub_revenue"]].head(12)""")

md("""## 4. Senior new-money bonds — revenue-wrap sizing

Principal is sized year-by-year so that net annual debt service is exactly
covered by the pledged net revenue at the target coverage ratio:

$$P_t = \\left\\lfloor \\frac{(\\text{NetRevenue}_t + \\text{DSRF earnings})/\\text{coverage} - \\text{balance}_t \\cdot r}{5000} \\right\\rfloor \\times 5000$$

with the par amount solved so it self-consistently amortizes.""")
code("""from ut_pid_model import CallProvisions
# Optional-redemption provisions: call-protected until the premium-call date,
# then callable at a premium (e.g. 103%), declining to par at the par-call date.
senior_calls = CallProvisions(
    premium_call_date=cfg.premium_call_date,
    par_call_date=cfg.par_call_date,
    premium_call_price=cfg.premium_call_price,
)
from ut_pid_model import size_senior_with_dynamic_dsrf
senior = size_senior_with_dynamic_dsrf(
    SeniorLienSizer(cfg, sm),
    name=f"Senior Bonds (Series {cfg.delivery_year}A)",
    rate=cfg.senior_interest_rate, coverage=cfg.dsc_senior, delivery=cfg.delivery,
    first_principal_year=cfg.senior_first_principal_year,
    final_year=cfg.senior_final_year, capi_end_year=cfg.capi_end_date.year,
    call_provisions=senior_calls, reoffering_yield=cfg.senior_reoffering_yield,
    coupon_scale=cfg.senior_coupon_scale, yield_scale=cfg.senior_yield_scale,
    term_bonds=cfg.senior_term_bonds,
)
print(f"Sized senior par: ${senior.par_amount:,.0f}   "
      f"DSRF (3-prong): ${senior.dsrf_deposit:,.0f}")
print(f"Optional redemption: call-protected to {senior_calls.premium_call_date} "
      f"-> {senior_calls.premium_call_price:.0f}% premium call "
      f"-> par call {senior_calls.par_call_date}")
schedule_dataframe(senior).head(12)""")

md("""**Optional redemption / premium call.** The bonds are call-protected until
the first optional redemption date, then callable at a premium (103%) declining
to par. This is what enables the refunding — the redemption price applied is
driven by the refunding date:""")
code("""from datetime import date
for d in [date(2028,1,1), cfg.premium_call_date, date(2030,6,1), cfg.par_call_date, date(2035,1,1)]:
    p = senior.call_price_on(d)
    print(f"{d}:  callable={senior.is_callable_on(d)!s:<5}  redemption price = {p if p else 'N/A (protected)'}")""")

md("""### Annual debt-service coverage — senior lien

For each collection year, **coverage = net pledged senior revenue ÷ senior net
debt service**, compared against the target factor (`dsc_senior`). Coverage is
thin during construction (revenue still ramping) and stabilizes above target as
the assessed-value base matures.""")
code("""from ut_pid_model import senior_coverage_dataframe
cov = senior_coverage_dataframe(cfg, sm, senior)
print(f"Target {cfg.dsc_senior:.2f}x  |  min {cov['coverage'].min():.2f}x  |  "
      f"max {cov['coverage'].max():.2f}x  |  years below target: "
      f"{int((~cov['meets_target']).sum())}")
cov.head(14)""")

md("""## 5. Senior surplus fund, dynamic subordinate par, Sources & Uses

The senior surplus / debt-service-reserve fund builds from the senior residual
to its target; excess cash flows to the subordinate lien. The **subordinate par
is sized dynamically** to the largest amount the residual surplus repays, and
the DSRF, sub par, and refunding amounts are all derived from the district's
taxable value and revenue (nothing hardcoded).""")
code("""from ut_pid_model import SurplusFund
surplus = SurplusFund(cfg, sm).build(senior, None, cfg.first_collection_year, senior.final_year)
sub_par = (cfg.sub_par if cfg.sub_par is not None
           else SubordinateLien(cfg, sm).size_par(
               senior, cfg.first_collection_year, senior.final_year, surplus_fund=surplus))
print(f"Surplus-fund target: ${surplus.target:,.0f}")
print(f"Subordinate par (sized to residual surplus): ${sub_par:,.0f}")

su = first_financing_sources_uses(cfg, senior, sub_par=sub_par)
print("\\nUSES OF FUNDS")
for k, v in su.uses.items():
    print(f"  {k:<28} ${v:>13,.0f}")
print(f"\\nTotal sources ${su.total_sources:,.0f}  ==  Total uses ${su.total_uses:,.0f}  (balanced={su.balanced})")
print(f">>> Developer reimbursement: ${su.reimbursement:,.0f}")""")

md("""## 6. Subordinate cash-flow lien

The subordinate lien is a **cash-flow bond**: coverage factor (`DSC_SUB`),
interest at the sub rate that **compounds while unpaid**, repaid only from
residual surplus after the senior lien once the reserve target is satisfied.""")
code("""sub = SubordinateLien(cfg, sm).size(sub_par, senior, cfg.first_collection_year,
                                    senior.final_year, surplus_fund=surplus)
print(f"Subordinate par: ${sub.par_amount:,.0f}   coverage: {sub.coverage:.2f}x   "
      f"rate: {sub.rate:.0%} (accreting)")
print(f"Total debt service: ${sub.total_payments:,.0f}  "
      f"(interest ${sub.total_interest_paid:,.0f} + principal ${sub.total_principal_paid:,.0f})")
print(f"Fully repaid: {sub.fully_repaid}   ending accrued interest: ${sub.ending_accrued_interest:,.0f}")
SubordinateLien(cfg, sm).to_dataframe(sub).head(18)""")

md("""## 7. Refunding — refinancing the senior new-money bonds

Once the Series 2024A senior bonds become callable, they are **refinanced**
with a larger Series 2029 refunding series at a lower rate (4.50% vs 6.00%).
Because the assessed-value base has grown and the rate has dropped, the
refunding can be sized larger than the bonds it refunds. After defeasing the
old bonds and paying costs, the surplus proceeds become **additional
reimbursement ("new money")** to the developer.""")
code("""# Surplus on hand and the sub escrow at the refunding date are derived from the
# accumulated surplus and the outstanding subordinate balance (not inputs).
_ref_year = cfg.delivery_refunding.year
_sf = {r.year: r for r in surplus.rows}.get(_ref_year)
surplus_on_hand = (cfg.refunding_surplus_on_hand if cfg.refunding_surplus_on_hand is not None
                   else round(_sf.reserve_balance, 2) if _sf else 0.0)
_srow = {r["year"]: r for r in sub.rows}.get(_ref_year)
sub_escrow = (cfg.refunding_sub_escrow if cfg.refunding_sub_escrow is not None
              else (_srow["principal_balance"] + _srow["accrued_balance"] if _srow else 0.0))

refunding = RefundingAnalysis(cfg, sm).run(senior, surplus_on_hand=surplus_on_hand,
                                           sub_escrow=sub_escrow)
print(f"Refunding delivered {cfg.delivery_refunding} at {refunding.call_price:.0f}% call price")
print(f"Refunding par sized:          ${refunding.refunding_bond.par_amount:,.0f}")
print(f"Callable principal refunded:  ${refunding.refunded_par_outstanding:,.0f}")
print(f"Refunding escrow (defeasance):${refunding.refunding_escrow:,.0f}")
print()
print("REFUNDING SOURCES & USES")
for k, v in refunding.sources_uses.sources.items(): print(f"  + {k:<32} ${v:>13,.0f}")
for k, v in refunding.sources_uses.uses.items():    print(f"  - {k:<32} ${v:>13,.0f}")
print(f"\\n>>> Additional reimbursement (NEW MONEY): ${refunding.new_money_reimbursement:,.0f}")""")

code("""schedule_dataframe(refunding.refunding_bond).head(12)""")

md("## 8. Combined result & Excel export")
code("""total = su.reimbursement + refunding.new_money_reimbursement
print(f"First-financing reimbursement:  ${su.reimbursement:,.0f}")
print(f"Refunding 'new money':          ${refunding.new_money_reimbursement:,.0f}")
print(f"TOTAL developer reimbursement:  ${total:,.0f}")

path = build_excel_report(cfg, sm, senior, su, refunding, sub, surplus, dev=dev,
                          output_path=os.path.join(OUTPUT_DIR, "ut_pid_model_output.xlsx"))
print(f"\\nExcel report written to: {path}")

# Tierra-style reimbursement memo (HTML), populated from the model.
from ut_pid_model import build_memo_html
memo_path = build_memo_html(cfg, sm, senior, su, sub, refunding, dev=dev,
                            output_path=os.path.join(OUTPUT_DIR, "ut_pid_model_memo.html"),
                            developer=cfg.developer or "[Developer / Master Developer]")
print(f"Reimbursement memo written to: {memo_path}")""")

md("""### CPA-style forecast exhibits — base + stress cases

A second workbook presents the model in the standard accountant's *Forecast of
Cash Balances and Cash Receipts and Disbursements* exhibit set, for the **base
case and two development stress scenarios**, mirroring the source forecast:

* **Exhibit A** — Base Case (100% of forecast absorption pace)
* **Exhibit B** — Alternative Scenario (80% pace — slower absorption)
* **Exhibit C** — Alternative Scenario (45% pace — much slower absorption)

Each set has the master cash-flow exhibit plus sub-exhibits 1–7. The bonds are
sized once in the base case; the stress cases reduce absorption (extending
build-out) and test the **same** debt service against the lower taxable value
and pledged revenue, producing lower debt-service coverage during construction.""")
code("""from ut_pid_model import build_scenarios, build_forecast_report
# Stress cases slow the *loaded* development (dev); DSRF, sub par, maturities derived from cfg.
scenarios = build_scenarios(cfg, dev)
for s in scenarios:
    closed = {y: round(c) for y, c in s.dev.home_closings.items()}
    net = s.senior.annual_net_ds()
    covs = [s.sm.net_senior_revenue(y) / net[y] for y in sorted(net) if net[y] > 0]
    print(f"Exhibit {s.exhibit} ({s.pace_factor:>4.0%} pace): build-out {max(closed)}, "
          f"min construction-era coverage {min(covs):.2f}x")

fpath = build_forecast_report(scenarios, output_path=os.path.join(OUTPUT_DIR, "ut_pid_forecast_exhibits.xlsx"))
print(f"\\nForecast exhibits written to: {fpath}")
import openpyxl
print("Sheets:", openpyxl.load_workbook(fpath).sheetnames)""")

md("""## 9. Scenario analysis example — what if rates fall further at refunding?

Because the refunding capability is parameterized, you can run a sensitivity by
overriding a single `ModelConfig` field — everything else (development, assessed
value, DSRF, sub par, maturities, refunding amounts) re-derives from the **loaded
`cfg`/`sm`/`senior`**. Here we sweep the refunding rate relative to the loaded
base rate (base, −50 bps, −100 bps); only that one field changes.""")
code("""import dataclasses
base_ref = cfg.senior_refunding_interest_rate
for delta in (0.0, -0.005, -0.010):
    c = dataclasses.replace(cfg, senior_refunding_interest_rate=base_ref + delta)
    rf = RefundingAnalysis(c, sm).run(senior, surplus_on_hand=surplus_on_hand,
                                      sub_escrow=sub_escrow)
    print(f"Refunding rate {c.senior_refunding_interest_rate:.2%} "
          f"({delta*1e4:+.0f} bps) -> par ${rf.refunding_bond.par_amount:>12,.0f}  "
          f"new money ${rf.new_money_reimbursement:>11,.0f}")""")

nb["cells"] = cells
with open("ut_pid_model_notebook.ipynb", "w") as f:
    nbf.write(nb, f)
print("Notebook written.")
