# Utah PID vs. Colorado metropolitan district

Where the two frameworks differ, why, and where the difference lives in the code.
Section references are to the Utah Code unless noted.

---

## 1. Enabling act and levy caps

**Colorado.** A metropolitan district is a special district under Title 32,
Article 1, C.R.S. Its debt service mill levy is capped by the service plan
approved by the municipality, and the cap is *adjusted* ("gallagherized") when
the statewide residential assessment rate changes, so that a fall in the
assessment rate does not gut the district's revenue.

**Utah.** A PID is organised under the Public Infrastructure District Act, Title
17D, Chapter 4. Three caps apply and the most restrictive controls:

| Source | Rate | Mills |
|---|---|---|
| § 17D-4-303 | 0.015 per dollar of taxable value | 15.000 |
| Governing document | deal-specific | Viridian: 5.000 |
| Indentures | deal-specific | Viridian: 3.000 |

There is no gallagherization analogue: a Utah cap is a fixed rate per dollar of
taxable value and does not move when the residential exemption changes. The
statutory cap does not bind a levy to pay principal of and interest on a voted
general obligation bond.

*Code:* `ModelConfig.mill_levy_governing_doc`, `ModelConfig.mill_levy_indenture`,
`ModelConfig.mill_levy_cap` (returns the controlling cap),
`ModelConfig.effective_ds_mill_levy` (holds the target down to it), and
`ModelConfig.validate()` (flags a levy above the statutory cap).

---

## 2. Residential assessment — the biggest single difference

**Colorado.** Residential property is assessed at the statutory residential
assessment rate (6.7% in the reference Silver Peaks model); vacant land and most
commercial property at 29%.

**Utah.** § 59-2-103 grants a **45% primary residential exemption** for up to one
acre of land per residential unit, so primary residential property is taxed on
**55% of fair market value**. Part-year residential property qualifies if used
residentially for 183 or more consecutive days.

**Builder inventory.** Utah Admin. Code R884-24P-52 lets the exemption reach
unoccupied property and property under construction: on a written declaration,
or where the assessor determines the property will qualify as a primary
residence when occupied, the residential exemption applies while unoccupied.
That is why developer-held finished lots are carried at 55% here, rather than at
Colorado's 29% vacant-land rate.

The practical effect: a Utah PID's taxable base per home is roughly eight times
the Colorado base for the same house, which is why a 3-mill Utah levy does the
work of a 60-plus-mill Colorado levy.

*Code:* `ModelConfig.resid_taxable_ratio` (0.55),
`ModelConfig.lot_inventory_taxable_ratio` (0.55),
`config.RESIDENTIAL_EXEMPTION_HISTORY`, `config.BUILDER_INVENTORY_HISTORY`,
and `SummaryModel.build`.

---

## 3. Reassessment cycle and the value lag

**Colorado.** Reappraisal runs on a biennial cycle: the level of value is set as
of 30 June the year before an odd-year reappraisal and is **held flat across the
two-year cycle**, so the base steps up only every other year.

**Utah.** County assessors update values annually based on a systematic review
of market data, with a detailed review of each parcel at least every five years
(§ 59-2-303.1). Value created in one calendar year appears on the next year's
roll, and there is **no hold** — the base steps up every year. That is the
substantive difference; the total lag to the debt service payment happens to be
the same two years in both states (see the note below).

*Code:* `ModelConfig.av_lag_years` (2 — see the note below),
`ModelConfig.reassess_frequency` ("Annual"),
`DeveloperProjections._is_reassess_year`, `ModelConfig.av_source_year`.

**A note on the lag number.** Both states run a two-year lag from value creation
to the debt service it supports, by different routes. Utah: created in year *V*
→ 1 January roll for *V+1* → billed 30 November of *V+1* → pays 1 March of
*V+2*. Colorado: created in *V* → June-30 level of value → roll for *V+1* →
collected in *V+2* → pays 1 December of *V+2*. The number is the same; what
differs is the two-year hold in between, which Utah does not have.

---

## 4. Tax calendar and debt service dates

| | Colorado | Utah |
|---|---|---|
| Lien / valuation date | 1 January | 1 January |
| Assessor completes | — | before 22 May (locally assessed) |
| Rate adopted | — | before 22 June |
| Valuation notice | — | by 22 July |
| Tax notice | — | by 1 November |
| Taxes due | half 28 Feb / half 15 Jun (or full 30 Apr) | **30 November**, single payment |
| Delinquency penalty | — | greater of 2.5% or $10; interest from 1 January at the federal funds target + 6%, floored at 7% and capped at 10% |
| PID penalty | — | additional 7% annually under the PID Act |
| Tax sale | — | May/June of the fifth year after assessment |

The single 30 November due date is why Utah PID debt service is structured with
principal on **1 March** and interest on 1 March / 1 September, against
Colorado's 1 December principal. Each year's collections are in hand before the
payment they support.

*Code:* `ModelConfig.prin_maturity` (3), `ModelConfig.int_maturity` (9),
`ModelConfig._snap_to_payment_date` (call and capitalized-interest dates land on
a payment date, not an anniversary of closing).

---

## 5. Truth in Taxation vs. TABOR

**Colorado.** TABOR requires voter approval for a new or increased mill levy;
the district's service plan cap is the practical ceiling.

**Utah.** A taxing entity proposing a rate above the **certified tax rate** — the
rate that reproduces last year's budgeted revenue excluding new growth — must
give notice and hold a public hearing (Truth in Taxation). But the PID Act
exempts the Senior and Subordinate Required Mill Levy from that notice and
hearing so long as the rate does not exceed the cap in the Act, the governing
document, or the indentures. A PID levy within its caps therefore carries no
annual political risk, which is a materially better credit feature than the
Colorado equivalent.

*Code:* narrative only — `memo.py`, the mill-levy assumption bullet.

---

## 6. Vehicle and personal property revenue

**Colorado.** Specific ownership tax on motor vehicles is distributed to taxing
entities and typically runs 6–8% of the property tax levy. The template models
it as a percentage add-on.

**Utah.** There is no specific ownership tax. Registered personal property pays a
statewide **uniform fee in lieu of ad valorem tax** (§ 59-2-405), and the revenue
collected in each county is distributed to each taxing entity "in the same
proportion in which revenue collected from ad valorem real property tax is
distributed." The Viridian indenture includes the allocation in Senior Property
Tax Revenues; the pricing model conservatively credits none of it.

*Code:* `ModelConfig.uniform_fee_prc` (0.00, range `UNIFORM_FEE_PRC`);
`SummaryRow.uniform_fee_revenue`.

---

## 7. County collection cost

**Colorado.** The county treasurer retains a fee (commonly 1.5%) out of the
district's distribution, so it is a direct haircut on pledged revenue.

**Utah.** Assessing and collecting is funded by a separate **multicounty
assessing and collecting levy** imposed on property under § 59-2-1602, not by a
deduction from the taxing entity's distribution. The Viridian indenture still
defines Senior Property Tax Revenues net of "the collection costs of the
County", so the input is retained — at 0.00%.

*Code:* `ModelConfig.county_collection_fee` (0.00, range
`COUNTY_COLLECTION_FEE`).

---

## 8. Operations levy and district administration

**Colorado.** Metro districts levy a separate operations and maintenance mill
levy, and the template carves operations out of pledged revenue up to an
assessed-value limit.

**Utah.** The Viridian PID levies for debt service only; ongoing operations sit
with the HOA or the city. District administration — accounting, audit, legal,
assessor and continuing-disclosure filings — is a real cost paid from pledged
revenue, so it is charged against the senior lien.

Two refinements over the Colorado module, both using inputs that already existed
there: the administration base **inflates** at `admin_growth_rate` (the Colorado
`summary.py` charges a flat amount and never reads the growth rate), and nothing
is charged before `district_cost_start_year` — by default two years after
closing, which is the first roll set with the bonds outstanding and therefore
the first year with a full year of collections to charge against. Both bring the
model onto the reference workbook.

*Code:* `ModelConfig.admin_cost`, `ModelConfig.admin_growth_rate`,
`ModelConfig.district_cost_start_year`, `ModelConfig.district_costs()` (used by
both `SummaryModel.build` and the report, so the two cannot disagree).

---

## 9. Agricultural land

**Utah.** Five or more contiguous acres in active agricultural use may be
assessed at agricultural-use value (the **Greenbelt Reduction**, § 59-2-503).
Withdrawal triggers a rollback tax for up to five years, plus penalties of up to
2% if the withdrawal application is not recorded within 120 days. At Viridian,
roughly 98 acres carried a Greenbelt Reduction of about $7.2 million against a
$13.1 million market value as of 1 January 2024, and the Developer is
contractually responsible for rollback taxes.

The model does not project greenbelt land as a revenue source — it is
extinguished as the ground develops, and the projection starts from the
developed program. Set `DevelopmentProjections.existing_lot_value` if a specific
deal needs the standing base carried.

*Code:* `ModelConfig.existing_vacant_land`, `ModelConfig.historical_av`,
`DeveloperProjections.existing_value_adjustments`.

---

## 10. Centrally assessed property

**Colorado.** The template carries oil and gas producing properties, assessed at
87.5% of the prior year's selling price.

**Utah.** The State Tax Commission centrally assesses properties operating across
county lines, public utilities, railroads, airline operating property,
geothermal resources, and mines with appurtenant machinery (§ 59-2-201). The
same input rows are retained under Utah names.

*Code:* `ModelConfig.centrally_assessed_value`,
`ModelConfig.centrally_assessed_equipment`,
`ModelConfig.centrally_assessed_ratio`, `ModelConfig.centrally_assessed_av`.

---

## Sources

- Utah Code Title 17D, Chapter 4 (Public Infrastructure District Act), esp.
  §§ 17D-4-301, 17D-4-303
- Utah Code Title 59, Chapter 2 (Property Tax Act), esp. §§ 59-2-103,
  59-2-201, 59-2-303.1, 59-2-405, 59-2-503, 59-2-1602
- Utah Admin. Code R884-24P-52 (criteria for determining primary residence)
- *Viridian Farm Public Infrastructure District No. 1*, Limited Offering
  Memorandum, September 2024 (in this repository)
- Colorado reference: `co_metro_model` on branch
  `claude/co-metro-district-model-hawhos` of the CO-Metro-District-Model
  repository — this model is a port of it


---

## 12. Module-by-module map

The Utah package mirrors `co_metro_model` file for file. Attribute names are
kept identical wherever the concept survives, and renamed only where the
Colorado name names a Colorado mechanism:

| Colorado | Utah | Why |
|---|---|---|
| `tabor_current` | `resid_taxable_ratio` | Utah has no TABOR ratio; it has an exemption |
| `tabor_service_plan` | `resid_taxable_ratio_prior` | no service plan, no gallagherization base |
| `gallagherization` | *(removed)* | caps are fixed rates; replaced by `mill_levy_cap` |
| `mill_levy_service_plan` | `mill_levy_governing_doc` (+ `mill_levy_indenture`) | the two caps that actually bind |
| `biennial_reassess_rate` | `reassess_rate` (+ `reassess_frequency`) | Utah revalues annually |
| `developed_lot_value` | `lot_inventory_taxable_ratio` | builder inventory, not a vacant-land class |
| `vacant_land_ratio` | `lot_inventory_ratio` | same |
| `tax_collect_so_prc` | `uniform_fee_prc` | uniform fee, not specific ownership tax |
| `county_treasurer_fee` | `county_collection_fee` | no treasurer haircut in Utah |
| `oil_gas_*` | `centrally_assessed_*` | § 59-2-201 centrally assessed property |
| `om_carveout` | `admin_cost` | charged against pledged revenue, not an operations levy |
| `metro_name` | `pid_name` | it is a public infrastructure district |
| `TABOR_HISTORY` | `RESIDENTIAL_EXEMPTION_HISTORY` | 25% (1982) → 45% (1995) |
| `VACANT_LAND_HISTORY` | `BUILDER_INVENTORY_HISTORY` | R884-24P-52 treatment |
| `strict_biennial_av` | `hold_value_flat` | retained as an option, off for Utah |

Everything else — `debt_service.py`, `sources_uses.py`, `refunding.py`,
`pricing.py`, `scenarios.py` — is the Colorado engine unchanged, because the
sizing, waterfall, call provisions and refunding mechanics are the same in both
states.

The port is mechanical: `tools/port_from_colorado.py` regenerates every module
from a Colorado checkout, applying the renames in this table, the display-label
substitutions, and the semantic patches above. Every patch is checked, so a
Colorado refactor that invalidates one fails the port instead of silently
leaving Colorado behaviour in a Utah model.

---

## 13. What rides in the port beyond the renames

Two upstream defects are fixed in the patch set and should go back to
`co_metro_model`:

* **`SubordinateLien.size_par` rounds to the nearest $1,000.** The bisection
  solves the largest par the residual cashflow retires, then rounds — and
  rounding up lands above it, so the base case reports the note as not fully
  repaid by a few thousand dollars. The port rounds down.
* **`SummaryModel.build` runs past the value builds.** The loop ends at
  `dev.last_year` (2067) while the value builds stop at the senior final
  maturity, so the Summary carries a decade of zero taxable value and fee-only
  negative revenue. The port bounds the loop to the builds. Nothing is sized
  past final maturity, so there is no numerical effect — only a projection that
  stops where the projection stops.

One Colorado change needed a Utah-specific adjustment rather than a straight
port: `residential_assessment_rate` now falls back to the historical rate table
for the roll year instead of the flat config ratio, which is right for Colorado
(the ratio moves every cycle) but would make the Inputs-page
`RESID_TAXABLE_RATIO` cell inert in Utah, where the ratio has been flat at 55%
since 1995. The Utah version honours an explicitly-changed input first, then
falls back to the table.
