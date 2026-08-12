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

*Code:* `ModelConfig.mill_levy_governing_doc`,
`ModelConfig.validate()` (flags a levy above 15.000 mills, and flags
`gallagherization = "Yes"` as a Colorado mechanism).

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
`ModelConfig.developed_lot_value` (0.55), `Model._build_taxable_value`.

---

## 3. Reassessment cycle and the value lag

**Colorado.** Reappraisal runs on a biennial cycle, and the template lags new
value two rows before it reaches the collection year.

**Utah.** County assessors update values annually based on a systematic review
of market data, with a detailed review of each parcel at least every five years
(§ 59-2-303.1). Value created in one calendar year therefore appears on the
next year's roll — a one-year lag.

*Code:* `ModelConfig.value_lag_years` (1),
`ModelConfig.reassess_frequency` ("Annual"),
`DevelopmentProjections.av_creation_lagged`.

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
`SummaryRow.tax_revenue_date` (one year after the assessment date).

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

*Code:* narrative only — memo § 2.

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

*Code:* `ModelConfig.uniform_fee_prc` (0.00, range `UNIFORM_FEE_PRC`, aliased
`TAX_COLLECT_SO_PRC`).

---

## 7. County collection cost

**Colorado.** The county treasurer retains a fee (commonly 1.5%) out of the
district's distribution, so it is a direct haircut on pledged revenue.

**Utah.** Assessing and collecting is funded by a separate **multicounty
assessing and collecting levy** imposed on property under § 59-2-1602, not by a
deduction from the taxing entity's distribution. The Viridian indenture still
defines Senior Property Tax Revenues net of "the collection costs of the
County", so the input is retained — at 0.00%.

*Code:* `ModelConfig.county_treasurer_fee` (0.00, range
`COUNTY_COLLECTION_FEE`).

---

## 8. Operations levy and district administration

**Colorado.** Metro districts levy a separate operations and maintenance mill
levy, and the template carves operations out of pledged revenue up to an
assessed-value limit.

**Utah.** The Viridian PID levies for debt service only; ongoing operations sit
with the HOA or the city. District administration — accounting, audit, legal,
assessor filings — is a real cost paid from pledged revenue, so the model
charges it against the senior lien and does **not** add it back when computing
what is available to the subordinate lien. (The Colorado template does add it
back, which double-counts; the Utah workbook in this repo already corrected
that, and the correction is carried here.)

*Code:* `ModelConfig.admin_cost_base`, `Model._apply_revenue`,
`Model._sub_waterfall`.

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

*Code:* `DevelopmentProjections.existing_lot_value`,
`SummaryRow.prior_roll_value`.

---

## 10. Centrally assessed property

**Colorado.** The template carries oil and gas producing properties, assessed at
87.5% of the prior year's selling price.

**Utah.** The State Tax Commission centrally assesses properties operating across
county lines, public utilities, railroads, airline operating property,
geothermal resources, and mines with appurtenant machinery (§ 59-2-201). The
same input rows are retained under Utah names.

*Code:* `ModelConfig.centrally_assessed_value`,
`ModelConfig.centrally_assessed_ratio` (ranges `CENTRALLY_ASSESSED_*`, aliased
`OIL_GAS*`).

---

## Sources

- Utah Code Title 17D, Chapter 4 (Public Infrastructure District Act), esp.
  §§ 17D-4-301, 17D-4-303
- Utah Code Title 59, Chapter 2 (Property Tax Act), esp. §§ 59-2-103,
  59-2-201, 59-2-303.1, 59-2-405, 59-2-503, 59-2-1602
- Utah Admin. Code R884-24P-52 (criteria for determining primary residence)
- *Viridian Farm Public Infrastructure District No. 1*, Limited Offering
  Memorandum, September 2024 (in this repository)
- Colorado reference: Silver Peaks No. 6 (PA-4) Metro District analysis, Tierra
  Financial Advisors (CO-Metro-District-Model repository)
