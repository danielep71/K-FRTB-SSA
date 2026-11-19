# 🧮 Calculation contracts

A contract states what one calculation must produce, from which inputs, under
which rules. The VBA implements the contract, the
[validation plan](VALIDATION_PLAN.md) tests against it, and a change to a
result it defines is a `behavior-change`.

The workbook calculates the own funds requirement for market risk under the
simplified standardised approach of `SRC-02`, the Basel Framework's chapter
MAR40 ([sources](SOURCES.md)). It is a prototype: the rules below are
implemented, none is validated, and nothing was run in Excel
([validation plan](VALIDATION_PLAN.md)).

## Common rules

These hold for every calculation:

- **Units and signs:** every amount is in currency units, unscaled. A
  position is entered as a value of zero or more with a *Direction*: `Long`
  is positive and `Short` negative. Every result is in the reporting
  currency, `ReportingCcy`. A rate, a weight or a coupon is a fraction: 8%
  is `0.08`.
- **Dates:** the reporting date is `AsOfDate`. A residual term is the number
  of calendar days from `AsOfDate` to the date, divided by `DaysPerYear`,
  365. A month is a twelfth of a year. `SRC-02` gives its time bands in
  months and years and names no day count: this is an assumption of the
  workbook.
- **Band limits:** a residual term equal to the upper limit of a band
  belongs to that band, as the source says for specific risk, *6 months or
  less*. For the time bands of general market risk the source does not say
  it: an assumption of the workbook.
- **Precision and rounding:** VBA `Double` arithmetic throughout, and no
  rounding: `SRC-02` requires none.
- **Currency conversion:** `FX_Table` on `Params` gives, for each currency
  and for gold, the units of reporting currency per unit, which the user
  enters as the spot rate of the reporting date; the reporting currency is
  listed with rate 1. Every position is converted before any other step.
  Because each charge is a sum of fixed rates applied to amounts, converting
  first gives the same result as converting each currency's charge.
- **Rejected rows:** a row that fails a test of its sheet contributes
  nothing. Its status gives the reason, `Checks` lists it, and the status
  line of the run reads `INCOMPLETE`. No value is corrected.
- **Failure:** an unusable parameter stops the run, which reads `STOPPED`,
  and no result is calculated. A run that raises an error clears every
  output and reads `FAILED`
  ([status and staleness](WORKBOOK_CONTRACT.md#status-and-staleness)).
- **Tolerances:** an expected value that a derivation gives in full
  precision is compared within an absolute bound of `1e-6` currency units
  and a relative bound of `1e-9`; a time band, which is a whole number,
  exactly. A value printed in a source is compared within half a unit of
  its last printed digit, with a relative bound of 0
  ([tolerance](TEST_CASES.md#tolerance)).
- **Inputs:** the *Inputs* tables below name each input as a dataset and a
  field of the [data contract](DATA_CONTRACT.md), which gives its column, its
  type and its allowed values, and says how a cell is read
  ([how a sheet is read](DATA_CONTRACT.md#reading-a-sheet)).

## Parameters

Every factor, time band, weight and threshold is a named value on the
`Params` sheet, with its reference value in `src/core/CORE_Config.bas`,
which `src/core/CORE_Parameters.bas` reads and checks. A value on `Params`
that differs from its reference makes the run a what-if, and `Checks` warns
(rule `C08`). Each reference value was compared with `SRC-02` on 2026-11-16
([verification record](SOURCES.md#verification-record)):

| Parameter | Reference value | Source | Compared |
| --- | --- | --- | --- |
| `ScalingIR`, `ScalingEQ`, `ScalingCO`, `ScalingFX` | 1.30, 3.50, 1.90, 1.20 | `SRC-02`, MAR40.2 | Equal |
| `RWAMultiplier` | 12.5 | `SRC-02`, MAR40.1 | Equal |
| Specific-risk table, `SR_Key` and `SR_Charge`, 18 rows | The charges of table 1 | `SRC-02`, MAR40.6, table 1 | Equal. The table gives the qualifying category one set of charges; the workbook lists it under the two investment grades. It gives *BB+ to B–* for governments and *Below BB–* for other issuers; the workbook lists each under the grades it spans |
| Time bands: `Ladder_Zone`, `Ladder_UpperHigh`, `Ladder_UpperLow`, `Ladder_RW`, 15 rows | The bands, zones and risk weights of table 4 | `SRC-02`, MAR40.26, table 4, and MAR40.28 for the zones | Equal. *1 month* is 1/12 of a year, and *3 months* 0.25 |
| `CouponThreshold` | 0.03 | `SRC-02`, MAR40.26 | Equal |
| `VerticalDisallowance` | 0.10 | `SRC-02`, MAR40.27 | Equal |
| `HorizontalZone1`, `HorizontalZone2`, `HorizontalZone3` | 0.40, 0.30, 0.30 | `SRC-02`, MAR40.28, table 5 | Equal |
| `HorizontalAdjacent`, `HorizontalZone13` | 0.40, 1.00 | `SRC-02`, MAR40.28, table 5 | Equal |
| `EQSpecificRate`, `EQGeneralRate` | 0.08, 0.08 | `SRC-02`, MAR40.43 | Equal |
| `EQIndexRate` | 0.02 | `SRC-02`, MAR40.47 | Equal |
| `FXChargeRate` | 0.08 | `SRC-02`, MAR40.61 | Equal |
| `COMDirectionalRate` | 0.15 | `SRC-02`, MAR40.72 | Equal |
| `COMGrossRate` | 0.03 | `SRC-02`, MAR40.73 | Equal |
| `DaysPerYear` | 365 | None: an assumption of the workbook | Not applicable |
| `GoldCode` | `XAU` | None: the code under which gold is entered, a convention of the workbook. MAR40.53 treats gold as a currency | Not applicable |

No parameter is listed as not yet verified.

## Contract template

Copy this section for each calculation.

### *Illustrative:* `CALC-01` Example calculation

**Purpose:** what the figure is, and where it is reported.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Example amount | `DS-01.notional` | Currency of `DS-01.currency` |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | Example: sum the notional amounts | `SRC-01`, Article 1(1) |

- **Step:** a whole number, once per step. A test case's `rule` and a
  traceability row name step 2 as `CALC-NN.2`, and
  `tools/check_test_cases.py` rejects a step that this table does not list.

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Example total | `illustrative_total_notional_amount` | Reporting currency | Whole calculation | None |

- **Quantity:** the lower snake case name that the test cases' expected
  files and the accuracy contracts use for the output
  ([expected-result format](TEST_CASES.md#expected-format)). A `_status`
  quantity, which reports a rejection, is not listed here: it belongs to
  *Failure behaviour*.

**Limits and non-goals:** cases the calculation does not cover, and what
happens if they occur in the input.

**Failure behaviour:** which conditions stop it, and what it reports: each
`_status` quantity, in lower snake case ending in `_status`, with the exact
texts it can hold, which the `invalid` test cases expect
([invalid inputs](TEST_CASES.md#invalid-inputs)).

**Validation:** the checks in the [validation plan](VALIDATION_PLAN.md) that
cover it, and the rows of its steps in its
[traceability table](VALIDATION_PLAN.md#traceability).

## Contracts

### `CALC-02` Interest rate risk: specific risk

**Purpose:** the capital requirement for the specific risk of debt
positions, the *Specific risk* of the interest rate row of `Results`.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Position value | `DS-06.position_value` and `DS-06.direction` | Currency of `DS-06.currency` |
| Currency and its rate | `DS-06.currency`; `DS-05.currency_id` and `DS-05.fx_rate` | Reporting currency per unit |
| Category and rating grade | `DS-06.sr_category` and `DS-06.rating_grade` | Text |
| Final maturity | `DS-06.final_maturity_date`; `DS-02.as_of_date` and `DS-02.days_per_year` | Date; days |
| Issue | `DS-06.issue_id` | Text; may be empty |
| Charges | `DS-03.key` and `DS-03.charge` | Fraction |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | A position whose category is `NONE` carries no specific risk. It is for the legs of swaps, forward rate agreements, forward FX contracts and interest rate futures, which the user enters so | `SRC-02`, MAR40.38 |
| 2 | Positions in the identical issue are netted: the rows with the same *Issue ID*, without regard to case. A row without an *Issue ID* is an issue of its own. No other offsetting is made | `SRC-02`, MAR40.5 |
| 3 | The charge rate is that of table 1, for the category, the rating grade and, where the table has them, the three ranges of the residual term to final maturity: 6 months or less, over 6 and up to 24 months, over 24 months | `SRC-02`, MAR40.6, table 1 |
| 4 | The charge of an issue is the absolute net value multiplied by the rate, and the requirement is the sum over the issues | `SRC-02`, MAR40.4 and MAR40.5 |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Net value of an issue | `ir_issue_net_value` | Reporting currency | Issue | None |
| Specific risk charge of an issue | `ir_issue_specific_risk_charge` | Reporting currency | Issue | None |
| Specific risk charge | `ir_specific_risk_charge` | Reporting currency | Currency, and whole calculation | None |

**Limits and non-goals:**

- The user classifies each position: government, qualifying or other, and
  its rating grade. The workbook does not apply the criteria of MAR40.7 to
  MAR40.12.
- The rows of an issue should have one category, grade and currency. When
  they do not, the issue takes the rate and the currency of its first row,
  and `Checks` warns (rule `C09`).
- Not supported: the national discretions of MAR40.7, MAR40.8 and MAR40.13;
  securitisation positions (MAR40.14); the cap at the maximum possible loss
  (MAR40.15); the offsets and the treatment of credit derivatives (MAR40.16
  to MAR40.21); and the correlation trading portfolio (MAR40.22).
- The workbook cannot tell an unsupported instrument from a bond. Entered
  as a debt position, it is calculated as one: a stated limit of the
  [data contract](DATA_CONTRACT.md#unsupported-input). Do not enter it. The
  engine reads `DS-06.instrument` and rejects a row whose instrument is not
  one of the seven values the contract allows; `Other` stays among them, so
  that an instrument without a value of its own can be entered. The list
  cannot tell a supported instrument from one entered under its name.
- A category and grade that table 1 does not list, such as `OTHER` with an
  investment grade, reject the row.

**Failure behaviour:** `ir_position_status`, on `IRCalc`, is `OK` or one of
the texts listed under `CALC-03`, which shares the rows.

**Validation:** the test cases
`illustrative-ir-specific-risk-maturity-limits` and
`illustrative-sample-ir-specific-risk`. Their expected values are
`illustrative`: derived within the project, they validate nothing. See the
[validation matrix](VALIDATION_PLAN.md#validation-matrix) and the rows of
`CALC-02` in the [traceability table](VALIDATION_PLAN.md#traceability).

### `CALC-03` Interest rate risk: general market risk, maturity method

**Purpose:** the capital requirement for general market risk of interest
rate positions by the maturity method, the *General market risk* of the
interest rate row of `Results`, with its components per currency on
`IR_Zones` and `IR_Ladder`.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Position value | `DS-06.position_value` and `DS-06.direction` | Currency of `DS-06.currency` |
| Currency and its rate | `DS-06.currency`; `DS-05.currency_id` and `DS-05.fx_rate` | Reporting currency per unit |
| Rate type and coupon | `DS-06.rate_type` and `DS-06.coupon` | Text; fraction |
| Dates | `DS-06.final_maturity_date` and `DS-06.next_repricing_date`; `DS-02.as_of_date` and `DS-02.days_per_year` | Date; days |
| Time bands | `DS-04.zone`, `DS-04.upper_bound_high_coupon`, `DS-04.upper_bound_low_coupon` and `DS-04.risk_weight`; `DS-02.coupon_threshold` | Years; fraction |
| Disallowances | `DS-02.vertical_disallowance` and the five `DS-02.horizontal_` fields | Fraction |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | Each currency has its own maturity ladder. The requirements of the currencies are summed, with no offsetting between them | `SRC-02`, MAR40.24 |
| 2 | A fixed-rate position is slotted by its residual term to final maturity, and a floating-rate position by its residual term to the next repricing date | `SRC-02`, MAR40.25 |
| 3 | The time band and its risk weight are those of table 4: the column for a coupon of 3% or more, or the column for a coupon of less than 3%. A floating-rate position uses the first column | `SRC-02`, MAR40.26, table 4 |
| 4 | The weighted position is the signed value multiplied by the risk weight of its band | `SRC-02`, MAR40.26 |
| 5 | In each band, the vertical disallowance is 10% of the smaller of the weighted longs and the weighted shorts | `SRC-02`, MAR40.27 |
| 6 | In each zone, the net positions of the bands are offset, and the horizontal disallowance is 40% of the matched amount in zone 1 and 30% in zones 2 and 3 | `SRC-02`, MAR40.28, table 5 |
| 7 | The net positions of the zones are offset: zones 1 and 2, then zones 2 and 3, at 40% of the matched amount, and then what is left of zones 1 and 3, at 100% | `SRC-02`, MAR40.28, table 5 |
| 8 | The net position charge is the absolute value of the weighted position that remains | `SRC-02`, MAR40.23 |
| 9 | The requirement of a currency is the sum of steps 5 to 8 | `SRC-02`, MAR40.23 |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Time band of a position | `ir_time_band` | Band number, 1 to 15 | Position | None |
| Weighted position | `ir_weighted_position` | Reporting currency | Position | None |
| Vertical disallowance | `ir_vertical_disallowance` | Reporting currency | Currency | None |
| Horizontal disallowance within the zones | `ir_horizontal_disallowance_within_zones` | Reporting currency | Currency | None |
| Horizontal disallowance between the zones | `ir_horizontal_disallowance_between_zones` | Reporting currency | Currency | None |
| Net position charge | `ir_net_position_charge` | Reporting currency | Currency | None |
| General market risk charge | `ir_general_market_risk_charge` | Reporting currency | Currency, and whole calculation | None |

**Limits and non-goals:**

- Assumptions: the day count and the band limits of the
  [common rules](#common-rules); that a floating-rate position uses the
  column for a coupon of 3% or more, whatever its coupon; and the order of
  step 7, which MAR40.28 does not state.
- A derivative is entered as its legs: the user converts it into positions
  in the underlying, as MAR40.31 to MAR40.34 and table 7 of MAR40.40
  describe, and enters each leg as a row. The workbook converts nothing.
- The workbook does not find the matched positions that MAR40.35 and
  MAR40.36 allow a bank to leave out: leave them out of the input, or they
  are slotted like any other.
- Not supported: the duration method (MAR40.29); a single ladder for
  currencies in which business is insignificant (MAR40.24 and MAR40.30);
  the alternative formulae for swaps (MAR40.37); and the national
  discretion of MAR40.13 to disallow offsetting.
- Gold and the reporting of FX positions belong to `CALC-06`. A row whose
  currency is the gold code is rejected.

**Failure behaviour:** `ir_position_status`, on `IRCalc`, is `OK` or:

- `INVALID: duplicate PositionID`
- `INVALID: Instrument must be Bond, FRN, Swap leg, Future leg, FRA leg, Forward leg or Other`
- `INVALID: currency not in FX_Table`
- `INVALID: gold is FX risk; enter it on FXPositions`
- `INVALID: Direction must be Long or Short`
- `INVALID: Position value must be a number`
- `INVALID: Position value must be >= 0 (Direction gives the sign)`
- `INVALID: Rate type must be Fixed or Floating`
- `INVALID: Final maturity date missing`
- `INVALID: Final maturity date is not a date`
- `INVALID: final maturity before AsOfDate`
- `INVALID: Coupon required for a fixed-rate position`
- `INVALID: Next repricing date required for a floating-rate position`
- `INVALID: Next repricing date is not a date`
- `INVALID: repricing date outside AsOfDate to final maturity`
- `INVALID: SR category must be GOV, QUAL, OTHER or NONE`
- `INVALID: no MAR40.6 charge for this category and rating grade`

The first test that fails gives the text. A rejected row enters neither
`CALC-02` nor this calculation.

**Validation:** the test cases `illustrative-invalid-ir-positions`,
`illustrative-ir-time-band-limits` and
`illustrative-sample-ir-general-market-risk`. Their expected values are
`illustrative`: derived within the project, they validate nothing. See the
[validation matrix](VALIDATION_PLAN.md#validation-matrix) and the rows of
`CALC-03` in the [traceability table](VALIDATION_PLAN.md#traceability).

### `CALC-04` Equity risk

**Purpose:** the capital requirement for equity positions, the equity row of
`Results`, with its components per market on `Buckets`.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Market value | `DS-07.market_value` and `DS-07.direction` | Currency of `DS-07.currency` |
| Currency and its rate | `DS-07.currency`; `DS-05.currency_id` and `DS-05.fx_rate` | Reporting currency per unit |
| Market | `DS-07.market` | Text: the national market |
| Type and identifier | `DS-07.type` and `DS-07.identifier` | Text |
| Rates | `DS-02.eq_specific_rate`, `DS-02.eq_index_rate` and `DS-02.eq_general_rate` | Fraction |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | Positions in the identical equity or stock index in a market are netted: the rows with the same market, type and identifier, without regard to case | `SRC-02`, MAR40.46 |
| 2 | Specific risk is 8% of the absolute net position of each single name | `SRC-02`, MAR40.42 and MAR40.43 |
| 3 | The charge for an index is 2% of the absolute net position of each index, in place of specific risk | `SRC-02`, MAR40.47, and table 8 of MAR40.52 |
| 4 | General market risk is 8% of the absolute net position of each market: the sum of its longs less the sum of its shorts, single names and indices together | `SRC-02`, MAR40.42 and MAR40.43 |
| 5 | The requirement is the sum of steps 2 to 4 over the stocks, the indices and the markets | `SRC-02`, MAR40.42 |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Net position of a stock or an index | `eq_net_position` | Reporting currency | Stock or index in a market | None |
| Specific risk charge | `eq_specific_risk_charge` | Reporting currency | Market, and whole calculation | None |
| Index charge | `eq_index_charge` | Reporting currency | Market, and whole calculation | None |
| General market risk charge | `eq_general_market_risk_charge` | Reporting currency | Market, and whole calculation | None |

**Limits and non-goals:**

- A derivative is entered as a position in its underlying, at the value
  MAR40.45 gives. The workbook converts nothing.
- Every position entered as `Index` is charged 2%. The source applies that
  rate only to a well-diversified index (MAR40.47): the user decides which
  index is one.
- Not supported: the arbitrage strategies of MAR40.48 to MAR40.50, and the
  offset of depository receipts (MAR40.51).
- The interest rate or FX exposure of the other leg of a contract (footnote
  14 of MAR40.44) is entered on `IRPositions` or `FXPositions`.
- `Buckets` lists up to 30 markets. The totals include every market, and
  `Checks` warns (rule `C10`).

**Failure behaviour:** `eq_position_status`, on `EQCalc`, is `OK` or:

- `INVALID: duplicate PositionID`
- `INVALID: Market missing`
- `INVALID: Type must be Single name or Index`
- `INVALID: stock or index identifier missing`
- `INVALID: currency not in FX_Table`
- `INVALID: Direction must be Long or Short`
- `INVALID: Market value must be a number`
- `INVALID: Market value must be >= 0 (Direction gives the sign)`

**Validation:** the test cases `illustrative-eq-matched-positions`,
`illustrative-invalid-eq-positions` and `illustrative-sample-equity`. Their
expected values are `illustrative`: derived within the project, they validate
nothing. See the [validation matrix](VALIDATION_PLAN.md#validation-matrix)
and the rows of `CALC-04` in the [traceability
table](VALIDATION_PLAN.md#traceability).

### `CALC-05` Commodities risk: simplified approach

**Purpose:** the capital requirement for commodity positions by the
simplified approach, the commodity row of `Results`, with its components per
commodity on `Buckets`.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Quantity | `DS-08.quantity` and `DS-08.direction` | The standard unit of the commodity, which `DS-08.unit` names |
| Spot price | `DS-08.spot_price` | Currency of `DS-08.currency`, per unit |
| Currency and its rate | `DS-08.currency`; `DS-05.currency_id` and `DS-05.fx_rate` | Reporting currency per unit |
| Commodity | `DS-08.commodity` | Text |
| Rates | `DS-02.com_directional_rate` and `DS-02.com_gross_rate` | Fraction |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | A position is valued at the current spot price: quantity multiplied by spot price, converted into the reporting currency | `SRC-02`, MAR40.68 and MAR40.73 |
| 2 | Long and short positions in a commodity are netted: the rows with the same *Commodity*, without regard to case. Different commodities are not offset | `SRC-02`, MAR40.67 |
| 3 | The directional charge is 15% of the absolute net position in each commodity | `SRC-02`, MAR40.72 |
| 4 | A further charge is 3% of the gross position in each commodity, long plus short | `SRC-02`, MAR40.73 |
| 5 | The requirement is the sum of steps 3 and 4 over the commodities | `SRC-02`, MAR40.72 and MAR40.73 |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Net position | `co_net_position` | Reporting currency | Commodity | None |
| Gross position | `co_gross_position` | Reporting currency | Commodity | None |
| Directional charge | `co_directional_charge` | Reporting currency | Commodity, and whole calculation | None |
| Charge on gross positions | `co_gross_charge` | Reporting currency | Commodity, and whole calculation | None |

**Limits and non-goals:**

- Not supported: the maturity ladder approach (MAR40.68 to MAR40.70), so
  *Maturity date* on `COPositions` is not used; and netting between
  subcategories of a commodity, a national discretion of MAR40.67.
- Two rows are one commodity only when their names are equal. The workbook
  does not check that their units are.
- A derivative is entered as a notional position in the commodity, as
  MAR40.71 describes. The workbook converts nothing.
- Gold is FX risk (MAR40.63): enter it on `FXPositions`. A row whose
  commodity is the gold code is rejected.
- A spot price is used as entered, also when it is negative.
- `Buckets` lists up to 30 commodities. The totals include every commodity,
  and `Checks` warns (rule `C11`).

**Failure behaviour:** `co_position_status`, on `COCalc`, is `OK` or:

- `INVALID: duplicate PositionID`
- `INVALID: Commodity missing`
- `INVALID: gold is FX risk; enter it on FXPositions`
- `INVALID: Direction must be Long or Short`
- `INVALID: Quantity must be a number`
- `INVALID: Quantity must be >= 0 (Direction gives the sign)`
- `INVALID: Spot price must be a number`
- `INVALID: currency not in FX_Table`

**Validation:** the test cases `illustrative-co-matched-positions`,
`illustrative-invalid-co-positions` and `illustrative-sample-commodities`.
Their expected values are `illustrative`: derived within the project, they
validate nothing. See the [validation
matrix](VALIDATION_PLAN.md#validation-matrix) and the rows of `CALC-05` in
the [traceability table](VALIDATION_PLAN.md#traceability).

### `CALC-06` Foreign exchange risk: shorthand method

**Purpose:** the capital requirement for positions in foreign currencies and
gold by the shorthand method, the FX row of `Results`, with the net position
per currency and the overall net open position on `Buckets`.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Net position | `DS-09.net_position` and `DS-09.direction` | Currency of `DS-09.currency`; for gold, the unit its rate is quoted for |
| Currency and its rate | `DS-09.currency`; `DS-05.currency_id` and `DS-05.fx_rate`; `DS-02.gold_code` and `DS-02.reporting_currency` | Reporting currency per unit |
| Rate | `DS-02.fx_charge_rate` | Fraction |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | The net position in each foreign currency and in gold is the sum of its rows, converted at the spot rate into the reporting currency | `SRC-02`, MAR40.60 |
| 2 | The overall net open position is the greater of the sum of the net long positions and the sum of the net short positions in the currencies, plus the net position in gold, whatever its sign | `SRC-02`, MAR40.60 |
| 3 | The requirement is 8% of the overall net open position | `SRC-02`, MAR40.61 |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Net position | `fx_net_position` | Reporting currency | Currency, and gold | None |
| Overall net open position | `fx_overall_net_open_position` | Reporting currency | Whole calculation | None |
| FX charge | `fx_charge` | Reporting currency | Whole calculation | None |

**Limits and non-goals:**

- The user enters the net open position in each currency: its components,
  which MAR40.55 to MAR40.58 list, are summed outside the workbook, and so
  is the valuation of forward positions (MAR40.58).
- The workbook does not derive FX positions from the other sheets. A bond,
  an equity or a commodity position in a foreign currency adds nothing to
  `FXPositions`: include it in the net position entered there.
- A row in the reporting currency carries no FX risk: it is excluded, and
  `Checks` reports it (rule `C12`).
- Not supported: the exemption of MAR40.62, and the use of internal limits
  as a proxy (footnote 22 of MAR40.60).

**Failure behaviour:** `fx_position_status`, on `FXCalc`, is `OK` or:

- `INVALID: duplicate PositionID`
- `INVALID: currency not in FX_Table`
- `INVALID: Direction must be Long or Short`
- `INVALID: Amount must be a number`
- `INVALID: Amount must be >= 0 (Direction gives the sign)`
- `EXCLUDED: reporting currency carries no FX risk`

**Validation:** the test case `mar40-61-table-9`, two of whose expected
values are published in table 9 of MAR40.61, and the test cases
`illustrative-fx-one-sided-positions`, `illustrative-invalid-fx-positions`
and `illustrative-sample-fx`. Every other expected value is `illustrative`:
derived within the project, it validates nothing. See the [validation
matrix](VALIDATION_PLAN.md#validation-matrix) and the rows of `CALC-06` in
the [traceability table](VALIDATION_PLAN.md#traceability).

### `CALC-07` Non-delta charges for options

**Purpose:** the additional requirement for the risks of options that the
delta-equivalent positions do not carry, per risk class: the *Non-delta
options* column of `Results`.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Charge per risk class | `DS-10.charge`, for each `DS-10.risk_class_id` | Reporting currency |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | The charge of each risk class is the value entered. The user calculates it outside the workbook, by one of the methods the source allows: the simplified approach, the gamma and vega charges of the delta-plus method, or the scenario approach | `SRC-02`, MAR40.74 to MAR40.85 |
| 2 | The charge is added to the requirement of its risk class before scaling | `SRC-02`, MAR40.2 and MAR40.75 |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Non-delta option charge | `option_charge` | Reporting currency | Risk class | None |

**Limits and non-goals:**

- The workbook calculates no option charge and checks none.
- Under the delta-plus method, the delta-equivalent position of an option
  is entered on the position sheets like any other position (MAR40.77 to
  MAR40.79). Under the simplified approach or the scenario approach, the
  option and its underlying are carved out: leave both out of the position
  sheets, and enter their requirement here.
- `Checks` reports the risk classes with a charge entered (rule `C13`).

**Failure behaviour:** a charge that is not a number, or is negative, is an
unusable parameter: the run reads `STOPPED`, and `Checks` lists the name of
the cell with `INVALID: must be a number` or `INVALID: must be 0 or more`.
An empty cell is zero.

**Validation:** the test case `illustrative-option-charges`. Its expected
values are `illustrative`: derived within the project, they validate nothing.
See the [validation matrix](VALIDATION_PLAN.md#validation-matrix) and the
rows of `CALC-07` in the [traceability
table](VALIDATION_PLAN.md#traceability).

### `CALC-08` Requirement and risk-weighted assets

**Purpose:** the own funds requirement for market risk under the simplified
standardised approach and its risk-weighted assets: the total row and the
headline of `Results`.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Requirement of each risk class | The outputs of `CALC-02` to `CALC-06` | Reporting currency |
| Option charges | The outputs of `CALC-07` | Reporting currency |
| Scaling factors and multiplier | `DS-02.scaling_ir`, `DS-02.scaling_eq`, `DS-02.scaling_co`, `DS-02.scaling_fx` and `DS-02.rwa_multiplier` | Number |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | The requirement of a risk class before scaling is the sum of its charges and of its option charge. For interest rate risk, the charges are specific risk and general market risk | `SRC-02`, MAR40.2 and MAR40.4 |
| 2 | The requirement after scaling is that sum multiplied by the scaling factor of the class: 1.30 for interest rate risk, 3.50 for equity risk, 1.90 for commodities risk and 1.20 for FX risk | `SRC-02`, MAR40.2 |
| 3 | The requirement is the simple sum of the four scaled requirements | `SRC-02`, MAR40.2 |
| 4 | The risk-weighted assets are the requirement multiplied by 12.5 | `SRC-02`, MAR40.1 |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Requirement before scaling | `requirement_before_scaling` | Reporting currency | Risk class | None |
| Requirement after scaling | `requirement_after_scaling` | Reporting currency | Risk class | None |
| Own funds requirement | `own_funds_requirement` | Reporting currency | Whole calculation | None |
| Risk-weighted assets | `risk_weighted_assets` | Reporting currency | Whole calculation | None |

**Limits and non-goals:**

- Whether a bank may use the simplified standardised approach is decided
  outside the workbook, and so is everything a jurisdiction's own rules add
  or change ([sources](SOURCES.md#decisions)).
- The requirement is that of the positions entered. The workbook holds no
  trading book boundary and no other part of the capital framework.
- When rows were rejected, the requirement is that of the rows used, and the
  status line reads `INCOMPLETE`.

**Failure behaviour:** an unusable value on `Params` stops the run, such as
a scaling factor that is not a number, an FX rate that is not greater than
zero, a currency listed twice or the reporting currency not listed with
rate 1. The run reads `STOPPED`, no result is calculated, and `Checks` lists
each value with its reason (rules `C05` to `C07` and `C14`). A position
sheet with more than 500 rows stops the run in the same way: `Checks` lists
the sheet and the rows found under `C14`.

**Validation:** the test cases `illustrative-requirement-fx-and-options-only`
and `illustrative-sample-requirement`. Their expected values are
`illustrative`: derived within the project, they validate nothing. See the
[validation matrix](VALIDATION_PLAN.md#validation-matrix) and the rows of
`CALC-08` in the [traceability table](VALIDATION_PLAN.md#traceability).

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| 2026-11-16 | The default tolerances are those the test-case guide suggests: `1e-6` absolute and `1e-9` relative for a derived value, and half a unit of the last printed digit for a published one | `Double` arithmetic agrees to far more digits on these sums and products, so a difference above the bound is a difference of method; a printed value is known only to its last digit | [20](https://github.com/danielep71/K-FRTB-SSA/issues/20) |
| 2026-11-16 | A value that `tests/derivations/ssa.py` computes is `illustrative`, and the script is not registered as an independent implementation | I wrote the script and the engine: the same reading of the text is in both, so their agreement cannot validate a rule. Only a value printed in `SRC-02`, or an implementation of another author, can | [20](https://github.com/danielep71/K-FRTB-SSA/issues/20) |
| 2026-11-16 | The calculations are numbered from `CALC-02` | `CALC-01` is the illustrative contract above, which the illustrative test cases still cite, and an ID is never reused | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-16 | A residual term is in calendar days over 365, a month is a twelfth of a year, and a term equal to the upper limit of a time band belongs to that band | `SRC-02` gives its bands in months and years, and names neither a day count nor the band of a limit. The prototype was built so | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-16 | The zones are offset in the order 1 and 2, 2 and 3, then 1 and 3 | MAR40.28 allows the offsetting between zones and gives its disallowances, not its order. The prototype was built so | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-16 | The user converts derivatives into positions, classifies issuers and calculates option charges | The workbook is a calculator of the charges of MAR40 on positions already expressed as the chapter requires. What it does not decide is listed as a limit of each contract | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
