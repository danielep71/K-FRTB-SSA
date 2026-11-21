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
- **Book:** a position row of `DS-06` to `DS-09` is in the trading book
  unless its `book` is `Non-trading`. A non-trading book position subject to
  foreign exchange or commodity risk is calculated as a trading book one by
  `CALC-05` and `CALC-06`, under both regimes; a non-trading book interest
  rate or equity position is outside the scope, is excluded by `CALC-03` and
  `CALC-04`, and `Checks` warns (rule `C17`). The run is not `INCOMPLETE` for
  it (`SRC-04`, Article 325(1); `SRC-11`, RBC25.28; `SRC-12`, MAR12.6). The
  technical standards of Article 325(9), `SRC-09` and `SRC-10`, value those
  positions for the alternative approaches only: under the simplified
  standardised approach the calculation at the value entered is an
  assumption.
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
| Specific-risk table, `SR_Key` and `SR_Charge`, 31 rows | The 18 charges of table 1, and 13 rows of the category `CRR`: 0%, half of the second category, the second category, 8% and 12% | `SRC-02`, MAR40.6, table 1; `SRC-04`, Article 336(1), table 1, and (3) | Equal, for the 18 rows. The 13 rows of the CRR were compared with Article 336 on 2026-11-20. The table gives the qualifying category one set of charges; the workbook lists it under the two investment grades. It gives *BB+ to B–* for governments and *Below BB–* for other issuers; the workbook lists each under the grades it spans |
| Time bands: `Ladder_Zone`, `Ladder_UpperHigh`, `Ladder_UpperLow`, `Ladder_RW`, 15 rows | The bands, zones and risk weights of table 4 | `SRC-02`, MAR40.26, table 4, and MAR40.28 for the zones | Equal. *1 month* is 1/12 of a year, and *3 months* 0.25 |
| `CouponThreshold` | 0.03 | `SRC-02`, MAR40.26 | Equal |
| `VerticalDisallowance` | 0.10 | `SRC-02`, MAR40.27 | Equal |
| `HorizontalZone1`, `HorizontalZone2`, `HorizontalZone3` | 0.40, 0.30, 0.30 | `SRC-02`, MAR40.28, table 5 | Equal |
| `HorizontalAdjacent`, `HorizontalZone13` | 0.40, 1.00 | `SRC-02`, MAR40.28, table 5 | Equal |
| `HorizontalZone13CRR` | 1.50 | `SRC-04`, Article 339(9), point (f) | Equal, on 2026-11-20 |
| `EQSpecificRate`, `EQGeneralRate` | 0.08, 0.08 | `SRC-02`, MAR40.43 | Equal |
| `EQIndexRate` | 0.02 | `SRC-02`, MAR40.47 | Equal |
| `EQIndexRateCRR` | 0 | `SRC-04`, Article 344 | Equal, on 2026-11-20 |
| `FXChargeRate` | 0.08 | `SRC-02`, MAR40.61 | Equal |
| `COMDirectionalRate` | 0.15 | `SRC-02`, MAR40.72 | Equal |
| `COMGrossRate` | 0.03 | `SRC-02`, MAR40.73 | Equal |
| `TransitionalMultiplier` | 0.9 | `SRC-06`, Article 495s(2) | Equal, on 2026-11-20 |
| `TransitionalEndDate` | 2029-12-31 | `SRC-06`, Article 495s(2) | Equal, on 2026-11-20 |
| `TotalOwnFunds` | None: the user's value, in the reporting currency | None: an input of the CRR regime | Not applicable |
| `FXDeMinimisRate` | 0.02 | `SRC-04`, Article 351 | Equal, on 2026-11-20 |
| `EligibilityShare325a`, `EligibilityCap325a` | 0.10, EUR 500,000,000 | `SRC-04`, Article 325a(1) | Equal, on 2026-11-20 |
| `EligibilityShare94`, `EligibilityCap94` | 0.05, EUR 50,000,000 | `SRC-04`, Article 94(1) | Equal, on 2026-11-20 |
| `Regime` | `Basel` | None: the regime the user chooses, `Basel` or `CRR` | Not applicable |
| `DaysPerYear` | 365 | None: an assumption of the workbook | Not applicable |
| `GoldCode` | `XAU` | None: the code under which gold is entered, a convention of the workbook. MAR40.53 treats gold as a currency | Not applicable |

`TransitionalMultiplier` and `TransitionalEndDate` are not values of MAR40: they
belong to the CRR regime, and `SRC-06` gives them; `FXDeMinimisRate` and
`HorizontalZone13CRR` and `EQIndexRateCRR` belong to it too, and `SRC-04` gives them. `Regime` selects between the
two regimes ([#74](https://github.com/danielep71/K-FRTB-SSA/issues/74)). Rule
`C08` counts `TransitionalMultiplier`, `TransitionalEndDate`, `FXDeMinimisRate`,
`HorizontalZone13CRR` and `EQIndexRateCRR` among the values that differ from their reference, as for the others, whatever the
regime.

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
| 5 | Under the `CRR` regime a position's category is `CRR` (or `NONE`), and its grade the risk weight that the issuer or the security would receive under the standardised approach for credit risk, `RW 0%`, `RW 10%`, `RW 20%`, `RW 50%`, `RW 75%`, `RW 100%` or `RW 150%`. The rate is 0% at `RW 0%`; 0.25%, 1.00% and 1.60% at `RW 20%` and `RW 50%`, by the three ranges of the residual term of step 3; half of those, 0.125%, 0.50% and 0.80%, at `RW 10%`, the covered bonds of Article 336(3); 8% at `RW 75%` and `RW 100%`; and 12% at `RW 150%`. The rest of steps 1, 2 and 4 is as under the Basel regime | `SRC-04`, Article 336(1), table 1, and (3) |

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
- A weight of 75%, which table 1 of Article 336 does not name, is charged as
  the category of 100%, 8%, by my decision of 2026-11-20
  ([#79](https://github.com/danielep71/K-FRTB-SSA/issues/79)). The text of the
  CRR does not say which category applies: this is an assumption.
- The user places a position in a category of Article 336 by entering the risk
  weight, which the workbook does not derive from an exposure class, a credit
  quality step or a rating. A security of the other qualifying items of
  Article 336(4) is entered at `RW 50%`, since the article gives it the
  category of 20% and 50%. The condition of Article 336(2) for an institution on
  the IRB approach, and the notification of the methodology of Article 336(4),
  are the institution's.
- Holdings of the institution's own debt instruments are disregarded for
  specific risk (Article 327(1)): enter them with category `NONE`, which keeps
  their general risk. A derivative leg is `NONE`, as under the Basel regime:
  Articles 328 to 330 do not say in terms that a leg carries specific risk, and
  this is an assumption.
- Not supported under the `CRR` regime: the cap of Article 335 at the maximum
  possible default-risk loss, which is optional, as MAR40.15 is, and which needs
  that loss as an input.

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
| Disallowances | `DS-02.vertical_disallowance` and the six `DS-02.horizontal_` fields; `DS-02.regime` | Fraction; text |
| Book | `DS-06.book` | Text |

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
| 10 | Under the `CRR` regime, the disallowance on the offset between zones 1 and 3 of step 7 is 150%, `HorizontalZone13CRR`, in place of 100%. The rest of steps 1 to 9 is as under the Basel regime | `SRC-04`, Article 339(9), point (f) |

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
  step 7, which MAR40.28 does not state. Under the `CRR` regime the order of
  step 7 is that of Article 339(5) to (7); paragraph 6 lets the institution
  reverse the first two offsets, and the workbook does not offer that choice.
- A derivative is entered as its legs: the user converts it into positions
  in the underlying, as MAR40.31 to MAR40.34 and table 7 of MAR40.40
  describe, and enters each leg as a row. The workbook converts nothing.
- The workbook does not find the matched positions that MAR40.35 and
  MAR40.36 allow a bank to leave out: leave them out of the input, or they
  are slotted like any other.
- Not supported: the duration method (MAR40.29), and under the `CRR` regime
  the duration-based calculation of Article 340, which needs the yield and
  the modified duration of each position as inputs and which the institution
  may choose instead of Article 339; a single ladder for
  currencies in which business is insignificant (MAR40.24 and MAR40.30);
  the alternative formulae for swaps (MAR40.37); and the national
  discretion of MAR40.13 to disallow offsetting.
- Gold and the reporting of FX positions belong to `CALC-06`. A row whose
  currency is the gold code is rejected.

**Failure behaviour:** `ir_position_status`, on `IRCalc`, is `OK` or:

- `INVALID: duplicate PositionID`
- `INVALID: Book must be Trading or Non-trading`
- `EXCLUDED: non-trading book position outside the market risk scope`
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
- `INVALID: SR category must be CRR or NONE under the CRR`
- `INVALID: no Article 336 charge for this risk weight`

The last two are the texts under the `CRR` regime: the first replaces the test of
the category, which then accepts `CRR` and `NONE` only, and the second replaces
the text before them for a category and grade that the table lacks. Under the
Basel regime the category `CRR` is rejected with the text that lists `GOV`,
`QUAL`, `OTHER` and `NONE`.

The first test that fails gives the text. A rejected row enters neither
`CALC-02` nor this calculation. Nor does an excluded row, which is a
non-trading book position: it is excluded before any test after the book, so
its other fields need not be valid, and `Checks` warns (rule `C17`) without
listing it.

**Validation:** the test cases `bcbs119-c2-maturity-method`,
`illustrative-invalid-ir-positions`, `illustrative-ir-time-band-limits`,
`illustrative-sample-ir-general-market-risk` and
`mar40-27-vertical-disallowance`. Two of them hold `published` values: the
weighted positions, the vertical and within-zone disallowances, the net
position charge and the charge of the worked example C.2 of `SRC-03`
(`bcbs119-c2-maturity-method`), and the weighted positions and the vertical
disallowance of the example of MAR40.27 (`mar40-27-vertical-disallowance`).
The expected values of the others are `illustrative`: derived within the
project, they validate nothing. No test case has an expected value under the
`CRR` regime, so step 10 is exercised by the VBA suite and by `verify.py`
only. Run under the `CRR` regime, the worked example C.2 gives 1,950,000 between
the zones and a charge of 5,080,000, which is 500,000 more than its published
values: 50% of the 1,000,000 that it offsets between zones 1 and 3, as step 10
says; the other components are unchanged (`tests/test_prototype.py`). See the
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
| Rates | `DS-02.eq_specific_rate`, `DS-02.eq_index_rate`, `DS-02.eq_index_rate_crr` and `DS-02.eq_general_rate`; `DS-02.regime` | Fraction; text |
| Book | `DS-07.book` | Text |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | Positions in the identical equity or stock index in a market are netted: the rows with the same market, type and identifier, without regard to case | `SRC-02`, MAR40.46 |
| 2 | Specific risk is 8% of the absolute net position of each single name | `SRC-02`, MAR40.42 and MAR40.43 |
| 3 | The charge for an index is 2% of the absolute net position of each index, in place of specific risk | `SRC-02`, MAR40.47, and table 8 of MAR40.52 |
| 4 | General market risk is 8% of the absolute net position of each market: the sum of its longs less the sum of its shorts, single names and indices together | `SRC-02`, MAR40.42 and MAR40.43 |
| 5 | The requirement is the sum of steps 2 to 4 over the stocks, the indices and the markets | `SRC-02`, MAR40.42 |
| 6 | Under the `CRR` regime the charge of step 3 is `EQIndexRateCRR`, 0, in place of 2%: a stock-index future is treated as one equity, whose general risk is in the net position of its market (step 4) and whose specific risk is ignored where the future is exchange traded and represents a relevant, appropriately diversified index. A position that does not meet that condition is entered as a `Single name`, and step 2 charges its specific risk. The rest of steps 1 to 5 is as under the Basel regime | `SRC-04`, Article 344 |

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
  index is one. Under the `CRR` regime every position entered as `Index` is
  charged 0: the user enters as `Index` only a stock-index future that is
  exchange traded and represents a relevant, appropriately diversified index,
  and any other as a `Single name`. The breakdown of a future into its
  constituent equities, which Article 344 allows with a notification to the
  competent authority, is the user's: enter the constituents as `Single name`
  rows. The notification is the institution's, and is not calculated.
- Not supported: the arbitrage strategies of MAR40.48 to MAR40.50, and the
  offset of depository receipts (MAR40.51); under the `CRR` regime the
  collective investment undertakings of Article 348, with 32% for position
  risk and 40% with foreign exchange risk, whose Articles 348 to 350 were not
  read in full and which would need a type of position, an input for the
  foreign exchange risk and an output of their own.
- A market is the national market that you enter. Under the `CRR` regime,
  Article 1 of `SRC-07` (Delegated Regulation (EU) No 525/2014), made under
  Article 341(3) of the CRR, defines it otherwise: for the euro area, all
  equities listed in stock markets located in Member States that have adopted
  the euro; for non-euro Member States and third countries, all equities
  listed in stock markets located within a national jurisdiction. Enter one
  market for all the equities listed in the euro area, and one for each
  national jurisdiction elsewhere. The workbook nets by the market you enter:
  it does not know where an equity is listed, nor which Member States have
  adopted the euro.
- The interest rate or FX exposure of the other leg of a contract (footnote
  14 of MAR40.44) is entered on `IRPositions` or `FXPositions`.
- `Buckets` lists up to 30 markets. The totals include every market, and
  `Checks` warns (rule `C10`).
- A non-trading book equity row is excluded. Not supported: a short equity
  position of the banking book that is not deducted or risk-weighted there,
  which RBC25.24 of `SRC-11` puts under the market risk rules: entered as
  `Non-trading`, it is excluded like any other.

**Failure behaviour:** `eq_position_status`, on `EQCalc`, is `OK` or:

- `INVALID: duplicate PositionID`
- `INVALID: Book must be Trading or Non-trading`
- `EXCLUDED: non-trading book position outside the market risk scope`
- `INVALID: Market missing`
- `INVALID: Type must be Single name or Index`
- `INVALID: stock or index identifier missing`
- `INVALID: currency not in FX_Table`
- `INVALID: Direction must be Long or Short`
- `INVALID: Market value must be a number`
- `INVALID: Market value must be >= 0 (Direction gives the sign)`

As under `CALC-03`, an excluded row is a non-trading book position, excluded
before any test after the book, and `Checks` warns (rule `C17`).

**Validation:** the test cases `illustrative-eq-matched-positions`,
`illustrative-invalid-eq-positions` and `illustrative-sample-equity`. Their
expected values are `illustrative`: derived within the project, they validate
nothing. No test case has an expected value under the `CRR` regime, so step 6
is exercised by the VBA suite and by `verify.py` only. See the [validation matrix](VALIDATION_PLAN.md#validation-matrix)
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
| Book | `DS-08.book` | Text |

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
- Under the `CRR` regime the workbook calculates the simplified approach of
  Article 360 only, with the rates of steps 3 and 4, which are the CRR's. It
  does not offer the choice that Article 355 gives the institution among
  Articles 359, 360 and 361. Not supported: the maturity ladder of Article 359
  (a spread rate of 1.50% in each of seven bands, a carry rate of 0.6% applied
  for each band into which an unmatched position is carried forward, and an
  outright rate of 15%), because paragraph 4 of that article does not say in
  which order an unmatched position is matched against the bands further out
  when several are open, and the order changes the matched positions to which
  the rates apply; the extended maturity ladder of Article 361 (the rates of
  its Table 2 by category of commodity, for an institution with significant
  commodities business and an appropriately diversified portfolio, which
  notifies the use), which has the same gap and needs the category of the
  commodity as an input; and the ancillary agricultural commodities business of
  Article 356, which the institution determines for its physical stock once a
  year, on conditions that it monitors, and notifies. A ladder also needs
  outputs of its own on `Buckets` and `Results`.
- Under the `CRR` regime Article 357(4) treats as positions in the same
  commodity the positions in different sub-categories that are deliverable
  against each other, and those in similar commodities that are close
  substitutes, where a correlation of at least 0.9 between price movements is
  clearly established over at least one year. Enter them under one commodity
  name: the workbook tests none of those conditions. Under Article 357(5) the
  commodities that the institution transferred in a repurchase agreement or
  lent are included: enter them as positions.
- Two rows are one commodity only when their names are equal. The workbook
  does not check that their units are.
- A derivative is entered as a notional position in the commodity, as
  MAR40.71 describes. The workbook converts nothing.
- Gold is FX risk (MAR40.63): enter it on `FXPositions`. A row whose
  commodity is the gold code is rejected.
- A non-trading book position is calculated as a trading book one, at the
  spot price entered: an assumption. Article 2 of `SRC-09` values it at the
  latest available fair value, measured at least monthly, under the
  alternative standardised approach only.
- A spot price is used as entered, also when it is negative.
- `Buckets` lists up to 30 commodities. The totals include every commodity,
  and `Checks` warns (rule `C11`).

**Failure behaviour:** `co_position_status`, on `COCalc`, is `OK` or:

- `INVALID: duplicate PositionID`
- `INVALID: Book must be Trading or Non-trading`
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
| Regime, total own funds and de minimis | `DS-02.regime`, `DS-02.total_own_funds` and `DS-02.fx_de_minimis_rate` | Text; reporting currency; fraction |
| Book | `DS-09.book` | Text |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | The net position in each foreign currency and in gold is the sum of its rows, converted at the spot rate into the reporting currency | `SRC-02`, MAR40.60 |
| 2 | The overall net open position is the greater of the sum of the net long positions and the sum of the net short positions in the currencies, plus the net position in gold, whatever its sign | `SRC-02`, MAR40.60 |
| 3 | The requirement is 8% of the overall net open position | `SRC-02`, MAR40.61 |
| 4 | Under the `CRR` regime, the requirement is zero where the overall net open position of step 2 does not exceed `FXDeMinimisRate`, 2%, of `TotalOwnFunds`, and that of step 3 where it exceeds it. Under the Basel regime there is no such threshold | `SRC-04`, Article 351 |

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
- Gold under the `CRR` regime is treated as MAR40.60 says, added to the greater
  of the two sums whatever its sign: an assumption. Article 352(4) of the CRR
  puts the net position in gold in the totals of long and short positions, and
  Article 351 adds the net gold position to the overall net foreign exchange
  position; the text read does not say which of the two readings is meant
  ([comparison](CRR_COMPARISON.md#foreign-exchange-risk-calc-06)).
- Not supported under the `CRR` regime: the lower requirements for closely
  correlated currencies (Article 354), whose list of currencies is left to
  technical standards and whose use needs a statistical test and a permission;
  and the exclusion of a position that hedges a capital ratio (Article 104c)
  or is deducted from own funds (Article 325(1)), which needs a permission.
  Leave such a position out of `FXPositions`: the workbook takes the net
  positions as entered.
- The de minimis needs `TotalOwnFunds` under the `CRR` regime, and the
  requirement of step 3 is the whole requirement above the threshold, not the
  excess over it.
- A non-trading book position is calculated as a trading book one, at the
  net position entered: an assumption. Article 1 of `SRC-09` values it at the
  last available accounting value, or fair value, updated at least monthly
  for the exchange rates, under the alternative standardised approach only;
  the valuation is the user's, as for any net position entered.
- Article 495u of `SRC-06` lets an institution eligible for the derogation of
  Article 94 calculate its non-trading book positions subject to foreign
  exchange or commodity risk by Article 325(2) until 2029-12-31. The workbook
  calculates every position so: it has no parameter for the article, does not
  know whether the institution uses the derogation, and does not check the end
  date. Not supported beyond that date: the alternative standardised approach
  that `SRC-09` specifies.

**Failure behaviour:** `fx_position_status`, on `FXCalc`, is `OK` or:

- `INVALID: duplicate PositionID`
- `INVALID: Book must be Trading or Non-trading`
- `INVALID: currency not in FX_Table`
- `INVALID: Direction must be Long or Short`
- `INVALID: Amount must be a number`
- `INVALID: Amount must be >= 0 (Direction gives the sign)`
- `EXCLUDED: reporting currency carries no FX risk`

Under the `CRR` regime, a `TotalOwnFunds` that is not a number greater than 0
is an unusable parameter: the run reads `STOPPED`, and `Checks` lists it with
`INVALID: must be a number greater than 0`.

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
| 3 | Under the `CRR` regime the user calculates the charge by one of the approaches of Delegated Regulation (EU) No 528/2014: the simplified approach, the delta-plus approach or the scenario approach. Whatever the approach, the delta-equivalent position of every option is entered on the position sheets, because the simplified approach deducts the risk-weighted delta equivalent amount from the gross amount, and the scenario approach deducts the delta effect, so the charge entered is net of the delta | `SRC-08`, Articles 1 to 9 and Annexes I and II |

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
  sheets, and enter their requirement here. Under the `CRR` regime, do not
  carve out: enter the delta-equivalent position of every option on the
  position sheets, and enter here the charge of Delegated Regulation (EU) No
  528/2014, which has the delta deducted (step 3). Entering a carved-out
  charge under the `CRR` regime counts the delta effect twice.
- Under the `CRR` regime the approaches have the conditions of Articles 1, 2
  and 7 of `SRC-08`: only an institution that purchases options and warrants
  alone may use the simplified approach; on an individual basis only the
  scenario approach and the delta-plus approach may be combined; the scenario
  approach needs a risk control unit, a notified scope and internal reporting.
  They are the institution's, and the workbook checks none. The delta that
  Article 329(1) of the CRR takes is the exchange's, or the institution's own
  model with the permission of the competent authority: enter it as given.
- Under the `CRR` regime the equity underlying of an option is each market as
  `SRC-07` defines it, so the euro area is one market, and the commodity
  underlying is each commodity that Article 357(4) treats as the same.
- The charge is added before scaling, under the `CRR` regime as under
  Basel, by a reading of Article 325(2) of the CRR: it multiplies the
  requirements of Chapters 2 to 4 by 1.3, 3.5, 1.2 and 1.9, and Articles
  329(2), 352(5) and 358(3), which require the other risks of options to be
  reflected, are in those chapters. The text says no more: this is an
  assumption.
- `Checks` reports the risk classes with a charge entered (rule `C13`).

**Failure behaviour:** a charge that is not a number, or is negative, is an
unusable parameter: the run reads `STOPPED`, and `Checks` lists the name of
the cell with `INVALID: must be a number` or `INVALID: must be 0 or more`.
An empty cell is zero.

**Validation:** the test case `illustrative-option-charges`. Its expected
values are `illustrative`: derived within the project, they validate nothing.
Step 3 is an instruction to the user and has no test.
See the [validation matrix](VALIDATION_PLAN.md#validation-matrix) and the
rows of `CALC-07` in the [traceability
table](VALIDATION_PLAN.md#traceability).

### `CALC-08` Requirement and risk-weighted assets

**Purpose:** the own funds requirement for market risk under the simplified
standardised approach and its risk-weighted assets: the total row, the row of
the transitional multiplier and the headline of `Results`.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Requirement of each risk class | The outputs of `CALC-02` to `CALC-06` | Reporting currency |
| Option charges | The outputs of `CALC-07` | Reporting currency |
| Scaling factors and multiplier | `DS-02.scaling_ir`, `DS-02.scaling_eq`, `DS-02.scaling_co`, `DS-02.scaling_fx` and `DS-02.rwa_multiplier` | Number |
| Regime | `DS-02.regime` | Text: `Basel` or `CRR` |
| Transitional multiplier | `DS-02.transitional_multiplier`, `DS-02.transitional_end_date` and `DS-02.as_of_date` | Factor; date |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | The requirement of a risk class before scaling is the sum of its charges and of its option charge. For interest rate risk, the charges are specific risk and general market risk | `SRC-02`, MAR40.2 and MAR40.4 |
| 2 | The requirement after scaling is that sum multiplied by the scaling factor of the class: 1.30 for interest rate risk, 3.50 for equity risk, 1.90 for commodities risk and 1.20 for FX risk | `SRC-02`, MAR40.2 |
| 3 | The requirement is the simple sum of the four scaled requirements | `SRC-02`, MAR40.2 |
| 4 | The risk-weighted assets are the requirement multiplied by 12.5 | `SRC-02`, MAR40.1 |
| 5 | The requirement of step 3 is multiplied by a transitional factor, and step 4 applies to the result. The factor is `TransitionalMultiplier`, 0.9, when the regime is `CRR` and the reporting date is on or before `TransitionalEndDate`, 2029-12-31, and 1 under the Basel regime and after that date | `SRC-06`, Article 495s(2) |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Requirement before scaling | `requirement_before_scaling` | Reporting currency | Risk class | None |
| Requirement after scaling | `requirement_after_scaling` | Reporting currency | Risk class | None |
| Transitional factor | `transitional_factor` | Factor | Whole calculation | None |
| Own funds requirement, after the transitional factor | `own_funds_requirement` | Reporting currency | Whole calculation | None |
| Risk-weighted assets | `risk_weighted_assets` | Reporting currency | Whole calculation | None |

**Limits and non-goals:**

- Whether a bank may use the simplified standardised approach is decided
  outside the workbook, and so is everything a jurisdiction's own rules add
  or change ([sources](SOURCES.md#decisions)).
- The requirement is that of the positions entered. The workbook holds no
  trading book boundary and no other part of the capital framework.
- The CRR regime changes step 5 and nothing else for now: its other rules are
  those of MAR40 until each difference listed in the
  [comparison](CRR_COMPARISON.md) is implemented.
- After `TransitionalEndDate` the factor is 1 and no finding is made: the
  transitional treatment has ended, and the requirement of step 3 is the
  result. `TransitionalEndDate` is read under the `CRR` regime only.
- When rows were rejected, the requirement is that of the rows used, and the
  status line reads `INCOMPLETE`.

**Failure behaviour:** an unusable value on `Params` stops the run, such as
a scaling factor that is not a number, an FX rate that is not greater than
zero, a currency listed twice or the reporting currency not listed with
rate 1. The run reads `STOPPED`, no result is calculated, and `Checks` lists
each value with its reason (rules `C05` to `C07` and `C14`), among them a
`Regime` that is neither `Basel` nor `CRR` and, under the `CRR`, a
`TransitionalEndDate` that is not a date. A position
sheet with more than 500 rows stops the run in the same way: `Checks` lists
the sheet and the rows found under `C14`.

**Validation:** the test cases `illustrative-requirement-fx-and-options-only`
and `illustrative-sample-requirement`. Their expected values are
`illustrative`: derived within the project, they validate nothing. See the
[validation matrix](VALIDATION_PLAN.md#validation-matrix) and the rows of
`CALC-08` in the [traceability table](VALIDATION_PLAN.md#traceability).

### `CALC-09` Eligibility for the simplified standardised approach and the small trading book derogation

**Purpose:** under the `CRR` regime, whether the size of the institution's business
allows the simplified standardised approach, Article 325a, and the derogation for
small trading book business, Article 94, and the months that give the finding. The
result is a finding on the `Eligibility` sheet and on `Checks`; it does not change
any requirement.

**Inputs:**

| Input | Dataset and field | Unit |
| --- | --- | --- |
| Regime | `DS-02.regime` | Text: `Basel` or `CRR` |
| Month, total assets and the aggregated long and short positions of each approach | `DS-11.month_end`, `DS-11.total_assets`, `DS-11.long_325a`, `DS-11.short_325a`, `DS-11.long_94` and `DS-11.short_94` | Date; euro |
| Share of total assets and amount, for each approach | `DS-02.eligibility_share_325a`, `DS-02.eligibility_cap_325a`, `DS-02.eligibility_share_94` and `DS-02.eligibility_cap_94` | Fraction; euro |

**Rules:**

| Step | Rule | Source |
| --- | --- | --- |
| 1 | The size of the business of a month is the absolute value of the aggregated long position plus the absolute value of the aggregated short position, for each approach. The user enters the two aggregates, valued as the Article says | `SRC-04`, Article 325a(2), point (f); Article 94(3), point (c) |
| 2 | The conditions are met in a month where the size is equal to or less than both the share of the total assets and the amount: 10% and EUR 500 million for the simplified standardised approach, 5% and EUR 50 million for the derogation. The size is compared with the share multiplied by the total assets, and with the amount | `SRC-04`, Article 325a(1); Article 94(1) |
| 3 | The run of a month is the number of consecutive months entered, that month included, in which the conditions are not met | `SRC-04`, Article 325a(5), point (a); Article 94(7), point (a) |
| 4 | The count of a month is the number of months, among the last 12 up to and including that month, in which the conditions are not met. Before the twelfth month entered, the months entered are counted | `SRC-04`, Article 325a(5), point (b); Article 94(7), point (b) |
| 5 | A month is an event where its run is 3 or more, or its count is more than 6 | `SRC-04`, Article 325a(5); Article 94(7) |
| 6 | The status is that of the last month entered: `CEASE` where that month is an event, otherwise `NOT MET` where the conditions are not met in it, otherwise `MET` | `SRC-04`, Article 325a(5); Article 94(7) |
| 7 | The restart condition is `YES` where 12 months are entered and the conditions are met in each of the last 12, otherwise `NO` | `SRC-04`, Article 325a(6); Article 94(8) |
| 8 | The assessment is made under the `CRR` regime only. Under the Basel regime the sheet is not read, no row is tested and no cell of the rows or of the figures of the result block is written, and the status is `NOT ASSESSED: Basel regime` | `SRC-04`, Article 325a |

**Outputs:**

| Output | Quantity | Unit | Scope | Rounding |
| --- | --- | --- | --- | --- |
| Size of the business | `eligibility_size` | Euro | Month, and approach | None |
| Share of total assets | `eligibility_share` | Fraction | Month, and approach | None |
| Conditions of the month | `eligibility_month_result` | `MET` or `NOT MET` | Month, and approach | None |
| Run | `eligibility_run` | Number of months | Month, and approach | None |
| Count of the last 12 months | `eligibility_count` | Number of months | Month, and approach | None |
| Event | `eligibility_event` | `YES` or `NO` | Month, and approach | None |
| Status | `eligibility_status` | Text | Approach | None |
| Restart condition | `eligibility_restart` | `YES` or `NO` | Approach | None |

An approach is `325a`, the simplified standardised approach, or `94`, the
derogation. The quantity of the expected files gives it after the name:
`eligibility_status_325a`.

**Limits and non-goals:**

- The user enters the aggregated long and short positions of each month. The
  workbook does not value positions, does not decide which positions Article
  325a(2) or Article 94(3) includes, and does not decide whether a position is
  long or short: the side is the side as entered, until the regulatory technical
  standard of Article 94(10) is registered ([sources](SOURCES.md)). The positions of
  Article 94(3) exclude foreign exchange and commodity positions, which those of
  Article 325a(2) include: the user enters two sets of totals.
- The amounts are in euro, as the Articles give them, and the workbook converts
  none: the share of the total assets does not depend on the currency, and the
  amounts are compared with the euro thresholds as entered. This is an assumption of the workbook.
- A month is the last day of a calendar month. The months entered follow one
  another without a gap; a month that is not entered is not assumed to meet or to
  miss the conditions.
- The sheet holds 24 rows, from row 5 to row 28, and reads each of them in
  order. A row whose month-end is empty is skipped, whatever its other cells
  hold, and the row entered before it is the reference of the next row for the
  test of one month after the previous row. The run and the count go over the
  valid rows, in the order of the sheet. A month in which an approach has no
  position is entered with its totals at 0: a blank total is not a number.
- Where a status reads `NOT ASSESSED`, the result block shows no month, size,
  share, run, count or restart condition of that approach, and a row that is
  valid still shows its figures, except under the Basel regime.
- Only the last month entered is judged. A month earlier in the series that is an
  event stays visible in its row, and no status of an approach already
  ceased or restarted is kept: the user decides which approach is in use.
- The notifications of Article 325a(3) and (4) and of Article 94(5) and (6), the
  obligation to cease within three months, the permission to restart of Article
  325a(6) and the prohibition of Article 325a(7) and Article 94(9) are the
  institution's obligations. The workbook documents them and calculates none.
- The status changes no requirement and stops no run: an institution that
  must cease the approach reads the finding and acts outside the workbook.
- The derogation of Article 94(2), its simpler requirement and its relief,
  is not calculated; for Article 495u of Delegated Regulation (EU) 2026/1221,
  see `CALC-06`.

**Failure behaviour:** `eligibility_month_status`, on `Eligibility`, is `OK` or:

- `INVALID: MonthEnd must be a date`
- `INVALID: MonthEnd must be the last day of a month`
- `INVALID: MonthEnd must be one month after the previous row`
- `INVALID: TotalAssets must be a number greater than 0`
- `INVALID: a long or short total must be a number`

A month is tested in the order of this list, and the first test that fails gives
the status. The test of one month after the previous row applies from the second row
entered, against the month-end of the row entered before it when that is a date.

A row that fails a test contributes nothing, `Checks` counts it (rule `C15`)
and the status line reads `INCOMPLETE`; the reason stays in the status of the row.
While a row is invalid, both statuses read `NOT ASSESSED: a row is invalid`. With no
month entered, they read `NOT ASSESSED: no month entered`. A status `NOT MET` or
`CEASE` is a warning of `Checks`, counted in rule `C16` once for each approach, and
the requirement is unchanged.

**Validation:** five registered test cases, `illustrative-eligibility-…` and
`illustrative-invalid-eligibility-months`, hold `illustrative` values that
`tests/derivations/ssa.py` derives from this contract; no `published` value covers
it yet. See the [validation matrix](VALIDATION_PLAN.md#validation-matrix) and the
rows of `CALC-09` in the [traceability table](VALIDATION_PLAN.md#traceability).

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| 2026-11-16 | The default tolerances are those the test-case guide suggests: `1e-6` absolute and `1e-9` relative for a derived value, and half a unit of the last printed digit for a published one | `Double` arithmetic agrees to far more digits on these sums and products, so a difference above the bound is a difference of method; a printed value is known only to its last digit | [20](https://github.com/danielep71/K-FRTB-SSA/issues/20) |
| 2026-11-16 | A value that `tests/derivations/ssa.py` computes is `illustrative`, and the script is not registered as an independent implementation | I wrote the script and the engine: the same reading of the text is in both, so their agreement cannot validate a rule. Only a value printed in `SRC-02`, or an implementation of another author, can | [20](https://github.com/danielep71/K-FRTB-SSA/issues/20) |
| 2026-11-16 | The calculations are numbered from `CALC-02` | `CALC-01` is the illustrative contract above, which the illustrative test cases still cite, and an ID is never reused | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-16 | A residual term is in calendar days over 365, a month is a twelfth of a year, and a term equal to the upper limit of a time band belongs to that band | `SRC-02` gives its bands in months and years, and names neither a day count nor the band of a limit. The prototype was built so | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-16 | The zones are offset in the order 1 and 2, 2 and 3, then 1 and 3 | MAR40.28 allows the offsetting between zones and gives its disallowances, not its order. The prototype was built so | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-16 | The user converts derivatives into positions, classifies issuers and calculates option charges | The workbook is a calculator of the charges of MAR40 on positions already expressed as the chapter requires. What it does not decide is listed as a limit of each contract | [18](https://github.com/danielep71/K-FRTB-SSA/issues/18) |
| 2026-11-21 | A position row has a book, and a blank book is the trading book. Under both regimes the non-trading book rows of `CALC-05` and `CALC-06` are calculated as trading book ones, at the value entered, and those of `CALC-03` and `CALC-04` are excluded, with a warning (`C17`); the book is tested after the PositionID, so an excluded row needs no other valid field. Article 495u is documented, not a parameter | Article 325(1) of the CRR, RBC25.28 and MAR12.6 put the foreign exchange and commodity risk of the non-trading book in the scope, and only that risk. `SRC-09` and `SRC-10` value those positions for the alternative approaches only, so the value entered is an assumption. An excluded row is out of the scope, not invalid. The workbook calculates Article 325(2) for every position, which is what Article 495u allows | [85](https://github.com/danielep71/K-FRTB-SSA/issues/85) |
