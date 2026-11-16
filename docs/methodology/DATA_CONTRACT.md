# 🧾 Data contract

This contract defines every input the workbook accepts: where it comes from,
its layout and units, what makes it invalid, and how an import is traced.
The [calculation contracts](CALCULATION_CONTRACTS.md) refer to its datasets
and fields by name.

The workbook imports no file. Every dataset is typed on a sheet of the
workbook: the parameters and the three tables of `Params`, the four position
sheets and `OptionCharges`. The [workbook contract](WORKBOOK_CONTRACT.md)
gives the place of each on its sheet.

## Datasets

This table and the field tables are rendered from the
[machine-readable contract](#machine-readable-contract).

<!-- BEGIN generated: data-contract-datasets -->
| ID | Description | Provider | Frequency | Granularity | File or sheet | Encoding |
| --- | --- | --- | --- | --- | --- | --- |
| *Illustrative:* `DS-01` | Example positions | Upstream system or manual input | Per reporting date | One row per position | `positions.csv` | `UTF-8` |
| `DS-02` | Parameters | Typed by the user | Per reporting date for the first two fields; the others are the reference values of SRC-02 unless a what-if changes them | One record: the Value column of the parameter list, one field per row | `Params` |  |
| `DS-03` | Specific-risk charges | Typed by the user | Fixed: the reference values of SRC-02 unless a what-if changes them | One row per category, rating grade and range of residual maturity; 18 rows | `Params` |  |
| `DS-04` | Maturity ladder | Typed by the user | Fixed: the reference values of SRC-02 unless a what-if changes them | One row per time band; 15 rows | `Params` |  |
| `DS-05` | FX rates | Typed by the user | Per reporting date | One row per currency, and one for gold; up to 20 rows | `Params` |  |
| `DS-06` | Interest rate positions | Typed by the user | Per reporting date | One row per debt position or derivative leg; up to 500 rows | `IRPositions` |  |
| `DS-07` | Equity positions | Typed by the user | Per reporting date | One row per position in a stock or an index; up to 500 rows | `EQPositions` |  |
| `DS-08` | Commodity positions | Typed by the user | Per reporting date | One row per position in a commodity; up to 500 rows | `COPositions` |  |
| `DS-09` | FX positions | Typed by the user | Per reporting date | One row per net position in a currency or in gold; up to 500 rows | `FXPositions` |  |
| `DS-10` | Non-delta option charges | Typed by the user | Per reporting date | One row per risk class; 4 rows | `OptionCharges` |  |
<!-- END generated: data-contract-datasets -->

A name with an extension is a file; one without is a sheet. A file is one
dataset; a sheet may hold several, one for each of its tables. *Encoding* is the
encoding the workbook reads a file in, as a Python codec name such as `UTF-8`
or `cp1252`. For a file, an empty cell (`null` in the contract) means UTF-8; a
sheet has no encoding, and its cell is always empty. The
[file-based test cases](TEST_CASES.md#file-based-cases) hold only these files.
Outside an `invalid` case, each decodes in its dataset's encoding; an
`invalid` case may hold one that does not, because that is what it tests.

## File format

No dataset of this workbook is a file, so this section and the
[reference validator](#reference-validator) apply only to the illustrative
`DS-01`, until a dataset is read from a file.

State, for each delimited text file the workbook imports, its encoding,
delimiter and decimal separator in `data-contract.json`. The rest of the
format is fixed, so that the [reference validator](#reference-validator) and
the workbook read a file the same way. In outline:

- dates are ISO 8601, `YYYY-MM-DD`;
- the first row is a header that names each column by its field name, and
  the columns may come in any order;
- a missing value is an empty cell, kept distinct from zero;
- a cell may be quoted with double quotes.

The [reference validator](#reference-validator) section gives the complete
rules, including how numbers and booleans are written, the line endings, a
blank line and a byte order mark.

A provider file in another format is converted before it reaches the
workbook, and the conversion is recorded as a *Decision* below.

## Fields

One table per dataset, after the file or sheet it is read from:

<!-- BEGIN generated: data-contract-fields -->
### *Illustrative:* `DS-01` Example positions

Read from `positions.csv`, delimited text in `UTF-8`, with `,` between cells and `.` as the decimal separator.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| *Illustrative:* `position_id` | Text | — | Yes | Any | Identifies the position |
| *Illustrative:* `currency` | Currency code, ISO 4217 | — | Yes | Any | Currency of the position's amounts |
| *Illustrative:* `notional` | Decimal | Currency in field `currency` | Yes | ≥ 0 | Absolute amount, never negative |

### `DS-02` Parameters

Read from the sheet `Params`.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| `as_of_date` | Date, `YYYY-MM-DD` | — | Yes | Any | Name AsOfDate. The reporting date |
| `reporting_currency` | Currency code, ISO 4217 | — | Yes | Any | Name ReportingCcy. Every result is in it; DS-05 lists it with rate 1 |
| `days_per_year` | Decimal | Days | Yes | ≥ 0 | Name DaysPerYear. Greater than 0. Reference value 365 |
| `coupon_threshold` | Decimal | Decimal fraction | Yes | ≥ 0 | Name CouponThreshold. A coupon below it uses the low-coupon time bands |
| `scaling_ir` | Decimal | Factor | Yes | ≥ 0 | Name ScalingIR. Greater than 0 |
| `scaling_eq` | Decimal | Factor | Yes | ≥ 0 | Name ScalingEQ. Greater than 0 |
| `scaling_co` | Decimal | Factor | Yes | ≥ 0 | Name ScalingCO. Greater than 0 |
| `scaling_fx` | Decimal | Factor | Yes | ≥ 0 | Name ScalingFX. Greater than 0 |
| `rwa_multiplier` | Decimal | Factor | Yes | ≥ 0 | Name RWAMultiplier. Greater than 0 |
| `vertical_disallowance` | Decimal | Decimal fraction | Yes | 0 to 1 | Name VerticalDisallowance |
| `horizontal_zone_1` | Decimal | Decimal fraction | Yes | 0 to 1 | Name HorizontalZone1 |
| `horizontal_zone_2` | Decimal | Decimal fraction | Yes | 0 to 1 | Name HorizontalZone2 |
| `horizontal_zone_3` | Decimal | Decimal fraction | Yes | 0 to 1 | Name HorizontalZone3 |
| `horizontal_adjacent` | Decimal | Decimal fraction | Yes | 0 to 1 | Name HorizontalAdjacent. Between zones 1 and 2, and 2 and 3 |
| `horizontal_zone_1_3` | Decimal | Decimal fraction | Yes | 0 to 1 | Name HorizontalZone13. Between zones 1 and 3 |
| `eq_specific_rate` | Decimal | Decimal fraction | Yes | 0 to 1 | Name EQSpecificRate |
| `eq_index_rate` | Decimal | Decimal fraction | Yes | 0 to 1 | Name EQIndexRate |
| `eq_general_rate` | Decimal | Decimal fraction | Yes | 0 to 1 | Name EQGeneralRate |
| `fx_charge_rate` | Decimal | Decimal fraction | Yes | 0 to 1 | Name FXChargeRate |
| `gold_code` | Text | — | Yes | Any | Name GoldCode. The code gold is entered under in DS-05 and DS-09. Reference value XAU |
| `com_directional_rate` | Decimal | Decimal fraction | Yes | 0 to 1 | Name COMDirectionalRate |
| `com_gross_rate` | Decimal | Decimal fraction | Yes | 0 to 1 | Name COMGrossRate |

### `DS-03` Specific-risk charges

Read from the sheet `Params`.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| `key` | Text | — | Yes | Any | Column SR_Key, name SR_Key. Category, rating grade and maturity range joined by a vertical bar, as the workbook builds it for each position |
| `category` | Text | — | Yes | `GOV`, `QUAL`, `OTHER` | Column Category. Repeats the first part of the key; not read |
| `rating_grade` | Text | — | Yes | `AAA to AA-`, `A+ to BBB-`, `BB+ to BB-`, `B+ to B-`, `Below B-`, `Unrated` | Column Rating grade. Repeats the second part of the key; not read |
| `residual_maturity` | Text | — | Yes | `Any`, `6 months or less`, `over 6 to 24 months`, `over 24 months` | Column Residual maturity. Describes the third part of the key; not read |
| `charge` | Decimal | Decimal fraction | Yes | 0 to 1 | Column Charge, name SR_Charge |
| `reference` | Text | — | No | Any | Column Reference. Not read |

### `DS-04` Maturity ladder

Read from the sheet `Params`.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| `band` | Integer | — | Yes | 1 to 15 | Column Band, name Ladder_Band. The row's position gives the band; the number is not read |
| `zone` | Integer | — | Yes | 1 to 3 | Column Zone, name Ladder_Zone |
| `upper_bound_high_coupon` | Decimal | Years | No | ≥ 0 | Name Ladder_UpperHigh. Inclusive. Empty: the band is open-ended, or does not exist for this coupon |
| `upper_bound_low_coupon` | Decimal | Years | No | ≥ 0 | Name Ladder_UpperLow. Inclusive. Empty: the band is open-ended |
| `label_high_coupon` | Text | — | No | Any | Name Ladder_LabelHigh. Shown on IR_Ladder |
| `label_low_coupon` | Text | — | No | Any | Name Ladder_LabelLow. Shown on IR_Ladder |
| `risk_weight` | Decimal | Decimal fraction | Yes | ≥ 0 | Column Risk weight, name Ladder_RW. 0 or more |
| `assumed_yield_change` | Decimal | Percentage points | No | Any | Column Assumed yield change (pp). For the duration method, which is not supported; not read |

### `DS-05` FX rates

Read from the sheet `Params`.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| `currency` | Currency code, ISO 4217 | — | Yes | Any | Name FX_Ccy. Listed once. The reporting currency is listed, and gold under the gold code |
| `description` | Text | — | No | Any | Not read |
| `rate` | Decimal | Units of reporting currency per unit | Yes | ≥ 0 | Name FX_Rate. Greater than 0. The spot rate of the reporting date; 1 for the reporting currency |

### `DS-06` Interest rate positions

Read from the sheet `IRPositions`.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| `position_id` | Text | — | Yes | Any | Column PositionID. Unique on the sheet, without regard to case |
| `portfolio` | Text | — | No | Any | Column Portfolio. Not read |
| `instrument` | Text | — | Yes | `Bond`, `FRN`, `Swap leg`, `Future leg`, `FRA leg`, `Forward leg`, `Other` | Column Instrument. Not read yet: issue 44 |
| `currency` | Currency code, ISO 4217 | — | Yes | Any | Column Currency. A code of DS-05, the gold code excepted |
| `direction` | Text | — | Yes | `Long`, `Short` | Column Direction |
| `position_value` | Decimal | Currency in field `currency` | Yes | ≥ 0 | Column Position value (ccy). The market value of the principal or of the notional underlying; Direction gives the sign |
| `rate_type` | Text | — | Yes | `Fixed`, `Floating` | Column Rate type |
| `coupon` | Decimal | Decimal fraction | No | Any | Column Coupon (decimal). Required for a fixed-rate position |
| `final_maturity_date` | Date, `YYYY-MM-DD` | — | Yes | Any | Column Final maturity date. Not before the reporting date |
| `next_repricing_date` | Date, `YYYY-MM-DD` | — | No | Any | Column Next repricing date (floating). Required for a floating-rate position: from the reporting date to the final maturity |
| `sr_category` | Text | — | Yes | `GOV`, `QUAL`, `OTHER`, `NONE` | Column SR category. NONE: no specific risk |
| `rating_grade` | Text | — | No | `AAA to AA-`, `A+ to BBB-`, `BB+ to BB-`, `B+ to B-`, `Below B-`, `Unrated` | Column Rating grade. Required unless the category is NONE. With the category, a key of DS-03 |
| `issue_id` | Text | — | No | Any | Column Issue ID (specific-risk netting). Rows with the same value are one issue; empty: an issue of its own |
| `comment` | Text | — | No | Any | Column Comment. Not read |

### `DS-07` Equity positions

Read from the sheet `EQPositions`.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| `position_id` | Text | — | Yes | Any | Column PositionID. Unique on the sheet, without regard to case |
| `portfolio` | Text | — | No | Any | Column Portfolio. Not read |
| `market` | Text | — | Yes | Any | Column Market. The national market |
| `type` | Text | — | Yes | `Single name`, `Index` | Column Type |
| `identifier` | Text | — | Yes | Any | Column Stock or index identifier. With the market and the type, it identifies what is netted |
| `currency` | Currency code, ISO 4217 | — | Yes | Any | Column Currency. A code of DS-05, the gold code excepted |
| `direction` | Text | — | Yes | `Long`, `Short` | Column Direction |
| `market_value` | Decimal | Currency in field `currency` | Yes | ≥ 0 | Column Market value (ccy). Direction gives the sign |
| `comment` | Text | — | No | Any | Column Comment. Not read |

### `DS-08` Commodity positions

Read from the sheet `COPositions`.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| `position_id` | Text | — | Yes | Any | Column PositionID. Unique on the sheet, without regard to case |
| `portfolio` | Text | — | No | Any | Column Portfolio. Not read |
| `commodity` | Text | — | Yes | Any | Column Commodity. Rows with the same name are netted. Never gold: issue 44 |
| `direction` | Text | — | Yes | `Long`, `Short` | Column Direction |
| `quantity` | Decimal | Units of the commodity | Yes | ≥ 0 | Column Quantity. Direction gives the sign |
| `unit` | Text | — | No | Any | Column Unit. The standard unit of the quantity; not read |
| `spot_price` | Decimal | Currency in field `currency` | Yes | Any | Column Spot price (ccy per unit). The current spot price of one unit, used as entered |
| `currency` | Currency code, ISO 4217 | — | Yes | Any | Column Currency. A code of DS-05, the gold code excepted |
| `maturity_date` | Date, `YYYY-MM-DD` | — | No | Any | Column Maturity date. For the maturity ladder approach, which is not supported; not read |
| `comment` | Text | — | No | Any | Column Comment. Not read |

### `DS-09` FX positions

Read from the sheet `FXPositions`.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| `position_id` | Text | — | Yes | Any | Column PositionID. Unique on the sheet, without regard to case |
| `portfolio` | Text | — | No | Any | Column Portfolio. Not read |
| `currency` | Currency code, ISO 4217 | — | Yes | Any | Column Currency. A code of DS-05, the gold code included. A row in the reporting currency is excluded |
| `direction` | Text | — | Yes | `Long`, `Short` | Column Direction |
| `amount` | Decimal | Currency in field `currency` | Yes | ≥ 0 | Column Amount (ccy). The net open position in the currency; for gold, in the unit its rate is quoted for. Direction gives the sign |
| `comment` | Text | — | No | Any | Column Comment. Not read |

### `DS-10` Non-delta option charges

Read from the sheet `OptionCharges`.

| Field | Type | Unit or currency | Required | Allowed values or range | Notes |
| --- | --- | --- | --- | --- | --- |
| `risk_class` | Text | — | Yes | `IR`, `EQ`, `CO`, `FX` | Column Risk class. A label of the template, not typed |
| `description` | Text | — | No | Any | Column Description. A label of the template; not read |
| `method` | Text | — | No | Any | Column Method. The method the charge was calculated by; not read |
| `charge` | Decimal | Reporting currency | No | ≥ 0 | Column Charge (rep ccy), names OptionsIR, OptionsEQ, OptionsCO and OptionsFX. Empty: zero |
| `comment` | Text | — | No | Any | Column Comment. Not read |
<!-- END generated: data-contract-fields -->

<a id="reading-a-sheet"></a>

### How a sheet is read

The field tables give the type and the values that a
[test fixture](TEST_CASES.md#fixture-format) holds. On a sheet, the engine
reads a cell this way:

- **Fields and columns.** A field is a column of its table, in the order of
  its dataset's table above; a field of `DS-02` is the *Value* cell of its
  parameter, found by its workbook name. *Not read* marks a column that no
  calculation uses: it is there for the reader.
- **Rows.** A position table is read from the row below its header to the
  last row with a value in its first column. A row whose `position_id` is
  empty is skipped.
- **Text and codes.** A text is read without the blanks at its ends, and a
  code or one of a list of allowed values without regard to case: `long`,
  `Long` and ` LONG ` are the same direction. Keys that group rows, such as
  `issue_id`, `identifier`, `market` and `commodity`, are compared without
  regard to case too.
- **Currencies.** A currency is valid when `DS-05` lists it, whatever its
  form: the engine does not test that a code is three letters of ISO 4217.
  Use ISO codes, and `XAU` for gold.
- **Numbers and dates.** A number is a numeric cell; a text that reads like a
  number is not one. A date is a date cell, or a number that is the serial of
  a date from 1 January 1900 to 31 December 9999: a date typed as `20311231`
  is not a date.
- **Error values.** A cell that holds an error value, such as `#N/A`, reads
  as the text `#ERROR`, and as not a number.
- **Nothing is corrected.** A value that fails a test is reported as below,
  never rounded, replaced or defaulted. An empty option charge is the one
  default: it is zero.

A dataset's identifier field is its required `text` field whose name ends in
`_id`, such as the illustrative `position_id`. A dataset whose records a
[JSON fixture](TEST_CASES.md#fixture-format) lists has exactly one, and its
value is the record's `id`.

State the unit of every amount, rate and period explicitly: a percentage, a
decimal fraction and basis points are different fields, not different
formats.

<a id="machine-readable-contract"></a>

## Machine-readable contract

[`data-contract.json`](data-contract.json), beside this guide, holds the
datasets and their fields. Edit it, never the generated tables, then run
`python tools/render_docs.py --write`. **Repository integrity** fails while a
table differs from it, and `python tools/check_data_contract.py`, which it
also runs, checks the file itself.

| Key | Holds |
| --- | --- |
| `schema_version`, `kind` | `1` and `data-contract` |
| `datasets` | One object per dataset, holding the columns of the *Datasets* table as `id`, `illustrative` (`true` or `false`), `description`, `provider`, `frequency`, `granularity`, `source` (the file or sheet) and `encoding` (a Python codec name for a file, `null` for a sheet), then `delimiter`, `decimal_separator` and `fields` |
| `delimiter`, `decimal_separator` | For a delimited text file, the one character between cells, and `.` or `,`; both `null` for a sheet or for a file of another kind |
| `fields` | One object per field: `name` (lower snake case, unique in the dataset), `type`, `unit`, `required` (`true` or `false`), `allowed` and `notes` |
| `type` | `text`, `integer`, `decimal`, `date` (ISO `YYYY-MM-DD`), `boolean` (`true` or `false`) or `currency` (three upper-case letters, ISO 4217) |
| `unit` | `null`, a text such as `Decimal fraction` or `Years`, or `{"currency_field": "<name>"}` for an amount whose currency is in a `currency` field of the same dataset |
| `allowed` | `null`; `{"values": [...]}`, texts matched exactly, for a `text`, `integer`, `date` or `currency` field; or `{"min": x, "max": y}`, either bound optional and both inclusive, for an `integer` or `decimal` field |

The file is strict JSON, with no duplicate key and no non-finite number, and
a text shown in a table cell holds no `|` and no line break.

<a id="reference-validator"></a>

### Reference validator

`python tools/check_data_contract.py --dataset DS-NN FILE...` validates input
files of a delimited text dataset against the contract.
`tools/check_test_cases.py` runs the same validator on every variant file of
a [file-based case](TEST_CASES.md#file-based-cases), except in an `invalid`
case. It reads a file this way:

- The file decodes in the dataset's encoding. A byte order mark stays part of
  the first column's name unless the encoding, such as `utf-8-sig` or
  `UTF-16`, removes it.
- The first row is a header that names each column by its field name, in any
  order. Every required field has a column.
- Each further row has one cell per column; a blank line is a row with no
  cell. LF and CRLF line endings are both read, and a cell may be quoted with
  double quotes.
- An empty cell is a missing value. Each other cell is matched whole: a
  number is written in ASCII digits, with an optional sign and the dataset's
  decimal separator, and no thousands separator, exponent, surrounding blank
  or line break; a boolean is `true` or `false`, in lower case; and a quoted
  line break makes a code, date or boolean invalid too.

It checks the layout and the values of a file, never what they mean: test a
key that must be unique, a reconciliation against control totals or a rule of
the next section in the import itself. Each finding carries one of these
codes, which keep their meaning:

<!-- BEGIN generated: data-contract-codes -->
| Code | Finding |
| --- | --- |
| `DC-ENCODING` | The file does not decode in the dataset's encoding |
| `DC-SYNTAX` | The file is not readable as delimited text, such as a quote left open |
| `DC-HEADER` | The file has no header row |
| `DC-DUPLICATE-COLUMN` | The header names a column twice |
| `DC-UNKNOWN-COLUMN` | The header names a column that is not a field of the dataset |
| `DC-MISSING-COLUMN` | A required field has no column |
| `DC-CELL-COUNT` | A row, a blank line included, has more or fewer cells than the header |
| `DC-REQUIRED` | A required field is empty |
| `DC-TYPE` | A value is not of its field's type |
| `DC-VALUE` | A value is not one of its field's allowed values |
| `DC-RANGE` | A number is outside its field's range |
| `DC-CURRENCY` | An amount is given without a currency in its currency field |
<!-- END generated: data-contract-codes -->

The codes belong to the validator. The rules below are the workbook's, and the
workbook reports its own rule IDs.

## Validation and invalid input

The workbook reports a finding with the identifier of a rule of its `Checks`
sheet, `C01` to `C14`, and, for a row, with a status text that gives the
reason. The texts are listed in the *Failure behaviour* of each
[calculation contract](CALCULATION_CONTRACTS.md).

| Rule | Rejects | Effect |
| --- | --- | --- |
| `C01` | A row of `DS-06` that fails a test: its identifier, currency, direction, value, rate type, coupon, dates, category or rating grade | The row is excluded and the run goes on. `IRCalc` gives the reason, `Checks` lists the row, and the status line reads `INCOMPLETE` |
| `C02` | A row of `DS-07` that fails a test: its identifier, market, type, stock or index, currency, direction or value | As `C01`, on `EQCalc` |
| `C03` | A row of `DS-08` that fails a test: its identifier, commodity, direction, quantity, spot price or currency | As `C01`, on `COCalc` |
| `C04` | A row of `DS-09` that fails a test: its identifier, currency, direction or amount | As `C01`, on `FXCalc` |
| `C05` | `DS-05` without the reporting currency at rate 1 | The run stops: no result is calculated, `Checks` lists the value, and the status line reads `STOPPED` |
| `C06` | A currency listed twice in `DS-05` | As `C05` |
| `C07` | A rate of `DS-05` that is not a number greater than 0 | As `C05` |
| `C14` | A value of `DS-02`, `DS-03`, `DS-04` or `DS-10` that the engine cannot use: a date that is not a date, a number that is not a number or is outside its range, a zone that is not 1, 2 or 3, an empty reporting currency or gold code | As `C05` |

Four rules report without rejecting:

| Rule | Reports | Effect |
| --- | --- | --- |
| `C08` | A value of `DS-02`, `DS-03` or `DS-04` that differs from its reference value | A warning: the run is a what-if, calculated with the value entered |
| `C09` | An issue of `DS-06` whose rows differ in category, rating grade or currency | A warning: the issue takes the rate and the currency of its first row |
| `C10`, `C11` | More than 30 markets in `DS-07`, or more than 30 commodities in `DS-08` | A warning: `Buckets` lists the first 30, and the totals include all |
| `C12`, `C13` | A row of `DS-09` in the reporting currency; an option charge entered in `DS-10` | For information: the row is excluded, and the charge is added |

- A row is tested in the order of its contract's list, and the first test
  that fails gives the reason.
- Invalid input is reported, never silently corrected, rounded or defaulted.
- A table of `Params` whose workbook name is missing, or has another length,
  is not a finding: the run fails, and its status line reads `FAILED`
  ([status and staleness](WORKBOOK_CONTRACT.md#status-and-staleness)).
- After a run that stopped or failed, the workbook shows no result of an
  earlier run as if it were current.

## Unsupported input

The workbook cannot tell an unsupported instrument from a supported one that
it is entered as. A securitisation position or a credit derivative entered
in `DS-06` as a bond is calculated as a bond, and an index that is not well
diversified entered in `DS-07` as `Index` is charged as one. This is a
stated limit: you decide what is entered, and the
[support matrix](README.md#support-matrix) says what must not be.

What a field can exclude, the contract excludes: `instrument` has a list of
allowed values, and a commodity is never gold. The engine does not enforce
either yet ([gaps](#gaps)).

## Gaps

Where the engine does not yet do what this contract says:

| Field or rule | The contract says | The engine does | Issue |
| --- | --- | --- | --- |
| `DS-06.instrument` | Required, one of seven values | It does not read the column: any text, or none, is accepted | [44](https://github.com/danielep71/K-FRTB-SSA/issues/44) |
| `DS-08.commodity` | Never gold | A commodity named as the gold code is calculated as a commodity | [44](https://github.com/danielep71/K-FRTB-SSA/issues/44) |
| [Volume](#volume) | At most 500 rows on a position sheet | It reads and calculates every row, with no limit | [45](https://github.com/danielep71/K-FRTB-SSA/issues/45) |

The formula workbook that the engine is compared with has a gap of its own:
it reads a code of `DS-05`, a direction and a rate type as typed, where the
engine reads them without the blanks at their ends
([issue 41](https://github.com/danielep71/K-FRTB-SSA/issues/41)).

## Point-in-time data

Every dataset is that of the reporting date, `as_of_date` of `DS-02`: the
positions held at its close, and the spot rates and spot prices of that
date. The workbook holds one reporting date at a time and no history. It
cannot check that a rate or a position belongs to the date: you enter them
for it.

## Volume

| Dataset | Largest supported | Beyond it |
| --- | --- | --- |
| `DS-06` to `DS-09` | 500 rows each | The sheets are formatted, validated and checked for 500 rows. The engine does not stop yet ([gaps](#gaps)) |
| `DS-05` | 20 currencies, gold included | The table has 20 rows |
| `DS-03`, `DS-04` | 18 and 15 rows | Fixed by `SRC-02` |
| Markets of `DS-07`, commodities of `DS-08` | Any number; 30 of each are listed on `Buckets` | The totals include all, and `Checks` warns (`C10`, `C11`) |

No run on full sheets has been timed, in Excel or elsewhere.

## Lineage

No file is imported, so the workbook records no file name and no hash of a
file. What a run records, where it can be read later:

- the time of the run, on the status line of `Results`, which every output
  sheet repeats;
- a fingerprint of every input the run read, in `Results_Fingerprint`, which
  a later check compares with the sheets
  ([status and staleness](WORKBOOK_CONTRACT.md#status-and-staleness));
- the rows used and the rows rejected per risk class, on `Results`, and each
  rejected row with its reason, on `Checks`;
- the reporting date and the reporting currency, in the headline of
  `Results`;
- the commit, the time and the configuration of the build, on the `README`
  sheet.

It does not record who entered the inputs or where they came from.

## Synthetic data

The repository holds synthetic data only: no real positions, counterparties,
clients or amounts. Test fixtures say they are synthetic and how they were
made. The sample data of the template are synthetic too: I wrote them in
`tools/prototype/ssa_reference.py` to exercise each rule, and they describe
no portfolio.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| 2026-11-16 | Every dataset is typed on a sheet; the workbook imports no file | The prototype was built so, and a simplified approach serves small books that are entered by hand | [19](https://github.com/danielep71/K-FRTB-SSA/issues/19) |
| 2026-11-16 | `Params` holds four datasets, `DS-02` to `DS-05`, and a sheet may hold several | The parameter list, the specific-risk charges, the maturity ladder and the FX rates are four tables with their own fields. A file stays one dataset | [19](https://github.com/danielep71/K-FRTB-SSA/issues/19) |
| 2026-11-16 | The contract accepts what the engine accepts: a code without regard to case and without the blanks at its ends, and any currency that `DS-05` lists | The inputs are typed by hand, and rejecting ` eur` would add findings without removing a risk. The types of the field tables are what a fixture holds | [19](https://github.com/danielep71/K-FRTB-SSA/issues/19) |
| 2026-11-16 | An unsupported instrument entered as a supported one is a stated limit; `instrument` has allowed values, and gold is never a commodity | No field can make the workbook recognise what is entered under another name. What a field can exclude, the engine is to reject | [19](https://github.com/danielep71/K-FRTB-SSA/issues/19), [44](https://github.com/danielep71/K-FRTB-SSA/issues/44) |
| 2026-11-16 | A position sheet holds at most 500 rows, and a run on more is to stop | The sheets, the formula workbook and the comparison cover 500 rows; results beyond them would have no format and no check | [19](https://github.com/danielep71/K-FRTB-SSA/issues/19), [45](https://github.com/danielep71/K-FRTB-SSA/issues/45) |
| 2026-11-16 | The project's datasets are numbered from `DS-02` | `DS-01` is the illustrative dataset, which the illustrative test cases still use, and an ID is never reused | [19](https://github.com/danielep71/K-FRTB-SSA/issues/19) |
