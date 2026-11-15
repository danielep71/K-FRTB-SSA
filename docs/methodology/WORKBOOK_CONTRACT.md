# 📊 Workbook contract

This contract defines what the workbook looks like:

- its sheets and their code names;
- the defined names the code addresses;
- which areas are inputs and which the code writes;
- how each operation reads and writes them;
- what the status line says, and what a failed run leaves on the sheets;
- what the build creates, and how each sheet is protected.

The [data contract](DATA_CONTRACT.md) defines the inputs and the
[calculation contracts](CALCULATION_CONTRACTS.md) the results; this contract
says where they sit in the workbook. The
[repository structure](../REPOSITORY_STRUCTURE.md) says which components may
touch a sheet, and the [house style](../VBA_HOUSE_STYLE.md) how they are
named.

## Principles

These principles hold for the workbook. The [decisions](#decisions) record
where the prototype set them:

- **Built from source, never committed.** The build creates the workbook from
  the exported source of one commit, as the
  [installation guide](../../INSTALLATION.md) describes, and saves it outside
  the checkout or under `test-results/`. Only a macro-free
  `src/workbook/Template.xlsx` may be tracked.
- **Datasets where the data contract says.** The *File or sheet* column of the
  data contract's *Datasets* table says whether a dataset is read from a file
  into memory or typed on a sheet. A dataset read from a file is never copied
  onto a sheet.
- **No regulatory calculation in sheet formulas.** VBA writes every
  calculated value as a value. The template's formulas only repeat or
  compare written values: the *Changed* flags of `Params`, the headline of
  `Results`, the status line that the other output sheets repeat and the
  comparisons of `TestCatalogue`. Input areas use data validation, and
  conditional formats read written values.
- **The adapter owns the sheets.** Only `M_Main`, the Excel adapter in
  `src/workbook/`, reads or writes a worksheet. It finds every cell and
  table through a workbook name ([defined names](#defined-names)), never
  through a hard-coded address, and uses a tab name only to show `Results`
  or `Checks` after a run. The core modules and the worksheet functions of
  `FRTBSSA_Formulas` never touch a workbook object, and the event procedures
  of `ThisWorkbook` only call `M_Main`.
- **Calculate first, then replace.** A run reads the inputs and calculates
  in memory, then clears every output area and writes the new results, the
  status line last.
- **No partial results.** A run that raises an error, before or while it
  writes, clears every output area and writes a status line that begins
  with `FAILED`. No result of that run or of an earlier one stays
  ([status and staleness](#status-and-staleness)).
- **Outdated results are marked.** A run stores a fingerprint of the inputs
  it read. A check compares it with the inputs on the sheets and marks the
  results `OUTDATED` when an input changed after the run
  ([status and staleness](#status-and-staleness)).

## Sheets

Tab order is the order of the table. Every sheet is visible.

| # | Tab name | Code name | Role | Code writes | Content |
| ---: | --- | --- | --- | --- | --- |
| 1 | `Home` | `shHome` | Reference | No | What the workbook calculates, how to use it and what each sheet holds |
| 2 | `Params` | `shParams` | Input | No | Reporting date and currency, the MAR40 constants, the specific-risk table, the time bands and the FX rates |
| 3 | `IRPositions` | `shIRPositions` | Input | No | Debt positions and interest rate derivative legs, one row per position |
| 4 | `EQPositions` | `shEQPositions` | Input | No | Equity and index positions, one row per position |
| 5 | `COPositions` | `shCOPositions` | Input | No | Commodity positions, one row per position |
| 6 | `FXPositions` | `shFXPositions` | Input | No | Net currency and gold positions |
| 7 | `OptionCharges` | `shOptionCharges` | Input | No | Non-delta option charges per risk class, calculated outside the workbook |
| 8 | `Results` | `shResults` | Output | Yes | Requirement per risk class, scaling, total and RWA, with the status line of the run |
| 9 | `IR_Zones` | `shIRZones` | Output | Yes | Interest rate charges per currency, with every offset of the maturity method |
| 10 | `IR_Ladder` | `shIRLadder` | Output | Yes | Weighted positions per currency and time band |
| 11 | `Buckets` | `shBuckets` | Output | Yes | Equity by market, commodities by commodity and the FX net open position |
| 12 | `IRCalc` | `shIRCalc` | Output | Yes | Every intermediate value per interest rate row |
| 13 | `EQCalc` | `shEQCalc` | Output | Yes | Every intermediate value per equity row |
| 14 | `COCalc` | `shCOCalc` | Output | Yes | Every intermediate value per commodity row |
| 15 | `FXCalc` | `shFXCalc` | Output | Yes | Every intermediate value per FX row |
| 16 | `Checks` | `shChecks` | Diagnostics | Yes | Errors, warnings and information lines, and the rejected rows with their reasons |
| 17 | `TestCatalogue` | `shTestCatalogue` | Diagnostics | No | The test catalogue: the published example and the re-calculation, with each result |
| 18 | `README` | `shREADME` | Reference | Build only | Methodology, conventions, scope and limitations, and the build stamp |

- **Code name:** `sh` plus the tab name without its underscore. The
  workbook's own document module is `ThisWorkbook`.
- **Role:**
  - `Reference`: what the workbook is, its limits and its build;
  - `Input`: typed by users;
  - `Parameter`: values a calibration writes and later runs read;
  - `Output`: results;
  - `Diagnostics`: findings, logs and test results;
  - `Test`: present only in a development build.

  No sheet has the `Parameter` or the `Test` role yet.
- **Code writes:**
  - `No`: code never writes the sheet; users type in its input cells, where
    it has any;
  - `Yes`: code replaces its content, and users never edit it;
  - `Append only`: code adds rows, and never changes or deletes them;
  - `Build only`: only the build writes it.

  `TestCatalogue` holds formulas that compare output cells with expected
  values; neither code nor users write it.

## Layout

- Columns A and B are margins. On every sheet `B1` holds the workbook's name,
  `B2` the sheet's title and `C3` a note on what the sheet holds.
- Every table starts in column C. Its first header cell carries a name
  ending in `_Header`, and its data start on the row below. `Buckets` holds
  three tables one below the other, each with a fixed number of rows and a
  total row.
- A position sheet is formatted for 500 rows. `M_Main` reads a table from
  the row below its header to the last row with a value in its first
  column.
- Input cells have a yellow fill, and the MAR40 constants on `Params` a grey
  one. No cell is unlocked and no sheet is protected yet
  ([protection](#protection)).
- `Results` holds the status line, `Results_Status`, and every other output
  sheet repeats it in `E1`, in the band at its top, with the formula
  `=Results_Status`: `IR_Zones`, `IR_Ladder`, `Buckets`, the four calculation
  sheets and `Checks` ([status and staleness](#status-and-staleness)).

## Defined names

The names that refer to one cell:

| Name | Sheet | Refers to | Written by |
| --- | --- | --- | --- |
| `build_commit` | `shREADME` | The commit the workbook was built from | The build |
| `build_time` | `shREADME` | The commit's committer time | The build |
| `build_configuration` | `shREADME` | `production` or `development` | The build |
| `AsOfDate` | `shParams` | The reporting date | Users |
| `ReportingCcy` | `shParams` | The reporting currency | Users |
| `DaysPerYear` | `shParams` | Calendar days per year for residual maturities | Users |
| `CouponThreshold` | `shParams` | Coupon below which the low-coupon time bands apply | Users |
| `ScalingIR` | `shParams` | Scaling factor, interest rate risk | Users |
| `ScalingEQ` | `shParams` | Scaling factor, equity risk | Users |
| `ScalingCO` | `shParams` | Scaling factor, commodity risk | Users |
| `ScalingFX` | `shParams` | Scaling factor, foreign exchange risk | Users |
| `RWAMultiplier` | `shParams` | RWA = requirement x multiplier | Users |
| `VerticalDisallowance` | `shParams` | Vertical disallowance on matched weighted positions per time band | Users |
| `HorizontalZone1` | `shParams` | Horizontal disallowance within zone 1 | Users |
| `HorizontalZone2` | `shParams` | Horizontal disallowance within zone 2 | Users |
| `HorizontalZone3` | `shParams` | Horizontal disallowance within zone 3 | Users |
| `HorizontalAdjacent` | `shParams` | Horizontal disallowance between adjacent zones (1-2, 2-3) | Users |
| `HorizontalZone13` | `shParams` | Horizontal disallowance between zones 1 and 3 | Users |
| `EQSpecificRate` | `shParams` | Equity specific risk, single names (on net position per stock) | Users |
| `EQIndexRate` | `shParams` | Equity further charge on net position per index contract | Users |
| `EQGeneralRate` | `shParams` | Equity general market risk (on net position per market) | Users |
| `FXChargeRate` | `shParams` | FX and gold, on the overall net open position | Users |
| `GoldCode` | `shParams` | The currency code treated as gold | Users |
| `COMDirectionalRate` | `shParams` | Commodity directional charge (net position per commodity) | Users |
| `COMGrossRate` | `shParams` | Commodity charge on gross positions at spot | Users |
| `IRPositions_Header` | `shIRPositions` | The first header cell of the interest rate positions | Users, in the rows below it |
| `EQPositions_Header` | `shEQPositions` | The first header cell of the equity positions | Users, in the rows below it |
| `COPositions_Header` | `shCOPositions` | The first header cell of the commodity positions | Users, in the rows below it |
| `FXPositions_Header` | `shFXPositions` | The first header cell of the FX and gold positions | Users, in the rows below it |
| `OptionsIR` | `shOptionCharges` | The non-delta option charge of interest rate risk | Users |
| `OptionsEQ` | `shOptionCharges` | The non-delta option charge of equity risk | Users |
| `OptionsCO` | `shOptionCharges` | The non-delta option charge of commodity risk | Users |
| `OptionsFX` | `shOptionCharges` | The non-delta option charge of FX and gold | Users |
| `IRCalc_Header` | `shIRCalc` | The first header cell of the rows calculated per interest rate position | `RunSSA`, below it |
| `EQCalc_Header` | `shEQCalc` | The first header cell of the rows calculated per equity position | `RunSSA`, below it |
| `COCalc_Header` | `shCOCalc` | The first header cell of the rows calculated per commodity position | `RunSSA`, below it |
| `FXCalc_Header` | `shFXCalc` | The first header cell of the rows calculated per FX position | `RunSSA`, below it |
| `IR_Ladder_Header` | `shIRLadder` | The first header cell of the maturity ladder | `RunSSA`, below it |
| `IR_Zones_Header` | `shIRZones` | The first header cell of the charges per currency | `RunSSA`, below it |
| `IR_GMR_Total` | `shIRZones` | General market risk, all currencies, on the total row | `RunSSA` |
| `IR_SR_Total` | `shIRZones` | Specific risk, all currencies, on the total row | `RunSSA` |
| `Buckets_EQ_Header` | `shBuckets` | The first header cell of the equity markets | `RunSSA`, below it |
| `EQ_General` | `shBuckets` | Equity general market risk, on the total row | `RunSSA` |
| `EQ_Specific` | `shBuckets` | Equity specific risk, on the total row | `RunSSA` |
| `EQ_Index` | `shBuckets` | The further charge on index contracts, on the total row | `RunSSA` |
| `Buckets_CO_Header` | `shBuckets` | The first header cell of the commodities | `RunSSA`, below it |
| `CO_Directional` | `shBuckets` | The commodity directional charge, on the total row | `RunSSA` |
| `CO_Gross` | `shBuckets` | The commodity charge on gross positions, on the total row | `RunSSA` |
| `Buckets_FX_Header` | `shBuckets` | The first header cell of the currencies | `RunSSA`, below it |
| `FX_ONOP` | `shBuckets` | The overall net open position | `RunSSA` |
| `FX_Charge` | `shBuckets` | The FX and gold requirement | `RunSSA` |
| `SSA_Requirement` | `shResults` | The requirement after scaling, on the total row | `RunSSA` |
| `Results_Header` | `shResults` | The first header cell of the requirement per risk class | `RunSSA`, in the rows below it |
| `Results_Status` | `shResults` | The status line of the last run | `RunSSA`, `ClearOutputs` and `CheckStatus` |
| `Results_Fingerprint` | `shResults` | The fingerprint of the inputs of the last run | `RunSSA`; cleared by `ClearOutputs` and by a run that fails |
| `SSA_RWA` | `shResults` | The risk-weighted assets, on the total row | `RunSSA` |
| `Checks_Header` | `shChecks` | The first header cell of the rules | `RunSSA`, in the count and result columns |
| `Checks_Errors` | `shChecks` | The number of errors of the last run | `RunSSA` |
| `Checks_Warnings` | `shChecks` | The number of warnings of the last run | `RunSSA` |
| `Checks_Findings_Header` | `shChecks` | The first header cell of the findings | `RunSSA`, below it |

`tools/devbuild/workbook.json` holds the cell of each of these names, and the
build stops when the template defines one elsewhere.

- Every name is workbook-scoped and visible, and refers to cells of this
  workbook by an absolute reference. External and relative references are
  not allowed.
- **Parameters:** the name of a parameter is its code in column C of
  `Params`, on its value cell. A non-delta option charge is `Options` plus
  the code of its risk class.
- **Table anchors:** `<Sheet>_Header`, or `<Sheet>_<Table>_Header` on a sheet
  with several tables, on a table's first header cell.
- **Totals:** a name of its own on each total that code or a formula reads.
- **State:** `Results_Status`, the status line of the last run,
  `Results_Fingerprint`, the fingerprint of its inputs, and the two counts
  `Checks_Errors` and `Checks_Warnings`.
- **Build stamp:** `build_commit`, `build_time` and `build_configuration`, on
  the `README` sheet.
- Before it reads or writes through a name, `M_Main` checks that the name
  exists and refers to a range; a missing name stops the operation with a
  layout error.

## Range names

Twelve names refer to one column of a table. The template defines them, and
the build copies them unchanged; `tools/devbuild/workbook.json` lists
one-cell names only, so it does not check them
([decisions](#decisions)).

| Name | Sheet | Refers to | Holds |
| --- | --- | --- | --- |
| `SR_Key` | `shParams` | `C31:C48` | The keys of the specific-risk table: category, rating grade and residual maturity |
| `SR_Charge` | `shParams` | `G31:G48` | The charge of each specific-risk key |
| `Ladder_Band` | `shParams` | `C53:C67` | The number of each time band |
| `Ladder_Zone` | `shParams` | `D53:D67` | The zone of each time band |
| `Ladder_UpperHigh` | `shParams` | `E53:E67` | The upper bound of each time band in years, coupon at or above `CouponThreshold` |
| `Ladder_UpperLow` | `shParams` | `F53:F67` | The upper bound of each time band in years, coupon below `CouponThreshold` |
| `Ladder_LabelHigh` | `shParams` | `G53:G67` | The label of each time band, coupon at or above `CouponThreshold` |
| `Ladder_LabelLow` | `shParams` | `H53:H67` | The label of each time band, coupon below `CouponThreshold` |
| `Ladder_RW` | `shParams` | `I53:I67` | The risk weight of each time band |
| `FX_Ccy` | `shParams` | `C72:C91` | The currency codes of the FX table |
| `FX_Rate` | `shParams` | `E72:E91` | The rate of each currency to the reporting currency |
| `Tests_Result` | `shTestCatalogue` | `M7:M25` | The result of each test, `PASS` or `FAIL` |

Users write the eleven names on `Params`; `Tests_Result` is a column of
formulas. `M_Main` reads every range name but two: no code reads
`Ladder_Band`, and the test suite reads `Tests_Result`.

## Areas

| Area | Sheet | Kind | Written by | Holds |
| --- | --- | --- | --- | --- |
| Parameters | `shParams` | Input | Users, in the *Value* column | One row per parameter, from `AsOfDate` to `COMGrossRate`: code, description, value, reference, MAR40 value and a *Changed* flag |
| Specific-risk table | `shParams` | Input | Users | 18 rows: key, category, rating grade, residual maturity and charge (`SR_Key`, `SR_Charge`) |
| Maturity ladder | `shParams` | Input | Users | 15 time bands: zone, upper bounds and labels by coupon, and risk weight (the `Ladder_` names) |
| FX table | `shParams` | Input | Users | Up to 20 currencies with their rate to the reporting currency (`FX_Ccy`, `FX_Rate`) |
| `IRPositions_Header` | `shIRPositions` | Input | Users | One row per debt position or derivative leg; the rows vary |
| `EQPositions_Header` | `shEQPositions` | Input | Users | One row per equity or index position; the rows vary |
| `COPositions_Header` | `shCOPositions` | Input | Users | One row per commodity position; the rows vary |
| `FXPositions_Header` | `shFXPositions` | Input | Users | One row per net currency or gold position; the rows vary |
| Option charges | `shOptionCharges` | Input | Users, in the *Charge* column | Four rows, one per risk class (the `Options` names) |
| `Results_Header` | `shResults` | Result | `RunSSA` | Four risk classes and a total row; below them a headline of four formulas |
| `Results_Fingerprint` | `shResults` | Parameter | `RunSSA` | One cell: the fingerprint of the inputs of the last run, which `CheckStatus` reads |
| `IR_Zones_Header` | `shIRZones` | Result | `RunSSA` | One row per currency of the FX table, and a total row |
| `IR_Ladder_Header` | `shIRLadder` | Result | `RunSSA` | One block of 15 time bands per currency of the FX table |
| `Buckets_EQ_Header` | `shBuckets` | Result | `RunSSA` | Up to 30 equity markets, and a total row |
| `Buckets_CO_Header` | `shBuckets` | Result | `RunSSA` | Up to 30 commodities, and a total row |
| `Buckets_FX_Header` | `shBuckets` | Result | `RunSSA` | One row per currency of the FX table, a total row, `FX_ONOP` and `FX_Charge` |
| `IRCalc_Header`, `EQCalc_Header`, `COCalc_Header`, `FXCalc_Header` | The four calculation sheets | Result | `RunSSA` | One row per input row, on the same row number, with its status and every intermediate value |
| `Checks_Header` | `shChecks` | Diagnostics | `RunSSA`, in the count and result columns | 14 rules, `C01` to `C14`, with severity, sheet, reference and message |
| `Checks_Findings_Header` | `shChecks` | Diagnostics | `RunSSA` | One row per rejected row or unusable parameter, with the reason |
| Test catalogue | `shTestCatalogue` | Diagnostics | Formulas | 19 tests: an output cell, its expected value, the difference, the tolerance and the result (`Tests_Result`) |

An area is named by its table anchor, which refers to its first header cell
only ([defined names](#defined-names)). `M_Main` derives the area from the
anchor: the rows below it, with the number of columns `CORE_Config` holds
for that table.

- **Input** areas hold what users type. The
  [data contract](DATA_CONTRACT.md) is to define each field's type, unit and
  valid values; until it does, the header row and the note in `C3` of each
  sheet describe them. A row that fails validation is rejected: it
  contributes nothing, its calculation row says why and `Checks` lists it.
  An unusable parameter stops the run. No value is corrected.
- **Result** areas hold the outputs of the run, in the reporting currency
  and in unscaled units. Equity markets or commodities beyond the 30 rows of
  `Buckets` are not listed; the totals include them and `Checks` warns.
- **Diagnostics** areas hold the rules, the findings and the test catalogue.
  The workbook has no operation log.
- The one **Parameter** area, a value that an operation writes and a later
  one reads, is `Results_Fingerprint`. `RunSSA` writes it, and `CheckStatus`
  compares it with the fingerprint of the inputs on the sheets. Do not type
  in it.

## Operations

| Entry point | Reads | Writes | Requires |
| --- | --- | --- | --- |
| `M_Main.RunSSA` | Every input area | Every result and diagnostics area but the test catalogue, `Results_Fingerprint` and `Results_Status` | Every name it uses |
| `M_Main.ClearOutputs` | Nothing | The same areas, cleared; `Results_Status` set to not calculated | Every name it uses |
| `M_Main.CheckStatus` | Every input area, `Results_Fingerprint` and `Results_Status` | `Results_Status`, and only when it must change | The names of the input areas and of the two cells |

- `RunSSA_Silent`, `ClearOutputs_Silent` and `CheckStatus_Silent` do the same
  without a message box or a change of the active sheet. They return the
  result line and re-raise an error.
- `ThisWorkbook` runs the check when the workbook is opened and before it is
  saved, through `M_Main.CheckStatus_OnEvent`. That entry point shows no
  message and raises no error, so that an unattended open or save is never
  stopped. When the check cannot be made, because a name is missing, it
  marks the results `OUTDATED`; `CheckStatus` reports the error. It leaves
  the workbook marked as saved when it was and the status line did not
  change.
- Each entry point is an adapter procedure that calls the engine,
  `CORE_Engine.Calculate`; it holds no regulatory formula.
- It captures and restores the Excel settings it changes: calculation
  mode, screen updating and events. It reports both the operation and the
  cleanup outcome
  ([house style](../VBA_HOUSE_STYLE.md#operation-and-cleanup-outcomes)).
- Clearing removes values, not formats, and puts back the labels of the
  total rows. It goes on after an area it cannot clear, and then stops with
  a layout error that names it.
- `FRTBSSA_Formulas` holds worksheet functions. They calculate from their
  arguments and read or write no sheet.

## Status and staleness

`Results_Status` holds one line, and every other output sheet repeats it in
`E1`:

| Line begins | Set by | Meaning |
| --- | --- | --- |
| `Not calculated` | The template and `ClearOutputs` | The output areas are empty |
| `OK` | `RunSSA` | Every input row was used; the line gives the number of positions and the engine version |
| `INCOMPLETE` | `RunSSA` | Rows were rejected and are excluded from the totals; `Checks` lists them |
| `STOPPED` | `RunSSA` | A parameter is unusable: no result was calculated, and only `Checks` was written |
| `FAILED` | `RunSSA` | The run raised an error: the output areas are empty, and the line gives the time and the error |
| `OUTDATED` | `CheckStatus`, before the line of the run | An input changed after the run: the results shown are not those of the inputs on the sheets |

The line of a run that ended, `OK`, `INCOMPLETE` or `STOPPED`, ends with the
time of the run, as `Run 2026-11-15 14:05:09.`: the local time of the
computer, without a time zone.

| Area kind | After a run that ends | After a run that raises an error |
| --- | --- | --- |
| Input | Unchanged: no operation writes it | Unchanged |
| Result | Cleared, then written, unless the run is `STOPPED` | Cleared |
| Diagnostics | Counts and findings rewritten | Cleared |
| Parameter | The fingerprint of the inputs of the run | Cleared |

**A run that fails.** Whether the error is raised while the inputs are read,
while the calculation runs or while the results are written, the run clears
every output area and writes the `FAILED` line before it reports the error.
The status line is the first cell cleared and the last one written, so a
run that is interrupted in between, with no error to handle, leaves
`Not calculated` and never the line of a run. An area that cannot be
cleared, because its name is missing, is named on the `FAILED` line. When
`Results_Status` itself cannot be written, the error message says so.

**An input changed after a run.** `RunSSA` stores in `Results_Fingerprint`
a fingerprint of everything it read: the parameters, the specific-risk
table, the ladder, the FX table, the option charges and the four position
tables. The check reads the same inputs again and compares:

- It runs on request, as `CheckStatus`; when the workbook is opened; and
  before it is saved. A change is therefore reported at the next of these
  operations, not at the moment it is typed. Run `CheckStatus`, or rerun,
  before you read the results of a workbook you have edited.
- When the fingerprints differ, it writes `OUTDATED:` before the line of the
  run. When an input is put back, they agree again, and it takes the mark
  away.
- It returns `CURRENT`, `OUTDATED`, `FAILED` or `NOT_CALCULATED`, and writes
  nothing when the line already says so.
- Any change counts: a text typed in another case, or a row added below a
  table, marks the results even where the calculation would not change.
- The fingerprint is two 31-bit hashes of the values, which tell a change
  from no change; it is not a cryptographic digest and proves nothing about
  who made the run.

The other output sheets show the line through a formula. With manual
calculation, they show the new line after the next recalculation.

Nothing of this is tested in Excel yet. In LibreOffice's VBA mode, which is
not a supported host, `TEST_Suite` covers a run that fails while it writes,
one that fails while it reads, and a parameter and a position changed after
a run.

## Build

The build writes, from one commit:

- the sheets of [Sheets](#sheets), in order, with their tab and code names;
- every defined name;
- the formats, the data validation of the input areas and the conditional
  formats, as the template holds them;
- the build stamp on the `README` sheet: the commit, its committer time
  (not the clock, so that a rebuild is byte-identical) and the configuration.
  A working tree with changes, or a build whose `--override` replaces
  modules, is stamped as not evidence-grade;
- the VBA components of that commit, as the
  [installation guide](../../INSTALLATION.md) lists them.

It writes no results; the input sheets hold the template's synthetic sample
data. A production build has no `Test` sheet, test
module or example; a development build adds them.

`tools/devbuild/build_workbook.py` builds the sheets, code names, defined
names, stamp and VBA project. It reads them from
`tools/devbuild/workbook.json`, whose sheets, roles and names must match the
tables above. It leaves the `Test` sheets and their names out of every
configuration but `development`. It builds the sheets in one of two ways:

- **Blank sheets**, without a `src/workbook/Template.xlsx`. Each sheet holds
  only the stamp, a title on the reference sheet and, to the left of each
  stamp cell, its label. Formats, data validation and print areas are not
  built on this path. This repository has a template, so its builds do not
  take it.
- **From the template**, `src/workbook/Template.xlsx`. The build
  starts from a copy of it, so the formats, data validation, print areas and
  protection are the template's. The template holds the sheets of the table
  other than `Test`, in order, and defines the one-cell names on the same
  cells; the build stops otherwise. The build sets the code names, writes the
  stamp values into the stamp cells, keeping their style, makes the workbook
  macro-enabled and adds the VBA project. A development build appends each
  `Test` sheet blank, with its names, so `Test` sheets come last in the
  table. Every other part of the template is copied unchanged, and two builds
  of the same commit and working tree stay byte-identical.

`tools/prototype/build_prototype.py` writes the template, with the sample
data, as a macro-free `.xlsx` workbook without personal document metadata;
it is not edited in Excel. **Repository integrity** fails while the tracked
template is not what the tool writes, and on macros, or on an author, last
editor, company or manager in it. The
[installation guide](../../INSTALLATION.md#building-a-workbook-from-source)
says how to write the template and how to run the build.

## Protection

The template carries the protection of each sheet and the lock flag of each
cell. The protection matrix, `docs/methodology/workbook-protection.json`,
records them. No sheet is protected yet, and the matrix declares no input or
output range, so each of its 18 entries reads like these two:

```json
{
  "schema_version": 1,
  "sheets": {
    "shParams": {"protection": {}},
    "shResults": {"protection": {}}
  }
}
```

[Issue 26](https://github.com/danielep71/K-FRTB-SSA/issues/26) tracks the
protection of the sheets and the unlocking of the input cells.

- **Sheets:** one entry for each sheet of the template, by code name. `Test`
  sheets are not in the template and have no entry.
- **`protection`:** the options of the sheet's protection, named as the
  attributes of the SpreadsheetML `sheetProtection` element. `sheet: true`
  turns protection on; every other option set to `true` blocks its action, or
  protects its part, while the sheet is protected. An option left out takes
  the standard's default: `false` for `sheet`, `objects`, `scenarios`,
  `selectLockedCells` and `selectUnlockedCells`, `true` for the others. `{}`
  is an unprotected sheet. Never put a password in the matrix.
- **`input`:** the cells and ranges, such as `E6` or `C5:P504`, that users
  type in. They must be unlocked.
- **`output`:** the cells and ranges that users must not change, such as
  results and labels. They must be locked. An input range never overlaps an
  output range; a cell in neither is not checked.

`tools/check_workbook_protection.py`, which **Repository integrity** runs,
compares the matrix with the template's sheets and styles. A cell takes the
style of its own cell, or else of its row when the row is formatted, or else
of its column, as Excel shows it; a style is locked unless it unlocks the
cell. A lock flag that differs from the matrix fails the check, and so does
a protection option, compared when the matrix or the template protects the
sheet; the options of a sheet that neither protects are not compared. Without the template the check reads the matrix alone, and
without either it passes as inactive; a template without a matrix fails. It
compares neither the password nor protection that VBA sets at run time with
`Worksheet.Protect`.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| 2026-11-14 | The sheets, their order and their names are those of the MAR40 prototype, as received | The VBA source, the template's formulas and its test catalogue refer to them | [16](https://github.com/danielep71/K-FRTB-SSA/issues/16) |
| 2026-11-14 | A parameter's name is its code, without a `cfg_` prefix | `M_Main` builds the name from the code in `CORE_Config`, and the template's formulas use it | [16](https://github.com/danielep71/K-FRTB-SSA/issues/16) |
| 2026-11-14 | Twelve names refer to a column of a table, not to one cell | `M_Main` reads the specific-risk table, the ladder and the FX table as columns of a fixed length. The build configuration checks one-cell names only, so the contract lists the range names apart | [16](https://github.com/danielep71/K-FRTB-SSA/issues/16) |
| 2026-11-14 | The build stamp sits on the `README` sheet, in `D38:D40` | It is the sheet that describes the workbook and its limits | [16](https://github.com/danielep71/K-FRTB-SSA/issues/16) |
| 2026-11-14 | The sheets keep formulas that repeat or compare written values | They show a changed constant, the headline and the test results without a run of their own; none calculates a requirement | [16](https://github.com/danielep71/K-FRTB-SSA/issues/16) |
| 2026-11-15 | One path builds the workbook: `tools/prototype/build_prototype.py` writes the template, and `tools/devbuild/build_workbook.py` builds the workbook from it. LibreOffice, where used, runs that workbook | Two build paths would let the workbook that is checked differ from the one that is built. The tool that inserted source text into a template and patched it on the way in is not kept | [17](https://github.com/danielep71/K-FRTB-SSA/issues/17) |
| 2026-11-15 | A date is a serial from 1 January 1900 to 31 December 9999, and no shorter limit applies to a final maturity | A number above that range, such as a date typed as `20311231`, is not a date in Excel or VBA. A long maturity is valid, and the last time band has no upper bound | [29](https://github.com/danielep71/K-FRTB-SSA/issues/29) |
| 2026-11-15 | A changed input is found by a fingerprint of the inputs, which a check compares on request, when the workbook is opened and before it is saved | A change event would mark the results at once, but its first write after a run empties Excel's undo history, it does not fire while events are off, and it cannot be exercised outside Excel | [25](https://github.com/danielep71/K-FRTB-SSA/issues/25) |
| 2026-11-15 | Every output sheet repeats the status line of `Results` with a formula, in `E1` | A reader of one sheet can tell whether it belongs to the last run, and one written cell holds the state | [25](https://github.com/danielep71/K-FRTB-SSA/issues/25) |
| 2026-11-14 | No sheet is protected and no cell is unlocked yet | The prototype is brought in as received; the change is tracked | [26](https://github.com/danielep71/K-FRTB-SSA/issues/26) |
