# 🧮 The prototype

This guide is for you if you want to use the workbook, or to understand what
it does before you read its contracts. It is a prototype: I have run it only
in LibreOffice's VBA mode, never in Excel, and I have not validated it for
regulatory reporting. [What has been checked](#what-has-been-checked) says
what exists.

## What it calculates

The market risk own funds requirement and the risk-weighted assets under the
simplified standardised approach of the Basel Framework's chapter MAR40
([`SRC-02`](methodology/SOURCES.md)):

- **Interest rate risk:** the specific risk of debt positions, netted per
  issue, and the general market risk by the maturity method, with one ladder
  per currency ([`CALC-02`](methodology/CALCULATION_CONTRACTS.md),
  [`CALC-03`](methodology/CALCULATION_CONTRACTS.md)).
- **Equity risk:** the specific risk, the index charge and the general market
  risk per national market (`CALC-04`).
- **Commodities risk:** the simplified approach, per commodity (`CALC-05`).
- **Foreign exchange risk,** gold included, by the shorthand method
  (`CALC-06`).
- **The requirement:** the four charges and the non-delta option charges you
  enter, each class multiplied by its scaling factor, and the risk-weighted
  assets as 12.5 times the total (`CALC-07`, `CALC-08`).

The [contracts](methodology/CALCULATION_CONTRACTS.md) state each rule with its
paragraph of MAR40.

## What you need

- Microsoft 365 Excel desktop on Windows, 64-bit, the host the
  [installation guide](../INSTALLATION.md#%EF%B8%8F-host-policy) names.
  Nothing here needs a lower macro-security level.
- A workbook you build from the source of one commit. The repository ships no
  workbook, and none is released: [How to build it](#how-to-build-it).
- Macros enabled in that workbook. Its code reads and writes the workbook
  itself and nothing else: [`SECURITY.md`](../SECURITY.md#%EF%B8%8F-runtime-risk-surface)
  says what each entry point does.

## The sheets

The workbook has 18 sheets, all visible, in this order
([workbook contract](methodology/WORKBOOK_CONTRACT.md#sheets)):

| Kind | Sheets | What you do |
| --- | --- | --- |
| Reference | `Home`, `README` | Read them. `Home` says what to do; `README` holds the conventions, the limits and the build stamp |
| Input | `Params`, `IRPositions`, `EQPositions`, `COPositions`, `FXPositions`, `OptionCharges` | Type in them. No code writes them |
| Output | `Results`, `IR_Zones`, `IR_Ladder`, `Buckets`, `IRCalc`, `EQCalc`, `COCalc`, `FXCalc` | Read them. A run replaces their content |
| Diagnostics | `Checks`, `TestCatalogue` | Read them: the findings of a run, and the comparison of the sample data with expected values |

## The inputs

Every sheet is protected, without a password, and only the input cells are
unlocked. They are the yellow cells, and on `Params` the grey cells of the MAR40
constants too: you type in them, and in nothing else. You can select any cell,
and change the width of a column and the height of a row. You cannot insert or
delete a row, sort or filter, or change a label, a header, a formula or a
result. The protection prevents accidents and is no security: to change a
locked cell, choose **Review → Unprotect Sheet**. The macros write the output
sheets without unprotecting them for you.

The [data contract](methodology/DATA_CONTRACT.md) defines every field, its
unit, its allowed values and how a cell is read.

- **`Params`:** the reporting date and currency, the regime, the MAR40
  constants, the specific-risk table, the time bands and the FX rates, one per
  currency and one for gold. The constants hold the reference values of MAR40;
  a value you change is a what-if, and `Checks` reports it. `Regime` is `Basel`
  by default. Set it to `CRR` and the requirement is multiplied by 0.9, the
  transitional multiplier of Article 495s(2) of Regulation (EU) No 575/2013, on
  a reporting date up to `TransitionalEndDate`, 2029-12-31; that multiplier is
  the first of the two rules that the `CRR` setting changes today. The second
  is the de minimis of Article 351: the FX requirement is zero while the overall
  net open position does not exceed 2% of `TotalOwnFunds`, which you enter and
  which the `CRR` requires. The [comparison](methodology/CRR_COMPARISON.md)
  lists the other differences.
- **`IRPositions`, `EQPositions`, `COPositions`, `FXPositions`:** one row per
  position, up to 500 rows on each sheet. The workbook converts nothing: a
  derivative is a row for each leg, the category and rating of an issuer are
  yours to enter, and the net position of a currency is the figure you
  derive outside the workbook.
- **`OptionCharges`:** the non-delta charge of each risk class, calculated
  outside the workbook, one value per class, added before scaling.

The template ships synthetic sample data. It reproduces the published example
of table 9 of MAR40.61 for foreign exchange risk, and is otherwise
illustrative: replace it with your own synthetic or real inputs before you
read a result.

## How to run it

Press **Alt+F8** and run one of the macros of `M_Main`:

| Macro | What it does |
| --- | --- |
| `RunSSA` | Reads every input, calculates, and writes every output sheet. It then shows a summary |
| `ClearOutputs` | Empties the output sheets and sets the status to not calculated |
| `CheckStatus` | Tells you whether the results shown are those of the inputs on the sheets |

`RunSSA_Silent`, `ClearOutputs_Silent` and `CheckStatus_Silent` do the same
without a message box. The workbook also runs the check when you open it and
before you save it.

The status line of `Results`, and `E1` of every other output sheet, is the
status of the last run:

| Begins | Meaning |
| --- | --- |
| `Not calculated` | The output sheets are empty |
| `OK` | Every input row was used |
| `INCOMPLETE` | Some rows were rejected, and are excluded from the totals; `Checks` lists each with its reason |
| `STOPPED` | A parameter is unusable: nothing was calculated, and `Checks` says which |
| `FAILED` | The run raised an error: the output sheets are empty, and the line gives the time and the error |
| `OUTDATED` | An input changed after the run: do not read the results until you run again |

A row is `OK`, `INVALID: <reason>` or `EXCLUDED: <reason>` on its calculation
sheet. A rejected row contributes nothing: read `Checks` before you rely on
a total.

`Results` gives, for each risk class, the specific risk, the general market
risk, the other charge, the options, the requirement before scaling, the
scaling factor and the requirement after scaling, then the total, a row for
the transitional multiplier (1 under the Basel regime) and the risk-weighted
assets. `IRCalc`, `EQCalc`, `COCalc` and `FXCalc` hold every
intermediate value of each row; `IR_Ladder`, `IR_Zones` and `Buckets` hold the
ladders, the offsets and the totals per market, commodity and currency.

## The worksheet functions

Four functions calculate from their arguments and read or write no sheet:

| Function | Gives |
| --- | --- |
| `FRTBSSA_TimeBand(ResidualMaturity, Coupon, Floating)` | The time band of a position, by table 4 of MAR40.26 |
| `FRTBSSA_LadderCharge(WeightedLong, WeightedShort)` | The general market risk charge of one ladder |
| `FRTBSSA_FXCharge(SumNetLong, SumNetShort, GoldPosition)` | The foreign exchange charge: 8% of the larger of the two sums, plus gold, as in table 9 of MAR40.61 |
| `FRTBSSA_CommodityCharge(NetPosition, GrossPosition)` | The commodity charge: 15% of the net position plus 3% of the gross position |

An input outside a function's domain gives an error value. The
[public API inventory](PUBLIC_API.txt) lists them as the supported API.

## How to build it

From a clone, at the commit you want, with Python 3.11 or later:

```sh
python tools/devbuild/build_workbook.py --configuration production
```

It writes `test-results/build/Workbook-production.xlsm`, which Git ignores:
the sheets of the template, the VBA modules of `src/` and a build stamp with
the commit. The `development` configuration adds the test modules. The
[installation guide](../INSTALLATION.md#%EF%B8%8F-building-a-workbook-from-source)
says how to open it, how to compile it and how to record an Excel run.

## What it does not do

- **Interest rate risk by the duration method,** commodities risk by the
  maturity ladder approach, and the delta-plus and scenario methods for
  options: the workbook uses the maturity method and the simplified
  approach, and you enter the non-delta charge of options.
- **Securitisation positions,** credit derivatives and the correlation
  trading portfolio. Nothing in an input identifies one: entered as a debt
  position, it is calculated as one. Do not enter it.
- **Any conversion for you:** derivatives into positions, the delta of an
  option, the category and rating of an issuer, and the net open position in
  each currency are inputs.
- **National discretions,** the equity arbitrage treatments, the offset of
  depository receipts, the exemption for insignificant foreign exchange
  business, and a jurisdiction's own implementation of MAR40.
- **Everything outside MAR40,** such as the sensitivities-based method of the
  standardised approach.
- **More than 500 rows** on a position sheet: the template is formatted for
  500, and a run on more stops. The status line reads `STOPPED`, no result
  is calculated, and `Checks` names the sheet and the rows found. The
  workbook rejects an instrument that is not in the list of `IRPositions`,
  but it cannot tell a supported instrument from one entered under its name.

The [support matrix](methodology/README.md#support-matrix) lists each feature
with its status, and each contract's *Limits and non-goals* name what it
leaves out.

## What has been checked

- **MAR40.** Every parameter of the engine equals the text of the Basel
  Framework chapter in force from 2023-01-01, compared on 2026-11-16.
- **Test cases.** 20 [test cases](methodology/TEST_CASES.md) run on the
  engine in LibreOffice and agree with their expected values. 16 values are
  published: two of the example of table 9 of MAR40.61, the overall net open
  position and the requirement; three of the example of MAR40.27, the
  vertical disallowance of one band; and eleven of the worked example of the
  maturity method in the 2005 update of the 1996 Amendment. The other 350 are
  `illustrative`: I derived them from the contracts, as I wrote the engine
  from them, so their agreement validates no rule.
- **Tests of the code.** Four VBA suites run under the
  [test harness](methodology/TEST_HARNESS.md), in LibreOffice only.
- **Excel.** The workbook has not been compiled or run in Excel, so no
  evidence of a run exists. The [Host behaviour](methodology/VALIDATION_PLAN.md#host-behaviour)
  table lists what the code relies on Excel to do, and which of it has been
  measured.

The [validation plan](methodology/VALIDATION_PLAN.md) gives each rule's status.
Do not use a result of this workbook for regulatory reporting, and do not read
it as a statement of compliance.
