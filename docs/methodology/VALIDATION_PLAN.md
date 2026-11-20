# ✅ Validation plan

This plan says how each [calculation contract](CALCULATION_CONTRACTS.md) is
shown to hold, and what evidence a result needs before anyone relies on it.

## Responsibilities

Implementing a calculation, checking it independently and accepting it are
separate steps, recorded separately even when one person does all three.
Code review and passing repository checks are necessary, but they are not
evidence that a figure is right: the static checks never run Excel or a
calculation.

## Reference values

An expected result is never taken from the workbook. Two kinds of value
validate a rule:

- **Published:** printed in a registered [source](SOURCES.md), such as a
  worked example of the regulation.
- **Independent:** computed by an external implementation by another author,
  registered in the source register with an `IMPL-` ID.

A value the project computes itself, by hand, in a separate spreadsheet that
shares no formulas with the workbook, or with a Python script under
`tests/derivations/`, is an `illustrative` cross-check: it catches defects,
but it validates no rule. Record how each value was derived and who derived
it. A reference script is a test tool, never a runtime dependency of the
workbook.

Each comparison has an absolute and a relative tolerance, with the reason for
them: the source's own rounding, or the precision of `Double` arithmetic.

The [numerical test cases](TEST_CASES.md) record each expected value with its
reference class (`published`, `independent` or `illustrative`), its source,
its derivation and its tolerance
([reference classes](TEST_CASES.md#reference-classes)). Only `published` and
`independent` values validate a rule: a rule reaches *Validated* in the
[traceability table](#traceability) only through one of them.

## Validation matrix

| Calculation | Reference values | Evidence | Status |
| --- | --- | --- | --- |
| *Illustrative:* `CALC-01` | `illustrative` cross-checks only: two `nominal` cases, a JSON fixture and a file-based case, whose total of two synthetic notional amounts a derivation script computes; empty input and invalid input rejected are still to cover; no `published` or `independent` value yet | Evidence record per release candidate | Not implemented |
| `CALC-02` | `illustrative` values that `tests/derivations/ssa.py` computes from the contract: a `nominal` case on the sample portfolio and a `boundary` case on the limits of table 1. No `published` or `independent` value. The rows are those of `CALC-03`, whose `invalid` case covers them | None: nothing was run in Excel ([issue 22](https://github.com/danielep71/K-FRTB-SSA/issues/22)). In LibreOffice, which is not a supported host, `tools/prototype/run_cases.py` finds every expected value in the workbook | Exercised |
| `CALC-03` | Published values in two `nominal` cases: the weighted positions, the vertical and within-zone disallowances, the net position charge and the charge of the worked example C.2 of `SRC-03` (`bcbs119-c2-maturity-method`), and the weighted positions and the vertical disallowance of the example of MAR40.27 of `SRC-02` (`mar40-27-vertical-disallowance`). `illustrative` values that `tests/derivations/ssa.py` computes from the contract in a `nominal` case on the sample portfolio, a `boundary` case on the limits of the time bands and an `invalid` case with a row for each input rule. The draft accuracy contract `ACC-03` covers the published quantities. No `independent` value | None: nothing was run in Excel ([issue 22](https://github.com/danielep71/K-FRTB-SSA/issues/22)). In LibreOffice, which is not a supported host, `tools/prototype/run_cases.py` finds every expected value in the workbook | Exercised |
| `CALC-04` | `illustrative` values that `tests/derivations/ssa.py` computes from the contract: a `nominal`, a `boundary` and an `invalid` case. No `published` or `independent` value | None: nothing was run in Excel ([issue 22](https://github.com/danielep71/K-FRTB-SSA/issues/22)). In LibreOffice, which is not a supported host, `tools/prototype/run_cases.py` finds every expected value in the workbook | Exercised |
| `CALC-05` | `illustrative` values that `tests/derivations/ssa.py` computes from the contract: a `nominal`, a `boundary` and an `invalid` case. No `published` or `independent` value | None: nothing was run in Excel ([issue 22](https://github.com/danielep71/K-FRTB-SSA/issues/22)). In LibreOffice, which is not a supported host, `tools/prototype/run_cases.py` finds every expected value in the workbook | Exercised |
| `CALC-06` | Two `published` values, the overall net open position and the requirement of the example of table 9 of MAR40.61 of `SRC-02`, in the `nominal` case `mar40-61-table-9`. `illustrative` values that `tests/derivations/ssa.py` computes from the contract in a `nominal`, a `boundary` and an `invalid` case. The draft accuracy contract `ACC-02` covers its two totals | None: nothing was run in Excel ([issue 22](https://github.com/danielep71/K-FRTB-SSA/issues/22)). In LibreOffice, which is not a supported host, `tools/prototype/run_cases.py` finds every expected value in the workbook | Exercised |
| `CALC-07` | `illustrative` values that `tests/derivations/ssa.py` computes from the contract: a `nominal` case. A charge that is not a number or is negative stops the run, which no fixture expresses: the project's suites have no assertion on it either | None: nothing was run in Excel ([issue 22](https://github.com/danielep71/K-FRTB-SSA/issues/22)). In LibreOffice, which is not a supported host, `tools/prototype/run_cases.py` finds every expected value in the workbook | Exercised |
| `CALC-08` | `illustrative` values that `tests/derivations/ssa.py` computes from the contract: a `nominal` case on the sample portfolio and a `boundary` case with positions in one risk class. No `published` or `independent` value | None: nothing was run in Excel ([issue 22](https://github.com/danielep71/K-FRTB-SSA/issues/22)). In LibreOffice, which is not a supported host, `tools/prototype/run_cases.py` finds every expected value in the workbook | Exercised |

The *Status* of a calculation is the least advanced
[traceability](#traceability) status of its rules: *Validated* only when every
rule is.

Cover, for each calculation, with the case categories of the
[test-case format](TEST_CASES.md#categories):

- typical cases (`nominal`) and the boundaries of every rule (`boundary`);
- invalid and missing input, which the [data contract](DATA_CONTRACT.md)
  rejects (`invalid`).

Cover with VBA tests, which add no category
([relation, end-to-end and volume tests](TEST_CASES.md#other-tests)):

- failure: an operation that fails clears its earlier results and marks them
  `FAILED` ([status and staleness](WORKBOOK_CONTRACT.md#status-and-staleness));
- repeatability: the same input gives the same output.

A relation test or the end-to-end case runner of the
[harness guide](TEST_HARNESS.md) can check either.

## Traceability

Each rule is traced from the paragraph of its source to the step of the
calculation contract that states it, to the VBA procedure that implements it
and to the test cases that test it, with its validation status:

| Source | Paragraph | Calculation step | VBA procedure | Test cases | Status |
| --- | --- | --- | --- | --- | --- |
| *Illustrative:* `SRC-01` | Article 1(1) | `CALC-01.1` | — | `illustrative-format-example`, `illustrative-file-import-example` | Not implemented |
| `SRC-02` | MAR40.38 | `CALC-02.1` | `CORE_InterestRate.CalculateInterestRate` | `illustrative-ir-specific-risk-maturity-limits`, `illustrative-sample-ir-specific-risk` | Exercised |
| `SRC-02` | MAR40.5 | `CALC-02.2` | `CORE_InterestRate.CalculateInterestRate` | `illustrative-ir-specific-risk-maturity-limits`, `illustrative-sample-ir-specific-risk` | Exercised |
| `SRC-02` | MAR40.6, table 1 | `CALC-02.3` | `CORE_InterestRate.SpecificRiskBand`, `CORE_InterestRate.SpecificRiskKey`, `CORE_Parameters.BuildSpecificRiskTable` | `illustrative-ir-specific-risk-maturity-limits`, `illustrative-sample-ir-specific-risk` | Exercised |
| `SRC-02` | MAR40.4 and MAR40.5 | `CALC-02.4` | `CORE_InterestRate.CalculateInterestRate` | `illustrative-ir-specific-risk-maturity-limits`, `illustrative-sample-ir-specific-risk` | Exercised |
| `SRC-02` | MAR40.24 | `CALC-03.1` | `CORE_InterestRate.CalculateInterestRate` | `bcbs119-c2-maturity-method`, `illustrative-sample-ir-general-market-risk` | Exercised |
| `SRC-02` | MAR40.25 | `CALC-03.2` | `CORE_InterestRate.CalculateInterestRate` | `bcbs119-c2-maturity-method`, `illustrative-ir-time-band-limits`, `illustrative-sample-ir-general-market-risk` | Exercised |
| `SRC-02` | MAR40.26, table 4 | `CALC-03.3` | `CORE_InterestRate.TimeBand`, `CORE_InterestRate.IsHighCoupon`, `CORE_Parameters.BuildLadder` | `bcbs119-c2-maturity-method`, `mar40-27-vertical-disallowance`, `illustrative-ir-time-band-limits`, `illustrative-sample-ir-general-market-risk` | Exercised |
| `SRC-02` | MAR40.26 | `CALC-03.4` | `CORE_InterestRate.CalculateInterestRate` | `bcbs119-c2-maturity-method`, `mar40-27-vertical-disallowance`, `illustrative-ir-time-band-limits`, `illustrative-sample-ir-general-market-risk` | Exercised |
| `SRC-02` | MAR40.27 | `CALC-03.5` | `CORE_InterestRate.LadderCharges` | `bcbs119-c2-maturity-method`, `mar40-27-vertical-disallowance`, `illustrative-ir-time-band-limits`, `illustrative-sample-ir-general-market-risk` | Exercised |
| `SRC-02` | MAR40.28, table 5 | `CALC-03.6` | `CORE_InterestRate.LadderCharges` | `bcbs119-c2-maturity-method`, `illustrative-ir-time-band-limits`, `illustrative-sample-ir-general-market-risk` | Exercised |
| `SRC-02` | MAR40.28, table 5 | `CALC-03.7` | `CORE_InterestRate.LadderCharges`, `CORE_InterestRate.Offset` | `bcbs119-c2-maturity-method`, `illustrative-ir-time-band-limits`, `illustrative-sample-ir-general-market-risk` | Exercised |
| `SRC-02` | MAR40.23 | `CALC-03.8` | `CORE_InterestRate.LadderCharges` | `bcbs119-c2-maturity-method`, `mar40-27-vertical-disallowance`, `illustrative-ir-time-band-limits`, `illustrative-sample-ir-general-market-risk` | Exercised |
| `SRC-02` | MAR40.23 | `CALC-03.9` | `CORE_InterestRate.LadderCharges` | `bcbs119-c2-maturity-method`, `mar40-27-vertical-disallowance`, `illustrative-ir-time-band-limits`, `illustrative-sample-ir-general-market-risk` | Exercised |
| `SRC-02` | MAR40.46 | `CALC-04.1` | `CORE_Equity.CalculateEquity` | `illustrative-eq-matched-positions`, `illustrative-sample-equity` | Exercised |
| `SRC-02` | MAR40.42 and MAR40.43 | `CALC-04.2` | `CORE_Equity.CalculateEquity` | `illustrative-eq-matched-positions`, `illustrative-sample-equity` | Exercised |
| `SRC-02` | MAR40.47, and table 8 of MAR40.52 | `CALC-04.3` | `CORE_Equity.CalculateEquity` | `illustrative-eq-matched-positions`, `illustrative-sample-equity` | Exercised |
| `SRC-02` | MAR40.42 and MAR40.43 | `CALC-04.4` | `CORE_Equity.CalculateEquity` | `illustrative-eq-matched-positions`, `illustrative-sample-equity` | Exercised |
| `SRC-02` | MAR40.42 | `CALC-04.5` | `CORE_Equity.CalculateEquity` | `illustrative-eq-matched-positions`, `illustrative-sample-equity` | Exercised |
| `SRC-02` | MAR40.68 and MAR40.73 | `CALC-05.1` | `CORE_Commodity.CalculateCommodity` | `illustrative-co-matched-positions`, `illustrative-sample-commodities` | Exercised |
| `SRC-02` | MAR40.67 | `CALC-05.2` | `CORE_Commodity.CalculateCommodity` | `illustrative-co-matched-positions`, `illustrative-sample-commodities` | Exercised |
| `SRC-02` | MAR40.72 | `CALC-05.3` | `CORE_Commodity.CalculateCommodity` | `illustrative-co-matched-positions`, `illustrative-sample-commodities` | Exercised |
| `SRC-02` | MAR40.73 | `CALC-05.4` | `CORE_Commodity.CalculateCommodity` | `illustrative-co-matched-positions`, `illustrative-sample-commodities` | Exercised |
| `SRC-02` | MAR40.72 and MAR40.73 | `CALC-05.5` | `CORE_Commodity.CalculateCommodity` | `illustrative-co-matched-positions`, `illustrative-sample-commodities` | Exercised |
| `SRC-02` | MAR40.60 | `CALC-06.1` | `CORE_ForeignExchange.CalculateForeignExchange` | `illustrative-fx-one-sided-positions`, `illustrative-sample-fx`, `mar40-61-table-9` | Exercised |
| `SRC-02` | MAR40.60 | `CALC-06.2` | `CORE_ForeignExchange.NetOpenPositionCharge` | `illustrative-fx-one-sided-positions`, `illustrative-sample-fx`, `mar40-61-table-9` | Exercised |
| `SRC-02` | MAR40.61 | `CALC-06.3` | `CORE_ForeignExchange.NetOpenPositionCharge` | `illustrative-fx-one-sided-positions`, `illustrative-sample-fx`, `mar40-61-table-9` | Exercised |
| `SRC-04` | Article 351 | `CALC-06.4` | `CORE_ForeignExchange.DeMinimisWaived`, `CORE_Engine.Calculate`, `CORE_Parameters.BuildRegime` | — | Not validated |
| `SRC-02` | MAR40.74 to MAR40.85 | `CALC-07.1` | `CORE_Parameters.BuildOptionCharges` | `illustrative-option-charges` | Exercised |
| `SRC-02` | MAR40.2 and MAR40.75 | `CALC-07.2` | `CORE_Engine.ResultsTable` | — | Not validated |
| `SRC-02` | MAR40.2 and MAR40.4 | `CALC-08.1` | `CORE_Engine.ResultsTable` | `illustrative-requirement-fx-and-options-only`, `illustrative-sample-requirement` | Exercised |
| `SRC-02` | MAR40.2 | `CALC-08.2` | `CORE_Engine.ResultsTable` | `illustrative-requirement-fx-and-options-only`, `illustrative-sample-requirement` | Exercised |
| `SRC-02` | MAR40.2 | `CALC-08.3` | `CORE_Engine.ResultsTable` | `illustrative-requirement-fx-and-options-only`, `illustrative-sample-requirement` | Exercised |
| `SRC-02` | MAR40.1 | `CALC-08.4` | `CORE_Engine.ResultsTable` | `illustrative-requirement-fx-and-options-only`, `illustrative-sample-requirement` | Exercised |
| `SRC-06` | Article 495s(2) | `CALC-08.5` | `CORE_Engine.ResultsTable`, `CORE_Parameters.BuildRegime`, `CORE_Parameters.TransitionalFactor` | — | Not validated |

- **Source and paragraph:** a `SRC-` ID of the [source register](SOURCES.md)
  and the exact paragraph, as in the step's *Source* column.
- **Calculation step:** `CALC-NN.<step>`, a step of the *Rules* table of
  that [calculation contract](CALCULATION_CONTRACTS.md). A rule implemented
  in several steps has one row per step.
- **VBA procedure:** `Module.Procedure` of a component under `src/`, or `—`
  while none implements the rule. Name every procedure that applies the rule.
- **Test cases:** the IDs of the [test cases](TEST_CASES.md) that test the
  step, or `—`. Each case tests the row's calculation.
- **Status:** one of the four below.

| Status | Meaning | The row names |
| --- | --- | --- |
| Validated | A `published` or `independent` case passes for the rule on an Excel run of the release candidate ([completeness policy](TEST_CASES.md#completeness-policy)) | A procedure, and a case with a `published` or `independent` output whose `rule` is the row's step |
| Exercised | Cases run the rule, but none is a `published` or `independent` case that passed: they are `illustrative`, or the reference cases have not passed yet | A procedure and at least one case |
| Not validated | Implemented, but no case tests it | A procedure and no case |
| Not implemented | No procedure implements it yet; its cases may come first | No procedure |

`python tools/check_test_cases.py`, which **Repository integrity** runs,
fails a row whose source is not in the register, whose step is not one that
the *Rules* table of a declared calculation lists, whose procedure is not in
a component under `src/`, whose case does not exist or tests another
calculation, or whose status contradicts what the row names. It cannot see an Excel run: whether a
`Validated` rule's case passed on the release candidate is the
[evidence record](#evidence-record)'s to show.

## Evidence record

Each run of a check in Excel records:

- the commit of the source and the workbook it was built into;
- the case, its expected value, that value's provenance and the tolerance;
- the observed value and the outcome;
- the Excel build, its bitness and the Windows version;
- who ran it, who reviewed it, and when.

A check that was not run is recorded as `NOT RUN` with the reason. It is
never a pass.

An Excel run of the harness keeps the host, who ran it, when and how it
ended in its [host record](TEST_HARNESS.md#host-record), also when it stopped
early. `tools/check_excel_run.py` rejects a record that lacks one of them or
leaves who ran it, when or how it ended `UNKNOWN`, and accepts an outcome
other than `PASS` only as evidence of that outcome.

### What an accuracy contract proves

A frozen [accuracy contract](TEST_CASES.md#accuracy-contracts) whose results
pass on an Excel run of the candidate shows that, on that host, the worst
error of the calculation over its calibration and holdout cases stays within
the threshold. The holdout cases played no part in setting the threshold,
and their values are `published` or `independent`. It proves nothing outside
those cases, regime and host. A draft contract, and any contract on
`illustrative` values, validates nothing. Record each contract's worst errors
with the evidence.

<a id="host-behaviour"></a>

## Host behaviour

The code relies on how Excel and the VBE behave: which settings an operation
changes, how a name is re-cased, what an import needs, how a window or an
event behaves. Such behaviour can differ by Excel build, channel, bitness and
policy. Record each behaviour that the code or a guide relies on here, with
whether it was measured:

| ID | Behaviour | Status | Probe | Host | Date | Result | Issue |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `HB-01` | The VBE re-cases every use of a name to its latest spelling, so an exported module can differ from its source only in the letter case of code outside strings and comments | `assumed` | The next Excel run by hand: the `Note:` lines of `tools/check_excel_run.py` | — | — | Not measured: no Excel run has recorded a re-cased export yet. `tools/check_excel_run.py` tolerates it: it compares that code without letter case and reports each re-cased identifier as a note | — |
| `HB-02` | Excel started through COM can follow the user's macro setting (`AutomationSecurity` is `msoAutomationSecurityByUI`, set by policy). A workbook it opens then has its macros disabled, and `Application.Run` fails with *Excel 4.0 (XLM) macros have been disabled*, because the VBA procedure is not found | `measured` | By hand: in PowerShell, open the built workbook through COM and `Run` `'<workbook>'!TEST_Harness.LastResult` | Microsoft 365 Excel 16.0 build 20430, channel not recorded, 64-bit; Windows 11 Enterprise 25H2, build 26200.9711 | 2026-11-03 | `AutomationSecurity` was `ByUI`, and `Run` failed with that message. `tools/devbuild/excel_run.py` changes no macro setting: it waits for **Enable Content** | — |
| `HB-03` | `Application.Run` reaches a public procedure of an `Option Private Module` module when the name is qualified with the workbook (`'Book.xlsm'!Module.Procedure`) | `measured` | As `HB-02`, in a session whose macros were enabled | As `HB-02` | 2026-11-03 | `'Workbook-development.xlsm'!TEST_Harness.LastResult` ran | — |
| `HB-04` | In a workbook opened through COM with its macros disabled, **Enable Content** enables them without reopening, and `Application.Run` then reaches its procedures | `measured` | `python tools/devbuild/excel_run.py` on a host with `HB-02`: click **Enable Content** in the yellow bar | As `HB-02` | 2026-11-03 | Excel showed the yellow bar; after the click, the certification ran to `excel-run: PASS` without reopening the workbook | — |
| `HB-05` | The project that `tools/devbuild/vba_project.py` writes holds no compiled code (`_VBA_PROJECT` version `0xFFFF`, no performance cache), so Excel reads it from source and compiles each module when it is first used: the code that runs is the source the build wrote, and no Trust access to the VBA project object model is needed | `assumed` | By hand: open a development build in Excel, check that `? TEST_Harness.LastResult` runs in the Immediate window, which compiles that module on its first use, then check that **Debug → Compile VBAProject**, which compiles the other modules, completes | — | — | Not measured by a probe of its own. `tools/devbuild/build_workbook.py` relies on it, and so does the identity by rebuild of `tools/check_excel_run.py`, which proves that the workbook Excel ran holds the source of the commit, not that Excel compiled it | — |
| `HB-06` | **End** in an error dialog, or **Run → Reset** in the VBE, resets the project and clears every module-level variable, so the harness's active run and recovery flag are lost and the stop is not detected | `assumed` | By hand: set a breakpoint on the line `TEST_Harness.BeginCase "text-exact"` of `TEST_Examples.CaseExactText`, type `TEST_Examples.RunExampleTests "<fixtures>"` with the full path of `tests/fixtures/` in the Immediate window, and when the run pauses there, choose **Run → Reset**; then check that `? TEST_Harness.LastResult` is empty and that the next run is not refused | — | — | Not measured. The harness guides tell you to check Excel's state by hand after such a stop | — |
| `HB-07` | The Immediate window keeps only its last 200 lines | `assumed` | By hand: print 250 numbered lines with `Debug.Print`, then check that the first line left is line 51 | — | — | Not measured. The installation guide tells you to save the whole log with `TEST_Harness.SaveLog` | — |
| `HB-08` | An error value typed in an input cell, or given by `CVErr`, is an error value through `Range.Value`, which `PlainText` reads as the text `#ERROR`; the text `#N/A` written into a cell formatted as text stays a text | `assumed` | The cases `error-values-in-inputs` and `reading-an-error-value` of the `probes` suite of `tests/modules/TEST_Probes.bas` | — | — | Not measured in Excel. In LibreOffice 24.2, which is not a supported host, a cell given an error value by VBA stays empty and a typed error value reads as 0 in the engine, so eight assertions of the two cases fail there | [22](https://github.com/danielep71/K-FRTB-SSA/issues/22) |
| `HB-09` | A text that reads as a number, written as it is into a cell that is not formatted as text, becomes a number; written with the cell formatted as text, and the cell given its format back, it keeps its digits | `assumed` | The case `text-written-into-a-cell` of the `probes` suite | — | — | Not measured in Excel. It passes in LibreOffice 24.2 | [22](https://github.com/danielep71/K-FRTB-SSA/issues/22) |
| `HB-10` | After a run, `E1` of each calculation sheet and of `Checks` shows the status line of `Results` | `assumed` | The case `status-on-the-other-sheets` of the `probes` suite | — | — | Not measured in Excel. It passes in LibreOffice 24.2 | [22](https://github.com/danielep71/K-FRTB-SSA/issues/22) |
| `HB-11` | `M_Main.CheckStatus_OnEvent` leaves a workbook with no unsaved change saved when the results are current, and leaves an unsaved change when it writes `OUTDATED:` before the status line | `assumed` | The case `save-state-after-the-check` of the `probes` suite | — | — | Not measured in Excel. It passes in LibreOffice 24.2 | [22](https://github.com/danielep71/K-FRTB-SSA/issues/22) |
| `HB-12` | `Workbook_BeforeSave` and `Workbook_Open` run the status check: saving a workbook whose input changed after a run writes `OUTDATED:`; opening a workbook that another application saved does the same; opening and closing a workbook whose inputs did not change asks for no save | `assumed` | By hand: [the host checks](../../INSTALLATION.md#host-checks-by-hand), steps 1 to 3 | — | — | Not measured: no procedure can raise the events, and none was exercised in Excel. In LibreOffice the tests call the procedure directly | [22](https://github.com/danielep71/K-FRTB-SSA/issues/22) |
| `HB-13` | The `Value` and the `Formula` of a range of several cells are tables with bounds from 1, and those of one cell are not | `assumed` | The case `areas-as-tables` of the `probes` suite | — | — | Not measured in Excel. It passes in LibreOffice 24.2 | [22](https://github.com/danielep71/K-FRTB-SSA/issues/22) |
| `HB-14` | A text written with a leading apostrophe into a cell that is formatted as text keeps the apostrophe in its value | `assumed` | By hand: format a cell as text, write `'IT` to its `Value` in the Immediate window, and read `? Range("A1").Value` | — | — | Not measured. The `cases` suite does not write an apostrophe, because of it: it writes a text with the cell formatted as text ([harness guide](TEST_HARNESS.md#the-projects-suites)) | [22](https://github.com/danielep71/K-FRTB-SSA/issues/22) |
| `HB-15` | A write by code to a locked cell of a protected sheet is refused, with error 1004; after `Protect` with `UserInterfaceOnly:=True` and the template's options it is accepted, the sheet stays protected for users, and `Protect` of a sheet that was protected without a password needs no password | `assumed` | The case `sheet-protection` of the `probes` suite, and [the host checks](../../INSTALLATION.md#host-checks-by-hand), step 6 | — | — | Not measured in Excel. LibreOffice 24.2, which is not a supported host, ignores `UserInterfaceOnly` and does not know the options of `Protect`, and refuses to clear a protected sheet: `M_Main.AllowWriting` unprotects the sheet there, so three assertions of the case fail | [26](https://github.com/danielep71/K-FRTB-SSA/issues/26) |
| `HB-16` | `UserInterfaceOnly` is not saved with the workbook, so a reopened workbook refuses a write by code until `AllowWriting` runs, and `Protect` called again marks the workbook as changed, which `CheckStatus_OnEvent` undoes when it wrote nothing | `assumed` | By hand: [the host checks](../../INSTALLATION.md#host-checks-by-hand), step 6, then the case `save-state-after-the-check` of the `probes` suite | — | — | Not measured in Excel | [26](https://github.com/danielep71/K-FRTB-SSA/issues/26) |

- **Status:** `measured` when a probe ran on the named host and its output is
  kept; `assumed` otherwise.
- **Probe:** a `TEST_` procedure under `tests/modules/` that observes the
  behaviour and prints what it saw, run with the VBA test harness. A
  measurement found by hand names the steps instead.
- **Host:** the Excel product, version and build, its channel and bitness,
  and the Windows version, as in the [evidence record](#evidence-record).
- **Result:** what was observed, and the model it supports. A measurement
  holds for the named host only; it is never generalised to others.

A guide that states host behaviour the code relies on cites its `HB-` entry.
An `assumed` entry is listed under *Limits* in the
[release certification](../release/RELEASE_CERTIFICATION.md) until it is
measured.

## Open decisions

| Decision | Options | Issue |
| --- | --- | --- |
| *Illustrative:* where the reference values that validate `CALC-01` come from | A worked example of a registered source (`published`); an external implementation by another author (`independent`, `IMPL-`) | Link |

None is open for the project's calculations: the choices are recorded below.

## Decisions

| Date | Decision | Reason | Issue |
| --- | --- | --- | --- |
| 2026-11-19 | `CALC-03`, general market risk by the maturity method, takes its `published` values from two sources: the example of MAR40.27 of `SRC-02`, and the worked example C.2 of `SRC-03`, the 2005 update of the 1996 Amendment, which the Basel Framework does not repeat | MAR40.27 prints the weighted amounts and the disallowance of one band; C.2 prints the weighted position of every leg and each charge of a portfolio of four instruments. Its ladder weights, disallowance percentages and zones for a coupon of 3% or more are those of MAR40.26 to MAR40.28, compared on 2026-11-19. Each is a `nominal` case, `mar40-27-vertical-disallowance` and `bcbs119-c2-maturity-method`; the draft `ACC-03` covers them | [47](https://github.com/danielep71/K-FRTB-SSA/issues/47) |
| 2026-11-19 | `CALC-06` keeps the example of table 9 of MAR40.61 as its only `published` value, for its steps 2 and 3. Step 1, the net position of a currency, has none | Table 9 prints the net positions as inputs, already in one currency: the sum of the positions of a currency is not printed anywhere in `SRC-02` or `SRC-03`. Its rules stay *Exercised* | [47](https://github.com/danielep71/K-FRTB-SSA/issues/47) |
| 2026-11-19 | `CALC-02`, `CALC-04`, `CALC-05`, `CALC-07` and `CALC-08` have no `published` value, and their rules stay *Exercised* | `SRC-02` prints no worked example of specific risk, of the equity or the commodity charge, of the scaling of the four classes or of the risk-weighted assets, and example C.3 of `SRC-03` is the maturity ladder approach for commodities and C.4 the delta-plus method, which the workbook does not calculate. The example of MAR40.76, a put option on shares, is the simplified approach for options, which the workbook does not calculate either: the user enters the non-delta charge | [47](https://github.com/danielep71/K-FRTB-SSA/issues/47) |
| 2026-11-19 | No `independent` implementation is registered, and no `IMPL-` ID exists | I searched the web on 2026-11-19 for another author's implementation of MAR40, open source or published, and found none that implements the simplified standardised approach with test cases; the ones returned cover credit risk or the model-based approach. The search was not exhaustive. Nothing is registered until one is found; `tools/check_independent.py` has nothing to run | [47](https://github.com/danielep71/K-FRTB-SSA/issues/47) |
| 2026-11-19 | A contract is frozen only where three holdout cases hold `published` or `independent` values, and none is frozen | `CALC-03` has two such cases and `CALC-06` one. The accuracy contracts `ACC-02` and `ACC-03` stay drafts, and no release may publish an accuracy for the calculation ([accuracy contracts](TEST_CASES.md#accuracy-contracts)) | [47](https://github.com/danielep71/K-FRTB-SSA/issues/47) |
| 2026-11-20 | The CRR regime is the simplified standardised approach of Article 325(2) of Regulation (EU) No 575/2013 as amended by Regulation (EU) 2024/1623, with Chapters 2 to 4 of Title IV of Part Three (Articles 326 to 361) and the factors of Article 325(2). The alternative standardised approach, the internal model approach, securitisation positions (Articles 337 and 338), CVA risk and the supervisory reporting templates are outside it | The simplified approach is the one the prototype's Basel rules correspond to: its four factors, 1.3, 3.5, 1.9 and 1.2, are those of Article 325(2). Everything else is a different calculation, and the data contract excludes securitisation positions | [74](https://github.com/danielep71/K-FRTB-SSA/issues/74) |
| 2026-11-20 | The text the regime follows is the CRR as consolidated, Regulation (EU) 2024/1623 and Delegated Regulation (EU) 2026/1221, each with its version and date in the [register](SOURCES.md) | Delegated Regulation (EU) 2026/1221 adds Articles 495s and 495u after the consolidated version of 2026-01-01, so no single text holds the regime. The register records the version of each text and whether the delegated act is in force | [74](https://github.com/danielep71/K-FRTB-SSA/issues/74) |
| 2026-11-20 | One workbook holds both regimes, selected by a `Regime` parameter. The Basel rules stay unchanged, and a rule of the CRR applies only when the parameter selects it | The engine, the data contract and the test harness are shared. The risk is that a rule of one regime changes the results of the other, which the registered Basel cases guard against: their expected values must not change | [74](https://github.com/danielep71/K-FRTB-SSA/issues/74) |
| 2026-11-20 | The multiplier of 0.9 of Article 495s(2) is a parameter with its end date, 2029-12-31, and the calculation date is an input of the workbook. `Results` shows the requirement before and after the multiplier | The multiplier applies until 2029-12-31, so the workbook must know the date of the calculation to apply it, and a reader must see what the multiplier changes | [74](https://github.com/danielep71/K-FRTB-SSA/issues/74) |
| 2026-11-20 | The size conditions of Article 325a (10% of total assets and EUR 500 million) and the derogation of Article 94 (5% of total assets and EUR 50 million) are calculated in the workbook, as a calculation contract of their own, with total assets and total own funds as inputs | Whether the approach may be used decides whether the result is meant to be used at all, and the de minimis of Article 351 needs total own funds in any case. The calculation takes the monthly series, the three consecutive months and the six of twelve of Article 325a(5) | [74](https://github.com/danielep71/K-FRTB-SSA/issues/74) |
