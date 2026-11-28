# 📜 Changelog

Notable changes to this repository, from a user's point of view. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions
follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It is
neither a commit log nor release evidence. **No version has been released.**

## 🧭 Maintenance policy

- Add a material change under `[Unreleased]` in the pull request that makes
  it. A change nobody would notice, such as a typo or a refactor with no
  effect, needs no entry.
- Describe the observable result: what exists, what a user or contributor
  can now do, what changed in a contract, and what to do about it.
- Link the issue or pull request that holds the detail.
- Record only evidence that was produced. A static check is not an Excel run:
  say what was and was not run in Excel.
- Once a version has been released, edit an existing entry of the next
  version rather than adding one that describes a change to its unreleased
  work: an entry describes what the release contains. The first release's
  section also records, as **Changed** and **Fixed**, how its additions
  changed while it was prepared.
- Entries move into a dated section only when a release is prepared, as
  [RELEASING.md](RELEASING.md) describes. After that, until the tag exists,
  a fix for a release gate that failed on the candidate records its change in
  the dated section being released
  ([step 2](RELEASING.md#2-integrate-and-freeze-the-candidate)). A published section is not
  rewritten; a demonstrable error is corrected in review, with a line
  `Corrected after release on YYYY-MM-DD: <reason>` in that section.

<a id="date-and-version-rules"></a>

### Date and version rules

`python tools/check_changelog.py` checks the headings, dates, order, link
references and `VERSION`, and that this introduction, the *Current status* of
[RELEASING.md](RELEASING.md#current-status) and the README describe the same
release state; **Repository integrity** also checks that every link reference
points to this repository. Once a release's tag exists, its section must equal
that section at the tag, unless it holds the dated correction line above; a
correction is reported as a note. What a version number means is checked
in review.

| Rule | Standard |
| --- | --- |
| Version | `MAJOR.MINOR.PATCH`, without a leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD`, with a real calendar date |
| Order | `[Unreleased]` first, then releases from newest to oldest, by version and date |
| Link references | One for `[Unreleased]` and one for every release; a release's points to `.../releases/tag/vX.Y.Z` |
| `0.y.z` | Initial development: no compatibility is promised between versions |
| Patch, minor, major | From `1.0.0`: a compatible correction, a compatible capability, an incompatible change to a calculation result, an input or output layout, or the supported VBA API |

| Category | Use for |
| --- | --- |
| **Added** | New capabilities, files, checks or tests |
| **Changed** | Changes to existing behaviour, contracts, tooling or documentation |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Validation** | Evidence actually produced |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories an entry needs.

## [Unreleased]

### Added

- `python tools/devbuild/excel_run.py --smoke` runs the example suite and every
  project suite that `.github/excel-evidence-policy.json` does not mark
  `"smoke": false`, which today leaves out the `cases` suite only. In
  LibreOffice that suite is 297 s of the 314 s of the whole certification, and the
  first full run in Excel took 7 min 44 s. The smoke run records into
  `test-results\excel-smoke\`, checks the logs it has against the counts of the
  policy and reports itself as `SMOKE RUN: not release evidence`; the full run,
  without `--smoke`, stays what a release needs, and `check_excel_run.py` fails on
  a smoke folder. Nothing was run in Excel
  ([#132](https://github.com/danielep71/K-FRTB-SSA/issues/132)).
- A step in the host checks by hand of `INSTALLATION.md` that compares the
  engine's run in Excel on the sample data with the formula workbook that Excel
  calculates, area by area, with `tools/prototype/compare_outputs.py`, so that a
  difference from the LibreOffice run has a procedure to find it. Nothing was
  run in Excel ([#124](https://github.com/danielep71/K-FRTB-SSA/issues/124)).
- FX table stop rules in the derivation script: `tests/derivations/ssa.py` now
  stops the run, as the contract states, on a rate that is not a number greater
  than 0 (`C07`), a currency listed twice (`C06`) and a reporting currency that
  is not listed or not at rate 1 (`C05`), and counts each in the `STOPPED` status
  line as the engine does. Five registered `invalid` cases cover them, one for
  each way a row can break, which raises the registered cases to 49 and the
  assertions of the workbook's `cases` suite to 874. The generated cases now
  mutate the `fx_rate` field too. The cases agree with the engine in LibreOffice
  and nothing was run in Excel
  ([#119](https://github.com/danielep71/K-FRTB-SSA/issues/119)).
- Generated cases: `tests/test_generated_cases.py` generates cases from the
  registered fixtures of `CALC-02` to `CALC-08` with fixed seeds and checks
  that the derivation script keeps five properties on them: the order of the
  records, the sign of a charge, its scale, the direction of the positions and
  an unusable position value that gives a status and never an exception. It
  needs no new dependency and runs in under a second. It found that the
  derivation script ignores the FX rate stop rules `C05` to `C07`, which
  [#119](https://github.com/danielep71/K-FRTB-SSA/issues/119) records. I did
  not add a fuzzer: the OpenSSF Scorecard *Fuzzing* check does not detect
  this approach, and `docs/methodology/TEST_CASES.md` states why. The test
  covers the derivation script, not the VBA engine, and nothing was run in
  Excel ([#69](https://github.com/danielep71/K-FRTB-SSA/issues/69)).
- Registered test cases of the `CRR` regime, with `illustrative` values:
  `illustrative-crr-ir-specific-risk`, a row at each risk weight of
  Article 336; `illustrative-crr-maturity-method-c2`,
  `illustrative-crr-sample-equity` and `illustrative-crr-sample-requirement`,
  the positions of a Basel case run under the CRR, so that each value that
  differs is the effect of a rule of the CRR (150% between zones 1 and 3, no
  charge on an index, the transitional multiplier); the de minimis of
  Article 351 at and just below the overall net open position; the
  transitional multiplier on its last date and the day after; and an
  `invalid` case on the category and the risk weight. The invalid interest
  rate case has a row with the category `CRR` under the Basel regime, and the
  requirement cases a new output, `transitional_factor`. Four cases test the
  parameters that stop the run: a regime that is neither `Basel` nor `CRR`,
  and under the `CRR` no total own funds and total own funds of 0, each with
  the new output `run_status`, the status line of `Results` without the time
  of the run; and the regime written `crr`, which is read as `CRR`. No
  existing expected value changed. `tests/derivations/ssa.py` now applies the
  rules of the CRR regime and the stops, and the `cases` suite has 42 cases
  and 864 assertions. The draft accuracy contract `ACC-04` covers the
  requirement of the CRR regime; it cannot be frozen, because the search
  recorded in `SOURCES.md` found no published worked example and no
  independent implementation of the CRR's simplified standardised approach.
  Nothing was run in Excel
  ([#86](https://github.com/danielep71/K-FRTB-SSA/issues/86)).
- The book of a position: a new input column *Book*, `Trading` or
  `Non-trading`, before *Comment* on `IRPositions`, `EQPositions`,
  `COPositions` and `FXPositions`; blank is `Trading`, so existing sheets and
  cases give the same results. Under both regimes a non-trading book FX or
  commodity row is calculated as a trading book one, at the value entered, as
  Article 325(1) of the CRR and MAR12.6 put that risk in the scope; a
  non-trading book interest rate or equity row is excluded before its other
  tests and counted by the new warning rule `C17` of `Checks`, and the run
  stays `OK`. Any other book rejects the row. The technical standards of
  Article 325(9), Delegated Regulations (EU) 2023/1577 and 2025/878, are now
  registered (`SRC-09`, `SRC-10`), with chapters RBC25 and MAR12 of the Basel
  Framework (`SRC-11`, `SRC-12`): the standards value those positions for the
  alternative approaches only, so the value entered is an assumption under the
  simplified standardised approach. Article 495u is documented, not a
  parameter. The input columns right of the book, *Comment*, moved one column
  right, and the `core` suite now has 17 cases and 230 assertions. Nothing was
  run in Excel ([#85](https://github.com/danielep71/K-FRTB-SSA/issues/85)).
- Four registered test cases of the book, `illustrative-ir-book`,
  `illustrative-eq-book`, `illustrative-co-book` and `illustrative-fx-book`:
  in each risk class a trading book row, a row with no book and non-trading
  book rows, which are excluded and offset nothing for interest rate and
  equity, and are calculated and offset the others for commodities and FX.
  The four `invalid` cases of the position sheets have two more rows: a book
  that is neither value, and a non-trading book row that breaks a later rule,
  which interest rate and equity exclude and commodities and FX reject. Their
  values are `illustrative`, and the `cases` suite now has 29 cases and 684
  assertions. Nothing was run in Excel
  ([#85](https://github.com/danielep71/K-FRTB-SSA/issues/85)).
- The eligibility for the simplified standardised approach, Article 325a of
  Regulation (EU) No 575/2013, and for the derogation for small trading book
  business, Article 94, under the `CRR` regime (`CALC-09`). On the new sheet
  `Eligibility` you enter, for up to 24 month-ends, the total assets and the
  aggregated long and short positions of each approach, in euro; the workbook
  tests the thresholds of each month (10% and EUR 500 million, 5% and EUR 50
  million, four new parameters on `Params`), the runs of three months and the
  counts of more than six in twelve, and gives each approach a status, `MET`,
  `NOT MET` or `CEASE`, and the restart condition of a year. A status `NOT MET`
  or `CEASE` is a warning of `Checks` (rule `C16`), and a month that fails a test
  is rejected (rule `C15`). The status changes no requirement and stops no run;
  the notifications and the other obligations of the two Articles are yours.
  Under the Basel regime the sheet is not read. The parameters, the ladder and the
  tables of `Params` below them moved down four rows. Nothing was run in Excel
  ([#84](https://github.com/danielep71/K-FRTB-SSA/issues/84)).
- Five registered test cases of `CALC-09`, `illustrative-eligibility-…` and
  `illustrative-invalid-eligibility-months`: the sample months, the share and
  the amount of each approach and one euro above them, the runs and counts that
  make an event, a month for each test a row can fail, and the Basel regime.
  Their values are `illustrative`, and every rule of `CALC-09` is now
  *Exercised*. A fixture may set the regime and the total own funds under the
  new optional key `parameters`, and lists the months of `DS-11` in the order
  of their rows; `tools/check_test_cases.py`, `tools/prototype/run_cases.py`
  and the `cases` suite read both, and the suite now has 25 cases and 637
  assertions. In LibreOffice every expected value is in the workbook; nothing
  was run in Excel ([#84](https://github.com/danielep71/K-FRTB-SSA/issues/84)).
- The options and the market of the equity risk under the `CRR` regime,
  compared with Delegated Regulations (EU) No 528/2014 and No 525/2014, whose
  Official Journal texts are now read and `verified` in the source register
  (`SRC-07`, `SRC-08`). Under the `CRR` regime you enter the delta-equivalent
  position of every option on the position sheets, whatever the approach, and
  the non-delta charge of Delegated Regulation (EU) No 528/2014 with the delta
  already deducted: the carve-out that the Basel regime asks for under the
  simplified and the scenario approaches would count the delta effect twice.
  You also enter one market for all the equities listed in the euro area, and
  one for each national jurisdiction elsewhere. The workbook checks neither,
  and nothing changes in the calculation. No test case has an expected value
  under the `CRR`, and nothing was run in Excel
  ([#83](https://github.com/danielep71/K-FRTB-SSA/issues/83)).
- The commodities risk under the `CRR` regime, compared with Articles 355 to
  361 of Regulation (EU) No 575/2013: the workbook calculates the simplified
  approach of Article 360, whose rates are those of the workbook, and nothing
  changes in the calculation. The maturity ladder of Article 359, the extended
  maturity ladder of Article 361 and the ancillary agricultural commodities
  business of Article 356 are not supported, and the positions that Article
  357(4) and (5) put in the same commodity or in the calculation are entered by
  you under one commodity name and as positions. No test case has an expected
  value under the `CRR`, and nothing was run in Excel
  ([#82](https://github.com/danielep71/K-FRTB-SSA/issues/82)).
- The equity risk of stock-index futures under the `CRR` regime: a position
  entered as `Index` has no charge of its own, where the Basel regime charges
  a further 2%, from the new parameter `EQIndexRateCRR` (0), because Article 344
  of Regulation (EU) No 575/2013 treats a stock-index future as one equity,
  whose general risk is in the net position of its market and whose specific
  risk is ignored where the future is exchange traded and represents a
  relevant, appropriately diversified index. A future that does not meet that
  condition is entered as a `Single name`. The collective investment
  undertakings of Article 348 are not supported. The results under Basel do not
  change, no registered test
  case has an expected value under the `CRR`, and nothing was run in Excel
  ([#81](https://github.com/danielep71/K-FRTB-SSA/issues/81)).
- The general market risk of debt instruments under the `CRR` regime: the
  offset between zones 1 and 3 of the maturity method is charged at 150%,
  Article 339(9), point (f), of Regulation (EU) No 575/2013, in place of the
  100% of MAR40.28, from the new parameter `HorizontalZone13CRR`. The bands,
  the weights and the other disallowances are the same as in MAR40, and the
  results under Basel do not change. The reversal of the first two offsets
  that paragraph 6 allows and the duration-based calculation of Article 340
  are not supported, and no registered test case has an expected value under
  the `CRR`; nothing was run in Excel
  ([#80](https://github.com/danielep71/K-FRTB-SSA/issues/80)).
- The specific risk of debt instruments under the `CRR` regime, by Article 336
  of Regulation (EU) No 575/2013: the category `CRR` of the specific-risk table,
  whose grade is the risk weight that the issuer or the security would receive
  under the standardised approach for credit risk (`RW 0%`, `RW 10%`, `RW 20%`,
  `RW 50%`, `RW 75%`, `RW 100%`, `RW 150%`), with 13 rows of rates: 0%, half of
  the second category for covered bonds at 10%, the second category by residual
  term at 20% and 50%, 8% and 12%. A weight of 75%, which the article does not
  name, is charged as 100%, by my decision. Under the `CRR` a position is `CRR`
  or `NONE`, and under Basel the category `CRR` is rejected. The cap of
  Article 335 and the other points of the article that need an input are not
  supported ([#79](https://github.com/danielep71/K-FRTB-SSA/issues/79)).
- The de minimis of Article 351 of Regulation (EU) No 575/2013 under the `CRR`
  regime: the FX requirement is zero while the overall net open position, gold
  included, does not exceed `FXDeMinimisRate`, 2%, of the new parameter
  `TotalOwnFunds`, which the `CRR` regime requires (a run without it stops).
  The treatment of gold stays that of MAR40.60, as an assumption: Articles 351
  and 352(4) do not settle it. The closely correlated currencies of Article 354
  and the exclusions of Article 104c are not supported
  ([#78](https://github.com/danielep71/K-FRTB-SSA/issues/78)).
- The regime of the workbook: a `Regime` parameter on `Params`, `Basel` by
  default, and `CRR`, which multiplies the requirement by the transitional
  multiplier of Article 495s(2) of Regulation (EU) No 575/2013 on a reporting
  date up to `TransitionalEndDate` (0.9 until 2029-12-31, parameters
  `TransitionalMultiplier` and `TransitionalEndDate`). `Results` has a row for
  the multiplier below the total, `SSA_Requirement` and `SSA_RWA` now refer to
  it, and `SSA_RequirementBeforeTransition` holds the total before it. Under
  `Basel` the factor is 1 and the results are those of before. An unknown regime,
  or a `CRR` without a last date, stops the run. Nothing else of the CRR is
  implemented: the other rules are those of MAR40, and nothing was run in
  Excel ([#77](https://github.com/danielep71/K-FRTB-SSA/issues/77)).
- A comparison of each rule of the calculation contracts with the CRR's
  simplified standardised approach, `docs/methodology/CRR_COMPARISON.md`: for
  each step of `CALC-02` to `CALC-08`, whether Article 325(2) and the articles
  it points to say the same, differ or add a rule, and the task that takes each
  difference. Among them, the CRR's rate of 150% between the zones 1 and 3 of
  the maturity method where the workbook has 100%, and the absence of a 2%
  index charge. It is a reading of the text: no rule of the CRR regime is
  implemented or validated ([#76](https://github.com/danielep71/K-FRTB-SSA/issues/76)).
- The texts of the CRR regime in the source register, `SRC-04` to `SRC-08`:
  the consolidated CRR of 2026-01-01, Regulation (EU) 2024/1623, Delegated
  Regulation (EU) 2026/1221 and the two delegated regulations of 2014 on the
  definition of *market* and the non-delta risk of options, with what was read
  and what was not, and a table of the technical standards that the CRR leaves
  to the EBA. No rule cites them yet, and nothing is validated by registering
  them ([#75](https://github.com/danielep71/K-FRTB-SSA/issues/75)).
- The scope of the CRR regime, recorded in the *Decisions* table of
  `docs/methodology/VALIDATION_PLAN.md`: the simplified standardised approach
  of Article 325(2) of Regulation (EU) No 575/2013 as amended by Regulation
  (EU) 2024/1623, selected by a `Regime` parameter of the one workbook, with
  the transitional multiplier of 0.9 of Article 495s(2) and a calculation of
  the size conditions of Article 325a. It is a decision only: no rule of the
  CRR is implemented or validated, and nothing was run in Excel
  ([#74](https://github.com/danielep71/K-FRTB-SSA/issues/74)).
- Two test cases with `published` values for the interest rate calculation,
  the maturity method, which had none: `mar40-27-vertical-disallowance`, the
  example of MAR40.27 (USD 100 million of weighted longs and USD 90 million of
  weighted shorts give a vertical disallowance of USD 9 million), and
  `bcbs119-c2-maturity-method`, the worked example C.2 of the Basel
  Committee's 2005 update of its 1996 market risk amendment, which the Basel
  Framework does not repeat: six legs, their weighted positions, the vertical
  and within-zone disallowances, the net position charge and a charge of
  $4,580,000. The amendment is registered as `SRC-03` for those printed
  values only, because the ladder weights, the disallowance percentages and
  the zones are those of MAR40.26 to MAR40.28. The engine reproduces them in
  LibreOffice, to a tolerance of 500 for the example that prints positions to
  0.01 mn. The validation plan records where each calculation's reference
  values come from: no other has a published example, no independent
  implementation was found, and the accuracy contract `ACC-03` is a draft, as
  a frozen one needs three holdout cases. Nothing was run in Excel (#47).
- The template protects its sheets. All 18 are protected, without a password,
  and only the input cells are unlocked: the yellow cells, and on `Params`
  the grey cells of the MAR40 constants, which a what-if may change. Users
  can select cells and change the width of a column and the height of a row,
  and cannot overwrite a label, a header, a formula or a result, insert or
  delete a row, sort or filter. The macros write the output sheets because
  each operation applies `UserInterfaceOnly` to them through the new
  `M_Main.AllowWriting`; the test modules apply it to every sheet they
  write. The protection matrix states the input and locked ranges of each
  sheet, `build_prototype.py` writes the template and the matrix from the
  colours of the sheets, and `tools/check_workbook_protection.py` compares
  them. A sheet that someone protected with a password makes an operation stop
  with an error that names it. Nothing was run in Excel: the behaviour is
  `HB-15` and `HB-16` of the validation plan, assumed, with a probe and a
  check by hand (#26).

- A guide to the prototype for the person who uses it, `docs/PROTOTYPE.md`:
  what the workbook calculates, what it needs, its 18 sheets and its inputs,
  how to run `RunSSA`, `ClearOutputs` and `CheckStatus` and read the status
  line, the four worksheet functions, how to build the workbook, what it does
  not do and what has been checked. The README and the installation guide
  link it. `SECURITY.md` now states, for each entry point of the workbook and
  for the project's test suites and the LibreOffice tools, what it reads,
  writes, saves or deletes, and that none starts a process or reaches the
  network. `INSTALLATION.md` records the host policy as a decision of the
  project, which keeps the defaults (#23).

- A `probes` suite for the first Excel run, `tests/modules/TEST_Probes.bas`:
  7 cases and 49 assertions that measure what the workbook relies on Excel
  to do. An error value typed in an input cell reads as `#ERROR` where an
  output shows that text, and no output cell holds an error value; a text
  that reads as a number keeps its digits when it is written through the
  cell's format; a range of several cells is a table; the status line shows
  on the other output sheets after a run; and the status check leaves the
  save state as the guides say. It is listed in the evidence policy, so an
  Excel run must hold its logs, and the `Host behaviour` table of the
  validation plan has an entry, `HB-08` to `HB-16`, for each behaviour,
  all `assumed`. `INSTALLATION.md` has the host checks that only a person
  can make: the two workbook events, the open and close without a change,
  the round trip and the protection of the sheets. The suite passes in
  LibreOffice except for eleven assertions, on error values and on
  protection, which LibreOffice cannot pass; nothing was
  run in Excel (#22). `tools/devbuild/lo_smoke.py` judges the RESULT line of
  every run on its own, so that a known failure never hides a host state
  that was not restored, a recovery that is required or a run that was
  refused.

- The registered test cases run in the workbook. `tests/modules/TEST_Cases.bas`
  is the `cases` suite of the test harness: 18 cases and 345 assertions, one
  case per test case and one assertion per expected output, plus the
  restoration of the sheets. Each case types its inputs on the input sheets,
  runs `RunSSA_Silent`, compares every expected output within its tolerance,
  records every number as a result, and puts the inputs back, formulas
  included.
  `tools/prototype/case_suite.py` writes the module from `tests/fixtures/`
  and `tests/expected/`, and a test fails while the tracked module is not
  what it writes. Its certification entry point is `CertifyCases`, and the
  evidence policy lists it with its results file, `cases-results.json`. The
  project has its first accuracy contract, `ACC-02`, a draft on the overall
  net open position and the requirement of the FX calculation. The suite
  passes in LibreOffice, where the contract holds; it has not run in Excel
  (#21).

- The workbook's VBA tests run under the test harness, as two suites. The
  `core` suite, 11 cases and 90 assertions in `tests/modules/TEST_Suite.bas`,
  calls the engine's procedures and the worksheet functions and touches no
  sheet. The `workbook` suite, 3 cases and 34 assertions in the new
  `tests/modules/TEST_Workbook.bas`, runs the whole calculation on the
  sample data and leaves the output sheets cleared. Each has an entry point
  to type, `TEST_Suite.RunCoreTests` and `TEST_Workbook.RunWorkbookTests`,
  and a certification entry point, `CertifyCore` and `CertifyWorkbook`, and
  the evidence policy lists both, so an Excel run must now hold their logs.
  The 124 assertions are those of the earlier module, with amounts now
  compared within the default tolerances and counts exactly.
  `tools/devbuild/lo_smoke.py` also certifies each project suite, whose run
  must pass. Both suites pass in LibreOffice; neither has run in Excel
  (#21).

- The project's test cases: 18 JSON fixtures with their expected results,
  on the seven calculations. One holds the two values printed in the
  example of table 9 of MAR40.61, as `published`; the others hold the
  template's sample portfolio, the limits of each calculation and a row for
  each input rule a JSON value can break, as `illustrative` values that
  `tests/derivations/ssa.py` computes from the calculation contracts.
  `tools/prototype/run_cases.py` runs the cases on the engine in
  LibreOffice, where the workbook holds all 327 expected values, and
  `tools/prototype/verify.py` runs them as its last step. The default
  tolerances are decided, and each rule of the traceability table names its
  cases. No
  rule is validated: no case has run in Excel, and only the FX calculation
  has a published value. To let a fixture list the FX rates, `DS-05` and
  `DS-10` have an identifier field (#20).
- The data contract defines every input of the workbook: nine datasets,
  `DS-02` to `DS-10`, one for each table that is typed on a sheet, with the
  type, unit, allowed values and column of every field; how a cell is read;
  the rules of `Checks` that reject or report input; the largest supported
  volume, 500 rows on a position sheet; and what a run records of its
  inputs. It lists three points where the engine does not yet do what the
  contract says, each with its issue. The calculation contracts name their
  inputs as datasets and fields. `tools/check_data_contract.py` lets a sheet
  hold several datasets, as `Params` does; a file stays one dataset (#19).
- The calculation has a registered regulatory basis. The Basel Framework's
  chapter MAR40, in force from 1 January 2023 and last updated on 5 July
  2024, is `SRC-02` of the source register, with its terms and what was
  compared with its text; no copy of it is kept. Seven calculation
  contracts, `CALC-02` to `CALC-08`, state the rules of the engine, each
  with its paragraph: interest rate specific risk, the maturity method,
  equity, commodities, foreign exchange, the option charges and the
  requirement. Every parameter is traced to its paragraph, and its
  reference value equals the text. The support matrix states what is
  supported, what the user enters and what is not supported, and the
  traceability table names the procedure of each rule. No rule is
  validated, and what the source leaves open is recorded as an assumption
  (#18).
- A prototype of the calculation: the VBA source of a MAR40 engine for
  interest rate, equity, commodity, and FX and gold risk; the command macros
  `RunSSA` and `ClearOutputs`; four worksheet functions; a test module; and
  a macro-free template of 18 sheets with synthetic sample data.
  `tools/devbuild/build_workbook.py` builds the workbook from them and
  stamps it with its commit. The prototype has not been compiled or run in
  Excel and is not validated: no evidence of an Excel run exists (#16).
- A project website, `index.html`, for GitHub Pages at
  <https://danielep71.github.io/K-FRTB-SSA/>: what the workbook calculates,
  how to build it and what has and has not been checked, with the same
  status as the README. `docs/WEBSITE.md` says how it is published and
  maintained (#30).
- The tools that write the template and check the engine, under
  `tools/prototype/`: `build_prototype.py` writes
  `src/workbook/Template.xlsx`, or the same workbook with every output as a
  formula; `ssa_reference.py` holds the sample data and a Python
  re-calculation of them; `compare_outputs.py` and `verify.py` compare the
  engine's output with the formula workbook in LibreOffice, through the
  workbook that `tools/devbuild/` builds. The tracked template is now what
  the tool writes, and a test fails when it is not. In LibreOffice 24.2 the
  engine and the formula workbook agree on every output cell of the sample
  data and of a scenario of invalid rows; this is not Excel evidence (#17).

### Changed

- `python tools/devbuild/excel_run.py` runs unattended: in the Excel instance
  it starts, and only there, it sets `AutomationSecurity` to
  `msoAutomationSecurityLow` before it opens the workbook, so its macros run
  without **Enable Content** where a policy makes programs follow the macro
  setting (`HB-02`). Excel does not save that property, and the command writes
  no registry value or Trust Center setting. The host record keeps
  `automation_security` as found and adds `automation_security_run`, the value
  of the run; the wait for **Enable Content** remains for a host where the
  macros are still disabled. `HB-18` records the behaviour, measured in Excel
  16.0 build 20527: a smoke run passed in 20 s with nobody touching Excel, and
  a new Excel kept `AutomationSecurity` as before
  ([#137](https://github.com/danielep71/K-FRTB-SSA/issues/137)).
- CodeQL now also analyzes every push to a branch other than Dependabot's, so the head of each pull request carries the CodeQL check that OpenSSF Scorecard counts for its SAST check. Before, only commits that a push to `main` or a release branch had analyzed carried it: 4 of the last 30 commits, although every pull request was analyzed ([#68](https://github.com/danielep71/K-FRTB-SSA/issues/68)). Pull requests are still analyzed without write permission. The Scorecard score changes only once this reaches `main`.
- The documents describe the two regimes. `README.md` is rewritten: what the
  approach is, the capabilities, the quick start, how a run works, the
  differences between the Basel and the `CRR` regime, the evidence classes,
  what has been checked and the known limitations. `docs/PROTOTYPE.md`, the
  website (`index.html`), the support matrix of the methodology guide and the
  `README` sheet of the template no longer say that the CRR is not modelled:
  they say what the `CRR` regime calculates, which rules it does not support,
  and that every value of the regime is `illustrative`, not validated, and not
  run in Excel, so a result is neither for regulatory reporting nor a statement
  of compliance with the CRR. The social preview is tagged *MAR40 · CRR*, the
  counts of test cases and expected values are those of the repository (44
  cases, 823 expected values in the 43 cases with a JSON fixture, 16 of them
  published), and `CRR_COMPARISON.md` says that its rules are implemented
  since it was written. No calculation and no expected value changes, and
  nothing was run in Excel
  ([#87](https://github.com/danielep71/K-FRTB-SSA/issues/87),
  [#117](https://github.com/danielep71/K-FRTB-SSA/issues/117)).
- The headline of `Results` shows the regime of the run, `Basel` or `CRR`,
  on a fifth row below the reporting date, in the formula workbook and in the
  template. It is a formula that repeats the `Regime` parameter, written as
  `CRR` or `Basel` whatever its case, so a figure no longer needs the
  transitional multiplier row to say which rules produced it. The input
  fingerprint moves from `E18` to `E19`, and the output area of `Results` in
  the protection matrix grows by one row. The template is built again; no
  calculation and no expected value changes, and nothing was run in Excel
  ([#87](https://github.com/danielep71/K-FRTB-SSA/issues/87)).

- `python tools/run_tests.py` runs the repository tests one module per
  process, as many at once as the machine has CPUs, and fails when any module
  fails or runs no test. `check_repository.py` and the `Minimum Python (3.11)`
  job call it instead of `python -m unittest discover -s tests`, which still
  runs the same tests one module after the other. `check_excel_run.py` reads
  each exported module and suite once per process instead of on every
  comparison, and reads the files of a commit with two Git processes instead
  of one per file; the development build compresses each module source once
  per process. The tests themselves are unchanged
  ([#107](https://github.com/danielep71/K-FRTB-SSA/issues/107)).
- The source register completes the verification of the CRR texts. It records
  the terms under which EUR-Lex allows its content to be reused, and keeps no
  local copy of `SRC-04` to `SRC-10` all the same; the Parliament's record
  that it did not object to Delegated Regulation (EU) 2026/1221 (`SRC-06`),
  which is in force; that Delegated Regulation (EU) No 525/2014 (`SRC-07`) is
  not amended; that Delegated Regulation (EU) 2016/861 replaced Article 2 of
  Delegated Regulation (EU) No 528/2014 (`SRC-08`), which a corrigendum of
  2014 also corrects, giving the gamma impact of its Annex I as
  ½ × Gamma × VU²; and that the corrigendum of 2025 to Regulation (EU)
  2024/1623 changes nothing in Articles 325 to 361. The contracts already
  state the replaced Article 2; no rule and no expected value
  changes ([#91](https://github.com/danielep71/K-FRTB-SSA/issues/91)).

- The VBA house style has a new layout, described under *Layout of a component*
  in `docs/VBA_HOUSE_STYLE.md`: a nine-field header on every component and
  procedure, section banners in each procedure body, an intent comment before
  each statement group, and one blank line between procedures. Every VBA
  component of `src/` and `tests/modules/` is converted, the generator of
  `TEST_Cases` writes it in the new layout, and `tools/format_vba.py` checks
  that layout only. No behaviour
  changes ([#101](https://github.com/danielep71/K-FRTB-SSA/issues/101)).

- The security policy links to the repository's private vulnerability
  reporting channel (#70).

- A run stops when a position sheet holds more rows than the workbook
  supports. The limit is `MAX_POSITION_ROWS`, 500, the rows the template is
  formatted, validated and coloured for, counted from the row below the
  header to the last row with a PositionID. On more, the status line reads
  `STOPPED`, no result is calculated, and `Checks` lists the sheet and the
  rows found under rule `C14`, which already stops the run on an unusable
  `Params` value; a table of exactly 500 rows runs as before. Before, such a
  run calculated every row and wrote results that no format, formula or check
  covered. The stop text of the status line reads `unusable input(s)`. A run
  of four full sheets took about 12 seconds in LibreOffice, and none has been
  timed in Excel. The data contract no longer lists the limit as a gap (#45).
- The engine rejects two inputs that the data contract excludes. An interest
  rate row whose instrument is not one of `Bond`, `FRN`, `Swap leg`,
  `Future leg`, `FRA leg`, `Forward leg` and `Other`, or whose instrument is
  blank, is `INVALID: Instrument must be ...` and is not calculated; before,
  the column was not read. A commodity row whose commodity is the gold code
  is `INVALID: gold is FX risk; enter it on FXPositions`; before, it was
  calculated as a commodity, which MAR40.63 excludes. `Other` stays among the
  instruments. The formula workbook applies the same two tests, the data
  contract no longer lists them as gaps, and `illustrative-invalid-ir-positions`
  and `illustrative-invalid-co-positions` hold a row for each. The list
  cannot tell a supported instrument from one entered under its name, which
  stays a stated limit. The figures of the sample data do not change (#44).
- The social preview says *Prototype* and that the workbook has not been run
  in Excel yet, where it said *In development*; its only figures are still
  the four scaling factors of MAR40.2. The sentence of the structure guide
  that called the public API inventory empty now describes its two tiers
  (#23).

- `TEST_Suite.Test_RunAll` no longer exists: run the two suites of the VBA
  tests by their entry points. `TEST_Suite.Test_RunCore` now returns the
  `RESULT` line of the harness, as `TEST_Workbook.Test_RunWorkbook` does,
  and `tools/prototype/verify.py` runs both as its seventh step.
  `tools/devbuild/lo_smoke.py` opens the workbook with a view, which the
  workbook's entry points need, and reports the error of each certification
  under `errors` (#21).
- The workbook no longer shows results it cannot stand behind (#25):
  - A run that raises an error, while it reads, calculates or writes, clears
    every output area and writes a status line that begins with `FAILED`,
    with the time and the error. Before, an error while writing left some
    tables under a status of `OK`, and an error before writing left the
    previous results unmarked.
  - The status line is written after every table, and ends with the time of
    the run.
  - `RunSSA` stores a fingerprint of the inputs it read in
    `Results_Fingerprint`. The new macro `CheckStatus`, which `ThisWorkbook`
    also runs when the workbook is opened and before it is saved, compares
    it with the inputs on the sheets and writes `OUTDATED:` before the
    status line when an input changed after the run.
  - `IR_Zones`, `IR_Ladder`, `Buckets`, the four calculation sheets and
    `Checks` repeat the status line of `Results` in `E1`.

  The template gains the fingerprint cell and the repeated status lines.
  The figures of a run that ends do not change. Tested in LibreOffice's VBA
  mode only, not in Excel.
- The template's README sheet no longer names another repository: its
  subtitle, its *Names* row and its *References* section describe this
  workbook only, and its *Source* row names this repository (#17).
- **Repository integrity** runs on pushes to the default branch and to
  release branches only, and on every pull request as before. A push to a
  task branch no longer starts a second run of the commit its pull request
  already checks; run a branch without a pull request by hand. The check
  `codeql-default-branch` is now `workflow-default-branch` and also fails
  while this workflow's filter omits the default branch (#32).
- The workbook contract describes the prototype's sheets, defined names,
  areas and operations instead of illustrative ones. It also records where
  the prototype does not yet follow the contract's defaults: a failed or
  outdated run is not marked on the sheets (#25), and no sheet is protected
  (#26).
- The README describes this repository: a workbook for the FRTB simplified
  standardised approach, its status, with no calculation and no release yet,
  and what is in place. It no longer holds the steps for setting up a new
  repository or the banner image (#3).
- The social preview image is this repository's own: the project name and
  scope, the four scaling factors of the simplified standardised approach
  and the build chain, with no result figure. The README shows it again,
  and `docs/assets/README.md` says how to render it from
  `docs/assets/social-preview.html` (#6).

### Fixed

- `README.md` has *Status* as a second-level heading again, `## 📌 Status`,
  which the rewrite of the README had made first level: a tool that reads the
  section whose second-level heading ends with "Status" found none. The text
  of the section and its anchor `#-status` do not change
  ([#135](https://github.com/danielep71/K-FRTB-SSA/issues/135)).
- A date typed as a number, such as `20311231`, in a cell formatted as a
  date no longer stops the run in Excel with the error 6 (Overflow): `M_Main`
  reads every input with `Value2`, which gives a date as its serial number, so
  the engine marks the row `INVALID` as the contract states. The engine already
  took a date as a number, and the results on valid inputs do not change. A new
  case of the `probes` suite, `date-typed-as-a-number`, measures that read, and
  the *Host behaviour* table of the validation plan records it as `HB-17`
  ([#127](https://github.com/danielep71/K-FRTB-SSA/issues/127)).
- The case `sheet-protection` of the `probes` suite protects `Results` with
  `UserInterfaceOnly:=False` stated, which is what the write it expects to be
  refused needs: measured by hand in Excel 16.0 build 20527, `Unprotect` does
  not take `UserInterfaceOnly` off a sheet and a `Protect` that omits the
  argument keeps what an earlier `AllowWriting` applied, so the first Excel
  run saw the write accepted. The `probes` suite now has 8 cases and 54
  assertions. The `cases` suite stopped at its 20th case in that run; the fix
  for it was not run in Excel
  ([#126](https://github.com/danielep71/K-FRTB-SSA/issues/126)).
- The formula workbook that `tools/prototype/build_prototype.py` writes reads
  a currency code of `FX_Table`, a direction and a rate type without the
  blanks at their ends and without regard to case, as the engine does. Before,
  a code such as `usd ` in `FX_Table` made 282 of 5,447 output cells differ,
  and a direction such as ` long ` made 99 differ, so the comparison reported
  differences that were not the engine's. The currencies shown on `IR_Ladder`,
  `IR_Zones` and `Buckets` are the trimmed, upper-case codes the engine
  writes. `tools/prototype/verify.py` types such codes in its scenario of
  invalid rows. The engine, the template and the sample data's result do not
  change (#41).
- The template no longer calls the re-calculation of the sample data
  independent. The `TestCatalogue` sheet and the sheet list on `Home` name it
  as the project's own re-calculation, whose values are illustrative, as the
  validation plan states (#54).

- The skeleton of a case of an end-to-end case runner, in the harness guide,
  passes the VBA checks: the handler records the case's error and resumes to
  the cleanup. The earlier skeleton raised the error again after
  `Resume CleanUp`, which `tools/check_vba_jumps.py` reports (#51).
- `tools/devbuild/lo_smoke.py` finds each file a suite saved by its whole
  name. It took the first file whose name ended with the one it looked for,
  so `results.json` could be read from another suite's results file (#21).
- Each calculation contract names its test cases and says which values are
  published. Seven contracts still said that no test case existed, after the
  cases were registered (#49).
- The formula workbook that `tools/prototype/build_prototype.py` writes
  groups, counts and looks up rows by typed keys as text, as the engine
  does. Before, a `*`, `?` or `~` in an Issue ID, a stock, a market, a
  commodity or a currency was read as a pattern, a leading `<`, `>` or `=` as
  a comparison, and an Issue ID entered as `POS:` and a PositionID netted
  with the position of that PositionID, so the comparison reported
  differences that were not the engine's. It also reads an error value in a
  typed text as the text `#ERROR`, as the engine does in Excel, so such a
  value stays on its row. `tools/prototype/verify.py` now includes such keys
  in its scenario of invalid rows, and checks error values on the formula
  workbook alone: LibreOffice gives the engine 0 for a cell that holds one.
  The engine, the template and the sample data's result do not change (#39).
- The workbook built from source runs in LibreOffice again. The engine's
  three error numbers are literals equal to `vbObjectError + 2049` to
  `+ 2051`: LibreOffice cannot evaluate `vbObjectError` in a constant
  expression and then compiled no module, so `tools/devbuild/lo_smoke.py`
  wrote no log. `tools/check_vba_source.py` now rejects such a constant, and
  `tools/devbuild/lo_macro.py` runs one macro of a built workbook in
  LibreOffice and saves the values. LibreOffice is not a supported host, and
  none of this is Excel evidence (#36).
- A final maturity or next repricing date typed as a number that is not a
  date, such as `20311231`, now rejects its interest rate row with a reason,
  and the run goes on. Before, the row was accepted: LibreOffice slotted it
  in the last time band with a status of `OK`, and Excel is expected to stop
  the whole run with an overflow. A reporting date that is not a date is an
  unusable parameter. A date is a serial from 1 January 1900 to 31 December
  9999; no shorter limit applies to a maturity. Tested in LibreOffice's VBA
  mode only, not in Excel (#29).
- Interest rate specific risk no longer nets a position without an Issue ID
  against an issue whose Issue ID equals that position's PositionID. Such a
  position is an issue of its own, and its *Issue key* on `IRCalc` reads
  `POS:` and its PositionID. Before, two such rows of opposite sign had a
  charge of zero with no warning. The sample data's result does not change.
  Tested in LibreOffice's VBA mode only, not in Excel (#28).
- `INSTALLATION.md` no longer lists a tested environment for a commit that is
  not in this repository's history, and the *Host behaviour* table of the
  validation plan no longer cites issues that do not exist here. The three
  measured behaviours keep their probe, host, date and result (#5).
- The guides describe this repository instead of a template: the structure
  guide states what exists here, with `VERSION` not existing yet; the house
  style names the `FRTBSSA` prefix; the label and traffic guides and the
  Dependabot header no longer give steps for a newly created repository; and
  the setup guide says that its setup is done (#9).
- The pull request guide no longer says that Codex needs an environment to
  review: it reviews this repository's pull requests without one (#10).

[Unreleased]: https://github.com/danielep71/K-FRTB-SSA/commits/main
