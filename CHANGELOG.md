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
