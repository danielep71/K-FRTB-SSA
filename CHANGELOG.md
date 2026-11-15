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

- The formula workbook that `tools/prototype/build_prototype.py` writes
  groups, counts and looks up rows by typed keys as text, as the engine
  does. Before, a `*`, `?` or `~` in an Issue ID, a stock, a market, a
  commodity or a currency was read as a pattern, a leading `<`, `>` or `=` as
  a comparison, and an Issue ID entered as `POS:` and a PositionID netted
  with the position of that PositionID, so the comparison reported
  differences that were not the engine's. `tools/prototype/verify.py` now
  includes such keys in its scenario of invalid rows. The engine, the
  template and the sample data's result do not change (#39).
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
