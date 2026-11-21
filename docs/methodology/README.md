# 📐 Methodology

These guides state, before and while the workbook is built, what each
calculation relies on, accepts, produces and how it is validated. Code review
and green checks show that the repository is consistent; only these guides,
and the evidence they require, show that the figures are right.

| Guide | States |
| --- | --- |
| [📚 Sources](SOURCES.md) | Every regulatory and technical source a rule relies on, the version used, and whether it has been verified. |
| [🧾 Data contract](DATA_CONTRACT.md) | The input datasets, their layout, units and validation, and how an import is traced. |
| [🧮 Calculation contracts](CALCULATION_CONTRACTS.md) | For each calculation: inputs, rules with their sources, outputs, limits and failure behaviour. |
| [✅ Validation plan](VALIDATION_PLAN.md) | Who checks what, the reference values and cross-checks for each calculation, how each rule traces to its procedure, cases and status, what an evidence record contains, and which Excel host behaviour has been measured. |
| [🧪 Numerical test cases](TEST_CASES.md) | The format of the synthetic inputs and expected results, with their provenance and tolerances, that the checks enforce. |
| [📊 Workbook contract](WORKBOOK_CONTRACT.md) | The workbook's sheets, code names, defined names and areas, what each operation writes, and how a failure marks earlier results. |
| [🔍 CRR comparison](CRR_COMPARISON.md) | For the CRR regime, each rule of the calculation contracts compared with Article 325(2) of the CRR and the articles it points to: what is the same, what differs, what the CRR adds, and the task that takes each. |
| [🧪 VBA test harness](TEST_HARNESS.md) | The harness that runs the VBA suites in Excel, its outcomes, the log and certification formats, and the evidence an Excel run leaves. |

Read the first five in that order: a rule cites a source, consumes the data
contract's fields, and is validated as the plan describes, with test cases in
the format the fifth guide defines. The workbook contract then places those
inputs and results in the workbook, and the harness guide says how the VBA
suites run in Excel and what evidence a run leaves.

## Support matrix

The workbook's scope, feature by feature:

| Feature | Status | Calculation or input | Notes |
| --- | --- | --- | --- |
| Interest rate risk: specific risk of debt positions, by table 1 of MAR40.6 | Supported | `CALC-02` | Netted per issue. Implementation and validation status per rule: [traceability](VALIDATION_PLAN.md#traceability) |
| Interest rate risk: general market risk by the maturity method | Supported | `CALC-03` | One ladder per currency |
| Equity risk: specific risk, the index charge and general market risk | Supported | `CALC-04` | Per stock or index, and per national market |
| Commodities risk by the simplified approach | Supported | `CALC-05` | Per commodity |
| Foreign exchange risk, gold included, by the shorthand method | Supported | `CALC-06` | From the net position entered per currency |
| The requirement with its scaling factors, and the risk-weighted assets | Supported | `CALC-08` | |
| Under the `CRR` regime, the size conditions of the simplified standardised approach (Article 325a) and of the small trading book derogation (Article 94) | Supported | `CALC-09`; `DS-11` | From the monthly totals entered; a finding that changes no requirement |
| Non-delta charges for options: the simplified approach, the gamma and vega charges of the delta-plus method, the scenario approach | Entered by the user | `CALC-07`; `DS-10.charge` | One value per risk class, added before scaling |
| Conversion of derivatives into positions in the underlying, and the delta-equivalent of an option | Entered by the user | The rows of `DS-06` to `DS-09` | The workbook converts nothing: each leg is a row |
| The category and the rating grade of an issuer | Entered by the user | `DS-06.sr_category` and `DS-06.rating_grade` | The workbook does not apply the qualifying criteria |
| The net open position in each currency, from its components | Entered by the user | `DS-09.net_position` | The workbook derives no FX position from the other sheets |
| Interest rate risk by the duration method | Not supported | `CALC-03` | No input selects it: the maturity method always applies |
| Securitisation positions, credit derivatives and the correlation trading portfolio | Not supported | `CALC-02` | No input identifies one: entered as a debt position, it is calculated as one. Do not enter it |
| Commodities risk by the maturity ladder approach | Not supported | `CALC-05` | The maturity date of a commodity position is not used |
| National discretions, the equity arbitrage treatments, the offset of depository receipts and the exemption for insignificant FX business | Not supported | `CALC-02` to `CALC-06` | Each contract's *Limits and non-goals* names them |
| A jurisdiction's own implementation, such as the European Union's | Not supported | [Sources](SOURCES.md#decisions) | The basis is the Basel Framework's chapter MAR40 |

- **Supported:** in scope. The workbook calculates it, as the named
  [calculation contract](CALCULATION_CONTRACTS.md) states, once that
  contract's rules are implemented; *Supported* does not say that they are.
  The [traceability table](VALIDATION_PLAN.md#traceability) gives each rule's
  status, from *Not implemented* to *Validated*.
- **Not supported:** out of scope. The contract's *Limits and non-goals* say
  what happens when the input holds it: it is reported as unsupported scope,
  never calculated silently
  ([input and output contracts](../REPOSITORY_STRUCTURE.md#input-and-output-contracts)).
  The workbook cannot meet this for an instrument entered as a supported
  one, as the rows say: it is a stated limit of the
  [data contract](DATA_CONTRACT.md#unsupported-input). What a field can
  exclude, the engine rejects: an instrument outside the list of `DS-06`,
  and a commodity that is gold.
- **Entered by the user:** the workbook does not calculate it: the user
  enters the value, which the [data contract](DATA_CONTRACT.md) or the
  [workbook contract](WORKBOOK_CONTRACT.md) defines and validates, and the
  workbook uses it as entered. Name that input in the third column.

## Filling in the guides

The tables of project entries start with entries marked *Illustrative*:
sources, datasets and their fields, validation rules, calculations, the
support matrix, the validation matrix and traceability rows, sheets, defined
names, areas and operations. Replace those entries with the project's own and
keep the columns. The data contract's datasets and fields are edited in its
[machine-readable contract](DATA_CONTRACT.md#machine-readable-contract), which
renders their tables. Tables that define a format or a status, such as the
traceability statuses, describe the method and have no illustrative entry.

Record each decision, with the issue where it was taken, in the *Decisions*
table of the guide it concerns. [Sources](SOURCES.md), the
[data contract](DATA_CONTRACT.md), the
[calculation contracts](CALCULATION_CONTRACTS.md) and the
[workbook contract](WORKBOOK_CONTRACT.md) have one; while it holds only its
illustrative row, nothing has been decided in that guide. The numerical test
cases and the VBA test harness define formats and tools, and have none. Two
tables of the [validation plan](VALIDATION_PLAN.md) differ:

- **Open decisions** lists the questions the plan leaves open, with their
  options. Replace its illustrative row with the project's questions, and
  remove each once the plan states the choice.
- **Host behaviour** has no illustrative entry: its entries record Excel
  behaviour that the repository's tools rely on. Keep them, and add the
  behaviour that the project's code relies on.

Byte-preserved copies of sources go under `docs/methodology/sources/`, whose
Git and editor rules are already set; [Sources](SOURCES.md) describes the
naming and the register entry, with its SHA-256, that each copy needs.
