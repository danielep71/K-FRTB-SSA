# 🔒 Security policy

## 📌 Supported versions

The latest release, the active release branch and the default branch are
supported on a best-effort basis. The latest release is the newest version
published as a GitHub Release, as [RELEASING.md](RELEASING.md) describes;
until one exists, only the two branches are supported. Reports must identify
a full commit SHA. Modified copies are unsupported unless the problem
reproduces in this repository.

## 📨 Reporting a vulnerability

Do **not** disclose a suspected vulnerability in an issue, pull request,
commit message, sample file or screenshot.

Report it privately through **Security → Report a vulnerability** on this
repository's GitHub page. The report stays visible only to you and the people
who manage the repository's security until an advisory is published.

If the page has no **Report a vulnerability** button, for example because the
repository is private, open an issue with the **🔒 Private security contact**
form. It asks only how to reach you; I reply with a private channel. Never put
vulnerability details in the issue.

Include:

- the full commit SHA and the affected file, procedure or workflow;
- the Excel/Office build, bitness and Windows version, where relevant;
- the impact and minimal reproduction steps, with synthetic data only; and
- any workaround you tested.

## ⏱️ Response targets

I maintain this repository alone; the targets are best effort, not a
contractual commitment:

- acknowledgement within 5 business days;
- an initial assessment within 10 business days of receiving enough evidence;
- an update at least every 14 days while the report is investigated.

Remediation and disclosure timing are agreed with the reporter according to
severity.

## 🎯 Scope

**In scope:** this repository's VBA source, workbook templates, tests and
tooling, and its GitHub Actions workflows and their configuration.

**Out of scope:**

- vulnerabilities in Excel, Office, Windows, GitHub, Python or Node.js
  themselves;
- organization-controlled macro and deployment policy;
- modified copies;
- ordinary defects without concrete security impact, such as an incorrect
  calculated figure. Report those as issues.

## 🤝 Safe harbor

I welcome good-faith security research within the scope above. If you report
privately as this policy describes, use only synthetic data and your own
copies and accounts, do not access, change or delete anyone else's data, do
not degrade this repository or the services it runs on, and give me
reasonable time to respond before any disclosure, I will not pursue or
support legal action against you for that research. This covers only claims
I could bring; it cannot authorize testing of GitHub, Microsoft or other
third-party systems, whose own policies apply.

## ⚙️ Runtime risk surface

A workbook's macros run with the rights of the person who opens it. This
section states what the code in this repository may do when it runs, so that
you can tell intended behaviour from a vulnerability.

| Component | Reads | Writes or deletes |
| --- | --- | --- |
| Test harness and example suite (`tests/modules/`), in the development workbook only | The illustrative test-case files under the fixtures folder it is given; Excel's `Calculation`, `DisplayAlerts`, `EnableEvents` and `ScreenUpdating`, the number of open workbooks, the active window's caption, the host's name, version, build and operating system, and the `build_commit` defined name of its own workbook | Its logs, `results.json` and the certification log in the evidence folder it is given, replacing a file of the same name; the example suite changes `Calculation` and `ScreenUpdating` and restores them. It deletes no file |
| The project's suites (`TEST_Suite`, `TEST_Workbook`, `TEST_Probes` and `TEST_Cases` in `tests/modules/`), in the development workbook only | The sheets and defined names of their own workbook, and the settings the harness compares; none reads a file | Their log, certification log and, for `TEST_Cases`, results file in the evidence folder they are given. `TEST_Workbook`, `TEST_Probes` and `TEST_Cases` run `RunSSA_Silent`, which rewrites the output sheets, and leave them cleared. `TEST_Cases` replaces the reporting date and currency, the FX table, the option charges and the four position sheets with those of each test case, and puts them back, values and formulas, changing the number format of a cell for the write of a text and restoring it. `TEST_Probes` appends rows under the position tables and removes them, changes `ScalingIR` and `ThisWorkbook.Saved` for one check and puts both back. None saves the workbook, deletes a file, starts a process, creates an external object or reaches the network. An error stops a case, not the run, and `TEST_Cases` puts the inputs back on that path too; a run stopped with **End** or **Reset** leaves a test case's inputs on the sheets, so build the workbook again |
| `tools/prototype/verify.py`, `run_cases.py` and `tools/devbuild/lo_macro.py`, optional repository tooling with LibreOffice | The checkout: the template, the test cases and the tools they call; the workbooks they build in a temporary folder | Work in a new temporary folder, or one you give with `--keep`: build the formula workbook and the development workbook there, start headless LibreOffice with a user profile in that folder, run a macro of the workbook and save the recalculated sheets as a macro-free `.xlsx`; compare the saved values with the expected values; delete the temporary folder when they end. They never save the workbook you give them, and call the same LibreOffice as `lo_smoke.py` |
| `tools/devbuild/excel_run.py`, repository tooling on Windows | The checkout; the evidence folder and the workbook it built, which its check reads; the Windows version from the registry; the `AccessVBOM` value, for the Trust access to the VBA project, under `Software\Policies\Microsoft\Office\<version>\Excel\Security` in HKCU and in HKLM and under `Software\Microsoft\Office\<version>\Excel\Security` in HKCU; Excel's `Version`, `Build`, `Path`, `AutomationSecurity` and `Hwnd` through COM, read only; the PE header of `EXCEL.EXE` in that path, for its bitness; the process id of its own Excel instance, from its `Hwnd` window, and that process's name and start time; and your Windows user name when `--operator` is not given | Builds the development workbook under `test-results/build/`; empties the evidence folder, but only one under `test-results/`, and otherwise requires an empty or new folder, and writes `host.txt` in it and its report beside it as `<folder>-report.txt`; starts its own Excel instance, sets its `Visible` to true and its `DisplayAlerts` to false, opens the workbook read-only without updating links, closes it without saving and quits that instance; if that instance is still running 30 seconds after it quits, ends its process, and only that one: it ends no process whose name is not `EXCEL` or whose start time is not the one it recorded when it started the instance; for its check, creates a temporary folder, adds there a temporary Git worktree of the commit, which it registers in the checkout's repository, builds a second workbook from it in that folder, and removes the worktree and the folder when the check ends; to measure the accuracy contracts, the check also writes the commit's test cases into another temporary folder, which it deletes after the measurement |
| `tools/devbuild/lo_smoke.py`, optional repository tooling with LibreOffice | The checkout: `.github/excel-evidence-policy.json`, the source of the development workbook when you give it no workbook, and the test cases under `tests/fixtures`; the workbook you give it, which must be one you built | Works in a new temporary folder: builds the development workbook there when you give it none; starts headless LibreOffice with a user profile in that folder, not yours, and sets that profile to load and run VBA; opens the workbook in that headless LibreOffice with `MacroExecutionMode` 4, so its macros run without a warning; runs the certification entry point of the example suite and of each project suite, which write their evidence into a folder there. The workbook suite and the cases suite run the calculation on the open copy: they rewrite its output sheets, and the cases suite replaces its inputs and puts them back; closes the workbook without saving it. It deletes the temporary folder, profile and evidence included, when it ends |
| `M_Main.RunSSA` and `RunSSA_Silent`, in the production and the development workbook | Every input area of its own workbook: the reporting date and currency, the MAR40 constants, the specific-risk table, the time bands and the FX table of `Params`, the four position sheets and `OptionCharges`; the defined names of those areas; and Excel's `Calculation`, `ScreenUpdating` and `EnableEvents` | Clears, then writes, every output area: `Results`, `IR_Zones`, `IR_Ladder`, `Buckets`, `IRCalc`, `EQCalc`, `COCalc`, `FXCalc` and `Checks`, with `Results_Fingerprint` and `Results_Status`. It sets calculation to manual, and screen updating and events off, for the run, and puts them back, on the success and the error path. `RunSSA` also shows a message box and activates `Checks` or `Results`; the silent variant does neither, returns a result line and re-raises an error. It saves nothing, deletes no sheet or file, starts no process, creates no external object, calls no native library and reaches no network: the calculation needs none |
| `M_Main.ClearOutputs` and `ClearOutputs_Silent` | The defined names of the output areas, and Excel's settings as above | Clears the values of the same output areas, not their formats, and sets `Results_Status` to *Not calculated*. It goes on after an area it cannot clear and then stops with an error that names it. `ClearOutputs` shows a message box on failure. As `RunSSA` for the settings, and for saving, deleting, processes, objects, libraries and the network |
| `M_Main.CheckStatus`, `CheckStatus_Silent` and `CheckStatus_OnEvent` | Every input area, `Results_Fingerprint` and `Results_Status`, and Excel's settings as above | `Results_Status`, and only when it must change: it writes or takes away `OUTDATED:` before the line of the run. `CheckStatus` shows a message box. `CheckStatus_OnEvent` is what `ThisWorkbook` runs when the workbook is opened and before it is saved: it shows no message, raises no error, saves nothing and cancels no save, and leaves the workbook marked as saved when it was and it wrote nothing. As `RunSSA` for the rest |
| Worksheet functions of `FRTBSSA_Formulas` | Their arguments | Nothing: they read and write no sheet, and return a value or an error value |

- **Excel settings.** The procedures need their workbook's macros to be
  enabled, by your organization's policy or by **Enable Content**, and
  nothing else: no lower macro-security level and no *Trust access to the VBA
  project object model* ([host policy](INSTALLATION.md#%EF%B8%8F-host-policy)).
  `tools/devbuild/excel_run.py` changes no macro setting: where Excel opens
  the workbook with macros disabled, it waits for **Enable Content**. Do not
  change a macro setting to run them.
- **Host-independent code.** The [dependency rules](docs/REPOSITORY_STRUCTURE.md#dependencies)
  forbid `src/core/` and `src/classes/` any Excel object model, file,
  registry, process or network access, external object model, native
  `Declare`, dynamic call, and `CreateObject` or `GetObject` other than
  `CreateObject("Scripting.Dictionary")`. `tools/check_vba_dependencies.py`
  enforces them by name, with the limits the
  [static checks](docs/VBA_HOUSE_STYLE.md#static-checks) list.
- **Other code.** Those rules do not restrict the facade, the workbook and
  form code, or the tests. For each entry point of the production workbook,
  state in the table above which sheets, defined names and files it reads,
  what it writes, saves or deletes, and whether it starts a process
  (`Shell`), creates an external object (`CreateObject`, `GetObject`), calls
  a native library (`Declare`) or reaches the network, and why.

## 🗝️ Data and secrets

- Use synthetic data only, in reports as everywhere else. Never attach a
  workbook or file that holds real portfolio, client or counterparty data.
- A workbook holds more than its visible cells: hidden and very hidden
  sheets, defined names, comments and their authors, links to other workbooks
  and data connections, document properties, cached pivot data and the VBA
  project. Inspect a workbook before you share it, as the
  [distribution review](docs/repo_governance/DISTRIBUTION.md#-workbooks)
  describes.
- Keep secret and key files out of the repository. `.gitignore` ignores them,
  but `git add -f` bypasses it, so **Repository integrity** also fails on a
  tracked file with one of these names, in any directory and in any letter
  case:
  - `.env` and `.env.*`, except `.env.example`;
  - `*.pem`, `*.key`, `*.p12`, `*.pfx`, `*.pvk`, `*.jks` and `*.keystore`;
  - `id_rsa*` and `id_ed25519*`, except their `.pub` public keys;
  - `*.kdbx`.

  The check looks at names only, and secret scanning matches known token
  patterns, so neither proves that a file holds no secret.
- If a secret, such as a token, password or key, has leaked in a file, a
  workbook, the history, an issue or a log, revoke or rotate it first, then
  report it. Deleting it does not undo its exposure
  ([handling a finding](docs/repo_governance/DISTRIBUTION.md#-handling-a-finding)).

## 📦 Before distribution

Before the repository is made public, a copy is shared or a release is
published, I record the
[distribution review](docs/repo_governance/DISTRIBUTION.md): license,
provenance, a scan of the whole history, the workbooks, and the issue and pull
request text.
