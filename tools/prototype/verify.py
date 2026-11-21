#!/usr/bin/env python3
"""Check the VBA engine against the formula workbook in LibreOffice, end to end.

Usage: python3 tools/prototype/verify.py [--uno-python PYTHON] [--keep FOLDER]

It needs LibreOffice and openpyxl. The steps:

1. Build the formula workbook with build_prototype.py, and the development
   workbook with tools/devbuild/build_workbook.py, from the checkout.
2. Recalculate the formula workbook: every test of its TestCatalogue passes.
3. Run RunSSA on the sample data and compare every output area with the
   formula workbook's.
4. Apply the negative scenario to both workbooks, run and compare again, the
   findings on Checks included. The scenario adds rejected rows, an issue
   that mixes specific-risk keys, a position without an Issue ID whose
   PositionID is another row's Issue ID, an Issue ID that reads as the key
   shown for that position, a maturity that is not a date, a changed
   constant and an option charge. It also adds keys that a criteria function
   would read as a pattern or as a comparison: a PositionID, Issue IDs, a
   stock, a market, commodities, and a currency of FX_Table that is a
   pattern for another one, with positions in it. And it adds error values
   where a text is typed and no output shows it: the Issue ID of a rejected
   row, and a direction.
5. Type error values where an output shows the text, on the formula workbook
   only: no output cell holds an error value, and each such text reads
   "#ERROR", as the engine reads it in Excel. LibreOffice gives the engine 0
   for a cell that holds an error value, so the two cannot be compared here.
   The same under the CRR regime for months of Eligibility with error values:
   no output cell holds one, and each such month is rejected as
   ssa.eligibility rejects it.
6. Make two Params values unusable: the run stops and lists both.
7. Run the three suites of the VBA tests, through TEST_Suite.Test_RunCore,
   TEST_Workbook.Test_RunWorkbook and TEST_Cases.Test_RunCases: each run
   passes. The last is the module that case_suite.py writes from the
   registered test cases.
8. Run the registered test cases, as run_cases.py does: the workbook holds
   every expected value.
9. Type series of months on Eligibility under the CRR regime, in both
   workbooks: the sample months, months that all meet the conditions, a
   series that ends in a run of three, one whose last month meets them after
   seven of twelve that do not, and rows of each invalid kind. The engine's
   calculated columns, result block and counts of Checks equal the formula
   workbook's and those of ssa.eligibility in tests/derivations/ssa.py.

Every macro runs through tools/devbuild/lo_macro.py, on the workbook the
build produces. --uno-python names the Python that provides LibreOffice's
UNO bridge, when it is not the one running this script. --keep keeps the
workbooks of the run in FOLDER; without it they go to a temporary folder.

It prints each step's result and VERIFY: PASS, and exits 0; at the first
step that fails it prints VERIFY: FAIL and exits 1. It exits 2 when
LibreOffice is missing. LibreOffice is not a supported host: this compares
two implementations of the same rules, and is never Excel evidence.
"""

from __future__ import annotations

import argparse
import datetime
import shutil
import subprocess
import sys
import tempfile
import unittest.mock
import warnings
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any

import openpyxl

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))
sys.path.insert(0, str(ROOT / "tests" / "derivations"))

import build_prototype  # noqa: E402
import build_workbook  # noqa: E402
import compare_outputs  # noqa: E402
import ssa  # noqa: E402
import ssa_reference  # noqa: E402

LO_MACRO = ROOT / "tools" / "devbuild" / "lo_macro.py"
RUN = "M_Main.RunSSA_Silent"
TESTS = ("TEST_Suite.Test_RunCore", "TEST_Workbook.Test_RunWorkbook",
         "TEST_Cases.Test_RunCases")
# The first row of a position table; its first column is C.
FIRST_ROW = 5
FIRST_COLUMN = 3

# A currency code that a criteria function would read as a pattern for USD.
WILD_CURRENCY = "U?D"
# openpyxl writes a text that is the name of an error value as that error value.
NOT_AVAILABLE = "#N/A"
# The text the engine reads an error value as, in Excel.
ERROR_TEXT = build_prototype.ERROR_TEXT

Findings = set[tuple[object, object, object, object]]


class Failure(Exception):
    """A step did not give what it must; the message says what it gave."""


def load(path: Path) -> Any:
    """The workbook for editing; a macro-enabled one keeps its VBA project."""
    return openpyxl.load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")


def named_cell(workbook: Any, name: str) -> Any:
    """The cell a one-cell defined name refers to."""
    sheet, _, cell = workbook.defined_names[name].attr_text.rpartition("!")
    return workbook[sheet.strip("'")][cell.replace("$", "")]


def first_cell(workbook: Any, name: str) -> Any:
    """The first cell of the range a defined name refers to."""
    sheet, _, cells = workbook.defined_names[name].attr_text.rpartition("!")
    return workbook[sheet.strip("'")][cells.replace("$", "").partition(":")[0]]


def append_rows(sheet: Any, first: int, records: Sequence[Sequence[object]]) -> None:
    for offset, record in enumerate(records):
        for column, value in enumerate(record, FIRST_COLUMN):
            sheet.cell(first + offset, column).value = value


def negative(source: Path, target: Path) -> None:
    """Save ``source`` with the negative scenario on its input sheets."""
    workbook = load(source)
    day = datetime.date
    append_rows(workbook["IRPositions"], FIRST_ROW + len(ssa_reference.IR), [
        # An instrument outside the list of the data contract, and none at all.
        ["IR-INST1", "T", "Securitisation", "EUR", "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "NONE", "", ""],
        ["IR-INST2", "T", None, "EUR", "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "NONE", "", ""],
        # The instrument is read without regard to case and the blanks at its ends.
        ["IR-INST3", "T", " swap LEG ", "EUR", "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "NONE", "", ""],
        # Not the blanks inside it: Excel's TRIM would collapse them, VBA's Trim$ does not.
        ["IR-INST4", "T", "Swap  leg", "EUR", "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "NONE", "", ""],
        # No coupon; a currency outside the FX table; a duplicate PositionID.
        ["IR-BAD1", "T", "Bond", "EUR", "Long", 1e6, "Fixed", None, day(2030, 1, 1), None,
         "OTHER", "BB+ to BB-", ""],
        ["IR-BAD2", "T", "Bond", "ZAR", "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "NONE", "", ""],
        ["ir-001", "T", "Bond", "EUR", "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "NONE", "", ""],
        # An issue of the sample with another specific-risk key: MIXED.
        ["IR-MIX", "T", "Bond", "EUR", "Long", 2e6, "Fixed", 0.04, day(2034, 3, 1), None,
         "QUAL", "A+ to BBB-", "BTP 4.00% 03/34"],
        # A direction and a rate type typed with blanks and in lower case.
        ["IR-FRN", "T", "FRN", "EUR", " short ", 3e6, " floating ", None, day(2029, 1, 1),
         day(2027, 1, 1), "GOV", "AAA to AA-", ""],
        # No specific-risk charge for this category and grade.
        ["IR-OTH", "T", "Bond", "EUR", "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "OTHER", "A+ to BBB-", ""],
        # A maturity typed as a number: not a date.
        ["IR-DATE", "T", "Bond", "EUR", "Long", 1e6, "Fixed", 0.05, 20311231, None,
         "NONE", "", ""],
        # A position without an Issue ID, and another whose Issue ID is that
        # position's PositionID: they do not net.
        ["R1", "T", "Bond", "EUR", "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "OTHER", "BB+ to BB-", ""],
        ["IR-R1", "T", "Bond", "EUR", "Short", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "OTHER", "BB+ to BB-", "R1"],
        # A PositionID that a criteria function would read as a pattern
        # matching every other one: it is unique, so the row is used.
        ["IR-*", "T", "Bond", "EUR", "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "NONE", "", ""],
        # Issue IDs that a criteria function would read as patterns or as a
        # comparison: four issues, none netted with another.
        ["IR-W1", "T", "Bond", "EUR", "Long", 2e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "OTHER", "BB+ to BB-", "ISS-?"],
        ["IR-W2", "T", "Bond", "EUR", "Short", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "OTHER", "BB+ to BB-", "ISS-A"],
        ["IR-W3", "T", "Bond", "EUR", "Short", 5e5, "Fixed", 0.05, day(2030, 1, 1), None,
         "QUAL", "A+ to BBB-", "ISS-*"],
        ["IR-W4", "T", "Bond", "EUR", "Long", 3e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "OTHER", "BB+ to BB-", "<>ISS-A"],
        # An Issue ID that reads as the key shown for the position R1, which
        # has no Issue ID: they do not net.
        ["IR-P1", "T", "Bond", "EUR", "Short", 4e5, "Fixed", 0.05, day(2030, 1, 1), None,
         "OTHER", "BB+ to BB-", "POS:R1"],
        # A currency that is a pattern for one of FX_Table, and one that
        # FX_Table holds as it is typed: the first is rejected, the second
        # has its own ladder.
        ["IR-C1", "T", "Bond", "E?R", "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "NONE", "", ""],
        ["IR-C2", "T", "Bond", WILD_CURRENCY, "Long", 7e6, "Fixed", 0.05, day(2031, 1, 1), None,
         "NONE", "", ""],
        ["IR-C3", "T", "Bond", WILD_CURRENCY, "Short", 2e6, "Fixed", 0.02, day(2031, 1, 1), None,
         "NONE", "", ""],
        # Error values where no output shows them: the Issue ID of a rejected
        # row, and a direction. Neither may reach another row.
        ["IR-E1", "T", "Bond", "EUR", "Long", 1e6, "Fixed", None, day(2030, 1, 1), None,
         "OTHER", "BB+ to BB-", NOT_AVAILABLE],
        ["IR-E2", "T", "Bond", "EUR", "#DIV/0!", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "NONE", "", ""],
    ])
    append_rows(workbook["EQPositions"], FIRST_ROW + len(ssa_reference.EQ), [
        ["EQ-BAD", "T", "IT", "Basket", "B", "EUR", "Long", 1e6],
        ["EQ-NEG", "T", "FR", "Single name", "BNP", "EUR", " Short", 2e6],
        # A stock and a market that would match others as patterns.
        ["EQ-W1", "T", "FR", "Single name", "A*", "EUR", "Long", 1e6],
        ["EQ-W2", "T", "FR", "Single name", "ABC", "EUR", "Short", 4e5],
        ["EQ-W3", "T", "F?", "Single name", "BNP", WILD_CURRENCY, "Long", 3e6],
        ["EQ-W4", "T", "DE", "Single name", "SAP", "US*", "Long", 1e6],
    ])
    append_rows(workbook["COPositions"], FIRST_ROW + len(ssa_reference.CO), [
        ["CO-BAD", "T", "Gold bar", "Long", "ten", "oz", 3500, "EUR"],
        ["CO-007", "T", "Wheat", "short ", 1000, "t", 250, "EUR"],
        # Gold is FX risk: a commodity that is the gold code is rejected, in any case.
        ["CO-GOLD", "T", " xau ", "Long", 1, "oz", 3500, "EUR"],
        # Commodities that would match others as patterns.
        ["CO-W1", "T", "C*", "Long", 10, "t", 9800, "USD"],
        ["CO-W2", "T", "Whea?", "Long", 500, "t", 250, WILD_CURRENCY],
        ["CO-W3", "T", "~Wheat", "Long", 100, "t", 250, "EUR"],
    ])
    append_rows(workbook["FXPositions"], FIRST_ROW + len(ssa_reference.FXPOS), [
        ["FX-007", "FX", "EUR", "Long", 5e6],
        ["FX-008", "FX", "JPY", "Long", -5],
        ["FX-009", "FX", "CHF", "Long", 4e6],
        ["FX-W1", "FX", "U*", "Long", 1e6],
        ["FX-W2", "FX", WILD_CURRENCY, "Short", 6e6],
        ["FX-E1", "FX", "CHF", NOT_AVAILABLE, 1e6],
    ])
    # USD, typed in FX_Table with a blank at each end and in lower case: the
    # positions in USD, and the ladders and sums of the currency, do not change.
    first_cell(workbook, "FX_Ccy").offset(next(
        row for row, (ccy, *_) in enumerate(ssa_reference.FX) if ccy == "USD")).value = " usd "
    # A currency of FX_Table that is a pattern for another one, USD.
    code = first_cell(workbook, "FX_Ccy").offset(len(ssa_reference.FX))
    code.value = WILD_CURRENCY
    first_cell(workbook, "FX_Rate").offset(len(ssa_reference.FX)).value = 0.5
    named_cell(workbook, "ScalingIR").value = 1.5
    named_cell(workbook, "OptionsFX").value = 1000
    workbook.save(target)


def error_keys(source: Path, target: Path) -> None:
    """Save ``source`` with error values typed where an output shows the text."""
    workbook = load(source)
    day = datetime.date
    append_rows(workbook["IRPositions"], FIRST_ROW + len(ssa_reference.IR), [
        # Two Issue IDs that both read "#ERROR": one issue.
        ["IR-K1", "T", "Bond", "EUR", "Long", 3e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "OTHER", "BB+ to BB-", NOT_AVAILABLE],
        ["IR-K2", "T", "Bond", "EUR", "Short", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "OTHER", "BB+ to BB-", "#VALUE!"],
        ["IR-K3", "T", "Bond", NOT_AVAILABLE, "Long", 1e6, "Fixed", 0.05, day(2030, 1, 1), None,
         "NONE", "", ""],
    ])
    append_rows(workbook["EQPositions"], FIRST_ROW + len(ssa_reference.EQ), [
        [NOT_AVAILABLE, "T", "IT", "Single name", "ENI", "EUR", "Short", 1e6],
        ["EQ-K2", "T", "IT", "Single name", "#REF!", "EUR", "Long", 2e6],
        ["EQ-K3", "T", NOT_AVAILABLE, "Index", "DAX", "EUR", "Long", 1e6],
    ])
    append_rows(workbook["COPositions"], FIRST_ROW + len(ssa_reference.CO), [
        ["CO-K1", "T", NOT_AVAILABLE, "Long", 100, "t", 250, "EUR"],
        ["CO-K2", "T", "Wheat", "Long", 100, "t", 250, NOT_AVAILABLE],
    ])
    append_rows(workbook["FXPositions"], FIRST_ROW + len(ssa_reference.FXPOS), [
        ["FX-K1", "FX", NOT_AVAILABLE, "Long", 1e6],
    ])
    workbook.save(target)


def error_months() -> list[Month | None]:
    """Months of Eligibility with an error value in the month end, in the total assets and in a
    long total, each a rejected row, and a row without a month end whose totals are error values,
    which is skipped."""
    day = datetime.date
    return [(day(2025, 10, 31), 8e9, 1e8, 0, 1e7, 0),
            (NOT_AVAILABLE, 8e9, 1e8, 0, 1e7, 0),
            (day(2025, 12, 31), 8e9, 1e8, 0, 1e7, 0),
            (day(2026, 1, 31), NOT_AVAILABLE, 1e8, 0, 1e7, 0),
            (day(2026, 2, 28), 8e9, "#VALUE!", 0, 1e7, 0),
            (None, NOT_AVAILABLE, NOT_AVAILABLE, 0, 1e7, 0),
            (day(2026, 3, 31), 8e9, 1e8, 0, 1e7, 0)]


def error_cells(path: Path) -> list[str]:
    """The cells of the output areas of a workbook of values that hold an error value."""
    workbook = openpyxl.load_workbook(path, data_only=True)
    return [f"{area.sheet}!{cell.coordinate}"
            for area in compare_outputs.areas(workbook, workbook)
            for row in area.rows for column in area.columns
            for cell in [workbook[area.sheet].cell(row, column)] if cell.data_type == "e"]


def unusable(source: Path, target: Path) -> None:
    """Save ``source`` with a reporting date and an option charge that are not usable."""
    workbook = load(source)
    named_cell(workbook, "AsOfDate").value = "not a date"
    named_cell(workbook, "OptionsIR").value = "abc"
    workbook.save(target)


def crr_positions(workbook: Any) -> None:
    """Give each interest rate row of the sample the category CRR and the risk weight that has its charges.

    Under the CRR the category of a row is CRR or NONE; the grade is the risk weight whose rates of
    Article 336 are the rates that the row's Basel category and grade have, so that the specific risk
    of the sample does not change."""
    weights = {(0.0,): "RW 0%", (0.0025, 0.01, 0.016): "RW 20%", (0.08,): "RW 100%", (0.12,): "RW 150%"}
    sheet = workbook["IRPositions"]
    for row in range(FIRST_ROW, FIRST_ROW + len(ssa_reference.IR)):
        category, grade = sheet.cell(row, 13).value, sheet.cell(row, 14).value
        if category in (None, "NONE"):
            continue
        charges = tuple(sorted({rate for (cat, grd, _), rate in ssa_reference.SR.items()
                                if (cat, grd) == (category, grade)}))
        sheet.cell(row, 13).value = "CRR"
        sheet.cell(row, 14).value = weights[charges]


def regime(source: Path, target: Path, name: str, end: object, own_funds: object = None,
           rows: str = "convert") -> None:
    """Save ``source`` with the regime, the last date of the transitional multiplier and the total own funds set.

    ``rows`` says what becomes of the interest rate rows: under the CRR they get the category CRR first
    ("convert"); "keep" leaves them as the Basel regime has them; "mixed" gives the first two rows the
    category CRR with a weight of the table and one that is not in it."""
    workbook = load(source)
    named_cell(workbook, "Regime").value = name
    named_cell(workbook, "TransitionalEndDate").value = end
    named_cell(workbook, "TotalOwnFunds").value = own_funds
    sheet = workbook["IRPositions"]
    if rows == "convert" and str(name).strip().upper() == "CRR":
        crr_positions(workbook)
    elif rows == "mixed":
        if str(name).strip().upper() == "CRR":
            crr_positions(workbook)
        for row, grade in ((FIRST_ROW, "RW 20%"), (FIRST_ROW + 1, "RW 33%")):
            sheet.cell(row, 13).value = "CRR"
            sheet.cell(row, 14).value = grade
    workbook.save(target)


def read_names(path: Path, *names: str) -> list[Any]:
    """The values of one-cell defined names in a workbook of values."""
    workbook = openpyxl.load_workbook(path, data_only=True)
    return [named_cell(workbook, name).value for name in names]


def findings(path: Path) -> Findings:
    """The findings of Checks: sheet, PositionID or code, row and reason."""
    sheet = openpyxl.load_workbook(path, data_only=True)["Checks"]
    # The findings table is the one whose first header cell reads "#".
    header = next(row for row in range(1, sheet.max_row + 1) if sheet.cell(row, 3).value == "#")
    found: Findings = set()
    for row in range(header + 1, sheet.max_row + 1):
        if sheet.cell(row, 4).value not in (None, ""):
            found.add(tuple(sheet.cell(row, column).value for column in (4, 5, 6, 9)))
    return found


def test_summary(path: Path) -> str:
    """The summary line of TestCatalogue, such as '19 of 19 tests PASS'."""
    value = openpyxl.load_workbook(path, data_only=True)["TestCatalogue"]["C4"].value
    return "" if value is None else str(value)


class Run:
    """The workbooks of one verification and the way to run a macro on them."""

    def __init__(self, folder: Path, uno_python: str) -> None:
        self.folder = folder
        self.uno_python = uno_python
        self.formula = folder / "formula.xlsx"
        self.formula_values = folder / "formula-values.xlsx"
        self.workbook = folder / "Workbook-development.xlsm"

    def macro(self, workbook: Path, macro: str, output: Path) -> str:
        """What ``macro`` returned; a Failure when LibreOffice did not run it."""
        result = subprocess.run(
            [self.uno_python, str(LO_MACRO), str(workbook), macro, str(output)],
            capture_output=True, text=True, check=False)
        lines = [line for line in result.stdout.splitlines() if line.startswith("RESULT: ")]
        if result.returncode != 0 or not lines:
            raise Failure((result.stdout + result.stderr).strip() or "lo_macro.py printed nothing")
        return lines[-1][len("RESULT: "):]

    def compare(self, vba_run: Path, formula: Path) -> str:
        compared, differences = compare_outputs.compare_workbooks(vba_run, formula)
        if differences:
            raise Failure("\n".join([f"{len(differences)} difference(s) in {compared} cells",
                                     *differences[:compare_outputs.SHOWN]]))
        return f"{compared} cells compared, no difference"


def build(run: Run) -> str:
    build_prototype.build(run.formula, template=False)
    try:
        built = build_workbook.build(ROOT, "development", run.workbook)
    except (build_workbook.BuildError, ValueError, OSError) as error:
        raise Failure(f"build: {error}") from error
    return f"formula workbook and development workbook of commit {built.commit}"


def formula_tests(run: Run) -> str:
    run.macro(run.formula, "-", run.formula_values)
    summary = test_summary(run.formula_values)
    count = summary.split(" ", 1)[0]
    if not count or not summary.startswith(f"{count} of {count} "):
        raise Failure(f"TestCatalogue of the formula workbook: {summary!r}")
    return summary


def sample(run: Run) -> str:
    output = run.folder / "vba-sample.xlsx"
    result = run.macro(run.workbook, RUN, output)
    if not result.startswith("RESULT=OK;"):
        raise Failure(result)
    return f"{result}\n{run.compare(output, run.formula_values)}"


def negative_scenario(run: Run) -> str:
    formula, workbook = run.folder / "negative-formula.xlsx", run.folder / "negative.xlsm"
    values, output = run.folder / "negative-formula-values.xlsx", run.folder / "vba-negative.xlsx"
    negative(run.formula, formula)
    negative(run.workbook, workbook)
    run.macro(formula, "-", values)
    result = run.macro(workbook, RUN, output)
    if not result.startswith("RESULT=INCOMPLETE;"):
        raise Failure(result)
    compared = run.compare(output, values)
    engine, reference = findings(output), findings(values)
    if engine != reference:
        raise Failure(f"findings differ: {sorted(map(str, engine ^ reference))}")
    return f"{result}\n{compared}\n{len(engine)} findings, the same in both"


def error_values(run: Run) -> str:
    formula, values = run.folder / "errors-formula.xlsx", run.folder / "errors-formula-values.xlsx"
    error_keys(run.formula, formula)
    run.macro(formula, "-", values)
    errors = error_cells(values)
    if errors:
        raise Failure(f"{len(errors)} output cell(s) hold an error value: "
                      + ", ".join(errors[:compare_outputs.SHOWN]))
    workbook = openpyxl.load_workbook(values, data_only=True)
    calc, first = workbook["IRCalc"], FIRST_ROW + len(ssa_reference.IR)
    # The Issue key and the net value of the two rows of the issue, the
    # currency and the status of the third row, and the PositionID typed as
    # an error value.
    read = ([calc.cell(first + offset, column).value for offset in (0, 1) for column in (19, 20)]
            + [calc.cell(first + 2, column).value for column in (5, 4)]
            + [workbook["EQCalc"].cell(FIRST_ROW + len(ssa_reference.EQ), 3).value])
    expected = [ERROR_TEXT, 2e6, ERROR_TEXT, 2e6, ERROR_TEXT, "INVALID: currency not in FX_Table",
                ERROR_TEXT]
    if not all(compare_outputs.same(*pair) for pair in zip(read, expected, strict=True)):
        raise Failure(f"read {read}, expected {expected}")
    # Months of Eligibility under the CRR regime with error values: only the formula
    # workbook, since LibreOffice's engine reads an error value typed in a cell as 0.
    months_formula, months_values = run.folder / "errors-months.xlsx", run.folder / "errors-months-values.xlsx"
    rows = error_months()
    with_months(run.formula, months_formula, rows)
    run.macro(months_formula, "-", months_values)
    errors = error_cells(months_values)
    differences = eligibility_differences(months_values, ssa.eligibility(sheet_months(rows), "CRR", {}), "formula")
    if errors or differences:
        raise Failure("\n".join([f"months with error values: {len(errors)} output cell(s) hold an error value, "
                                 f"{len(differences)} difference(s) from ssa.eligibility",
                                 *(errors + differences)[:compare_outputs.SHOWN]]))
    return (f"no error value in an output area; every such text reads {ERROR_TEXT}; "
            "months of Eligibility with error values are rejected as ssa.eligibility rejects them")


def stopped_run(run: Run) -> str:
    workbook, output = run.folder / "unusable.xlsm", run.folder / "vba-unusable.xlsx"
    unusable(run.workbook, workbook)
    result = run.macro(workbook, RUN, output)
    listed = findings(output)
    if not result.startswith("RESULT=STOPPED; errors=2;") or len(listed) != 2:
        raise Failure(f"{result}; findings: {sorted(map(str, listed))}")
    return result


def crr_regime(run: Run) -> str:
    """The CRR regime on the sample data: the transitional factor and the de minimis in both workbooks, and the stops."""
    day = datetime.date
    reporting = ssa_reference.AS_OF
    last = day(2029, 12, 31)
    # The overall net open position of the sample is 335 million: 2% of 10 billion is 200 million, of 17 billion
    # 340 million, of 16 billion 320 million.
    cases: tuple[tuple[str, str, object, float, float, bool], ...] = (
        ("the default last date", "crr", last, 1e10, 0.9, False),
        ("the reporting date as last date", "CRR", reporting, 1e10, 0.9, False),
        ("the day before the reporting date as last date", "CRR",
         reporting - datetime.timedelta(days=1), 1e10, 1.0, False),
        ("own funds whose 2% is above the position", "CRR", last, 1.7e10, 0.9, True),
        ("own funds whose 2% is below the position", "CRR", last, 1.6e10, 0.9, False))
    lines: list[str] = []
    # The sample's rows are given the CRR categories that have their specific-risk charges, so that the
    # requirement before the multiplier is the reference model's under the CRR (150% between zones 1 and 3,
    # no charge on an index), unless the FX requirement is waived.
    baseline = ssa_reference.calc_all(1.50, 0.0)["total"]
    for label, name, end, own_funds, factor, waived in cases:
        formula, workbook = run.folder / "crr-formula.xlsx", run.folder / "crr.xlsm"
        values, output = run.folder / "crr-formula-values.xlsx", run.folder / "vba-crr.xlsx"
        regime(run.formula, formula, name, end, own_funds)
        regime(run.workbook, workbook, name, end, own_funds)
        run.macro(formula, "-", values)
        result = run.macro(workbook, RUN, output)
        if not result.startswith("RESULT=OK;"):
            raise Failure(f"{label}: {result}")
        compared = run.compare(output, values)
        for path in (output, values):
            before, after, rwa, onop, charge = read_names(
                path, "SSA_RequirementBeforeTransition", "SSA_Requirement", "SSA_RWA", "FX_ONOP", "FX_Charge")
            expected_charge = 0.0 if waived else 0.08 * onop
            if not (waived or compare_outputs.same(before, baseline)):
                raise Failure(f"{label}, {path.name}: before {before}, the reference model {baseline}")
            if not (compare_outputs.same(after, before * factor)
                    and compare_outputs.same(rwa, after * 12.5)
                    and compare_outputs.same(charge, expected_charge)):
                raise Failure(f"{label}, {path.name}: before {before}, after {after}, RWA {rwa}, "
                              f"FX charge {charge} on {onop}, factor expected {factor}, "
                              f"waived expected {waived}")
        lines.append(f"{label}: factor {factor}, FX {'waived' if waived else 'charged'}, {compared}")
    # A regime, a last date and total own funds that stop the run.
    stops: tuple[tuple[str, str, object, object], ...] = (
        ("an unknown regime", "EU", last, 1e10),
        ("the CRR with a last date that is no date", "CRR", "never", 1e10),
        ("the CRR without total own funds", "CRR", last, None))
    for stop, name, end, funds in stops:
        workbook, output = run.folder / "crr-stop.xlsm", run.folder / "vba-crr-stop.xlsx"
        regime(run.workbook, workbook, name, end, funds)
        result = run.macro(workbook, RUN, output)
        listed = findings(output)
        if not result.startswith("RESULT=STOPPED; errors=1;") or len(listed) != 1:
            raise Failure(f"{stop}: {result}; findings: {sorted(map(str, listed))}")
        lines.append(f"{stop}: stopped, one finding")
    return "\n".join(lines)


def zone_rows() -> list[list[object]]:
    """Two swap legs in one currency that leave a position to offset between zones 1 and 3 only.

    The long leg reprices in nine months and the short one in twelve years, so that the zone 2 is empty
    and the whole of the zone 1 position is matched against the zone 3."""
    day = datetime.timedelta
    return [
        ["IR-Z13A", "T", "Swap leg", "USD", "Long", 100e6, "Floating", None,
         ssa_reference.AS_OF + day(days=3650), ssa_reference.AS_OF + day(days=270), "NONE", "", ""],
        ["IR-Z13B", "T", "Swap leg", "USD", "Short", 100e6, "Floating", None,
         ssa_reference.AS_OF + day(days=5000), ssa_reference.AS_OF + day(days=4380), "NONE", "", ""]]


def crr_zones(run: Run) -> str:
    """The offset between zones 1 and 3 of both regimes, against the reference model, in both workbooks."""
    lines: list[str] = []
    extra = [(*row, "") for row in zone_rows()]
    for name, rate, field in (("Basel", 1.00, "SSA_Requirement"),
                              ("CRR", 1.50, "SSA_RequirementBeforeTransition")):
        paths = []
        for source, suffix in ((run.formula, "xlsx"), (run.workbook, "xlsm")):
            staged = run.folder / f"zone-source.{suffix}"
            workbook = load(source)
            append_rows(workbook["IRPositions"], FIRST_ROW + len(ssa_reference.IR), zone_rows())
            workbook.save(staged)
            target = run.folder / f"zone-{name}.{suffix}"
            regime(staged, target, name, datetime.date(2029, 12, 31), 1e10)
            paths.append(target)
        formula, workbook_path = paths
        values, output = run.folder / "zone-formula-values.xlsx", run.folder / "vba-zone.xlsx"
        run.macro(formula, "-", values)
        result = run.macro(workbook_path, RUN, output)
        if not result.startswith("RESULT=OK;"):
            raise Failure(f"{name}: {result}")
        compared = run.compare(output, values)
        with unittest.mock.patch.object(ssa_reference, "IR", [*ssa_reference.IR, *extra]):
            expected = ssa_reference.calc_all(rate, 0.0 if name == "CRR" else 0.02)["total"]
        for path in (output, values):
            (given,) = read_names(path, field)
            if not compare_outputs.same(given, expected):
                raise Failure(f"{name}, {path.name}: {field} {given}, the reference model {expected} "
                              f"with {rate:.0%} between zones 1 and 3")
        lines.append(f"{name}: {rate:.0%} between zones 1 and 3, requirement {expected:,.2f}, {compared}")
    return "\n".join(lines)


def crr_categories(run: Run) -> str:
    """The categories of the other regime, and a weight that the table lacks, in both workbooks."""
    lines: list[str] = []
    cases = (("the CRR with the Basel categories", "CRR", "keep"),
             ("Basel with the categories CRR", "Basel", "mixed"),
             ("the CRR with a weight that is not in the table", "CRR", "mixed"))
    for label, name, rows in cases:
        formula, workbook = run.folder / "cat-formula.xlsx", run.folder / "cat.xlsm"
        values, output = run.folder / "cat-formula-values.xlsx", run.folder / "vba-cat.xlsx"
        regime(run.formula, formula, name, datetime.date(2029, 12, 31), 1e10, rows)
        regime(run.workbook, workbook, name, datetime.date(2029, 12, 31), 1e10, rows)
        run.macro(formula, "-", values)
        result = run.macro(workbook, RUN, output)
        if not result.startswith("RESULT=INCOMPLETE;"):
            raise Failure(f"{label}: {result}")
        compared = run.compare(output, values)
        engine, reference = findings(output), findings(values)
        if engine != reference:
            raise Failure(f"{label}: findings differ: {sorted(map(str, engine ^ reference))}")
        lines.append(f"{label}: {len(engine)} findings, the same in both; {compared}")
    return "\n".join(lines)


# Eligibility: the rows of the months, the first column of the calculated columns
# (Month status, I), the result block (D32:J33) and the rows of rules C15 and C16 on Checks.
ELIGIBILITY_ROWS = 24
STATUS_COLUMN, RESULT_ROW, RESULT_COLUMN = 9, 32, 4
C15_ROW, C16_ROW, COUNT_COLUMN = 21, 22, 6
APPROACHES = ("325a", "94")
FIGURES = ("size", "share", "result", "run", "count", "event")
RESULTS = ("status", "last_month", "size", "share", "run", "count", "restart")
Month = tuple[object, object, object, object, object, object]


def months(pattern: str, *, last: datetime.date = ssa_reference.AS_OF) -> list[Month]:
    """One row per letter, the last month in the month of ``last``: M where both approaches
    meet the conditions, N where Article 325a misses them and Article 94 meets them. With total
    assets of 8 billion the amounts bind: 400 and 600 million for Article 325a, 40 million for
    Article 94."""
    rows: list[Month] = []
    for k, flag in enumerate(pattern):
        size = 600e6 if flag == "N" else 400e6
        rows.append((ssa_reference.month_end(last, len(pattern) - 1 - k), 8e9,
                     0.6 * size, -0.4 * size, 28e6, 12e6))
    return rows


def eligibility_scenarios() -> tuple[tuple[str, list[Month | None]], ...]:
    """The series of months, each with its rows from row 5 down; None leaves a row blank."""
    day = datetime.date
    invalid: list[Month | None] = [
        (day(2025, 10, 31), 8e9, 1e8, 0, 1e7, 0),
        # A month end that is a text, one that is a number but no date, one that is not the
        # last day of its month and one that is not one month after the row before it.
        ("2025-11-30", 8e9, 1e8, 0, 1e7, 0),
        (20251231, 8e9, 1e8, 0, 1e7, 0),
        (day(2026, 1, 15), 8e9, 1e8, 0, 1e7, 0),
        (day(2026, 3, 31), 8e9, 1e8, 0, 1e7, 0),
        # Total assets of 0, a short total that is a text, a row without a month end and one
        # whose month end holds only blanks, which is read as empty.
        (day(2026, 4, 30), 0, 1e8, 0, 1e7, 0),
        (day(2026, 5, 31), 8e9, 1e8, "none", 1e7, 0),
        (None, 8e9, 1e8, 0, 1e7, 0),
        ("   ", 8e9, 1e8, 0, 1e7, 0),
        (day(2026, 6, 30), 8e9, 1e8, 0, 1e7, 0)]
    return (
        ("the sample months", list(ssa_reference.ELIGIBILITY)),
        ("24 months that meet the conditions", list(months("M" * 24))),
        ("a run of three after two blank rows", [None, None, *months("MMMMMMMMNNN")]),
        ("seven of the last twelve not met, the last met", list(months("MMMMNMNNMNNMNNMM"))),
        ("a row of each invalid kind", invalid))


def with_months(source: Path, target: Path, rows: Sequence[Month | None]) -> None:
    """Save ``source`` with the CRR regime and ``rows`` on Eligibility, the other rows blank."""
    regime(source, target, "CRR", datetime.date(2029, 12, 31), 1e10)
    workbook = load(target)
    sheet = workbook["Eligibility"]
    for offset in range(ELIGIBILITY_ROWS):
        record = rows[offset] if offset < len(rows) else None
        for column in range(6):
            sheet.cell(FIRST_ROW + offset, FIRST_COLUMN + column).value = (
                None if record is None else record[column])
    workbook.save(target)


def sheet_months(rows: Sequence[Month | None]) -> list[Month]:
    """The 24 rows of Eligibility as ssa.eligibility reads them: a blank row is six empty cells."""
    blank: Month = (None, None, None, None, None, None)
    padded = [blank if row is None else row for row in rows]
    return padded + [blank] * (ELIGIBILITY_ROWS - len(padded))


def same_cell(given: object, expected: object) -> bool:
    """A cell of the engine against a value of ssa.eligibility: a date as its serial number."""
    if isinstance(expected, datetime.date) and not isinstance(expected, datetime.datetime):
        expected = datetime.datetime.combine(expected, datetime.time())
    return compare_outputs.same(given, expected)


def eligibility_differences(path: Path, expected: dict[str, Any], source: str = "VBA") -> list[str]:
    """The cells of Eligibility and the counts of rules C15 and C16 that differ from ``expected``,
    what ssa.eligibility gives; ``source`` names the workbook in the messages."""
    workbook = openpyxl.load_workbook(path, data_only=True)
    sheet, checks = workbook["Eligibility"], workbook["Checks"]
    differences: list[str] = []

    def check(cell: Any, value: object) -> None:
        if not same_cell(cell.value, value):
            differences.append(f"Eligibility!{cell.coordinate}: {source}={cell.value!r} ssa.py={value!r}")

    for offset, item in enumerate(expected["rows"]):
        row = FIRST_ROW + offset
        check(sheet.cell(row, STATUS_COLUMN), None if item is None else item["status"])
        for k, approach in enumerate(APPROACHES):
            figures = None if item is None else item[approach]
            for j, key in enumerate(FIGURES):
                check(sheet.cell(row, STATUS_COLUMN + 1 + 6 * k + j), None if figures is None else figures[key])
    for k, approach in enumerate(APPROACHES):
        for j, key in enumerate(RESULTS):
            value = expected[key] if key == "last_month" else expected[key][approach]
            check(sheet.cell(RESULT_ROW + k, RESULT_COLUMN + j), value)
    rejected = sum(1 for item in expected["rows"] if item is not None and item["status"] != ssa.OK)
    flagged = sum(1 for status in expected["status"].values() if status in ("NOT MET", "CEASE"))
    for row, count in ((C15_ROW, rejected), (C16_ROW, flagged)):
        cell = checks.cell(row, COUNT_COLUMN)
        if not compare_outputs.same(cell.value, count):
            differences.append(f"Checks!{cell.coordinate}: {source}={cell.value!r} ssa.py={count}")
    return differences


def eligibility(run: Run) -> str:
    """Series of months under the CRR regime: the engine against the formula workbook and ssa.eligibility."""
    lines: list[str] = []
    for label, rows in eligibility_scenarios():
        formula, workbook = run.folder / "months-formula.xlsx", run.folder / "months.xlsm"
        values, output = run.folder / "months-formula-values.xlsx", run.folder / "vba-months.xlsx"
        with_months(run.formula, formula, rows)
        with_months(run.workbook, workbook, rows)
        run.macro(formula, "-", values)
        result = run.macro(workbook, RUN, output)
        expected = ssa.eligibility(sheet_months(rows), "CRR", {})
        # The sample positions have the CRR categories: only an invalid month makes the run incomplete.
        invalid = any(item is not None and item["status"] != ssa.OK for item in expected["rows"])
        if not result.startswith("RESULT=INCOMPLETE;" if invalid else "RESULT=OK;"):
            raise Failure(f"{label}: {result}")
        compared = run.compare(output, values)
        differences = eligibility_differences(output, expected)
        if differences:
            raise Failure("\n".join([f"{label}: {len(differences)} difference(s) from ssa.eligibility",
                                     *differences[:compare_outputs.SHOWN]]))
        block = openpyxl.load_workbook(output, data_only=True)["Eligibility"]
        statuses = [block.cell(RESULT_ROW + k, RESULT_COLUMN).value for k in range(len(APPROACHES))]
        lines.append(f"{label}: {statuses[0]} and {statuses[1]}, as ssa.eligibility gives; {compared}")
    return "\n".join(lines)


def test_suite(run: Run) -> str:
    results = []
    for macro in TESTS:
        result = run.macro(run.workbook, macro, run.folder / f"vba-{macro}.xlsx")
        if not result.startswith("RESULT=PASS;"):
            raise Failure(f"{macro}: {result or 'no result'}; python3 tools/devbuild/lo_smoke.py "
                          "prints the failures of the run")
        results.append(result)
    return "\n".join(results)


def test_cases(run: Run) -> str:
    # Imported here: run_cases imports this module.
    import run_cases

    count, compared, report = run_cases.check(run, [])
    if count == 0 or any(line.startswith("  ") for line in report):
        raise Failure("\n".join(report) or "no test case found")
    return f"{count} case(s), {compared} expected value(s), all in the workbook"


STEPS: tuple[tuple[str, Callable[[Run], str]], ...] = (
    ("Build the formula workbook and the development workbook", build),
    ("Recalculate the formula workbook", formula_tests),
    ("RunSSA on the sample data against the formula workbook", sample),
    ("Negative scenario on both workbooks", negative_scenario),
    ("Error values typed as texts, on the formula workbook", error_values),
    ("Unusable Params values stop the run", stopped_run),
    ("The CRR regime: the transitional factor and the stops", crr_regime),
    ("The categories of each regime", crr_categories),
    ("The offset between zones 1 and 3 of each regime", crr_zones),
    ("The three suites of the VBA tests", test_suite),
    ("The registered test cases on the engine", test_cases),
    ("Eligibility under the CRR regime against the formula workbook and ssa.eligibility", eligibility),
)


def verify(run: Run) -> bool:
    for number, (title, step) in enumerate(STEPS, 1):
        print(f"\n== {number}. {title}")
        try:
            print(step(run))
        except Failure as failure:
            print(f"{failure}\n\nVERIFY: FAIL (step {number}: {title})")
            return False
    print("\nVERIFY: PASS (LibreOffice; never Excel evidence)")
    return True


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--uno-python", default=sys.executable,
                        help="the Python that provides the UNO bridge (default: this one)")
    parser.add_argument("--keep", type=Path, metavar="FOLDER",
                        help="keep the workbooks of the run in FOLDER")
    args = parser.parse_args(argv)
    # The negative scenario types a number into a date cell on purpose.
    warnings.filterwarnings("ignore", message="Cell .* is marked as a date", category=UserWarning)
    if shutil.which("soffice") is None:
        print("verify: SKIP (needs LibreOffice; this is never Excel evidence)")
        return 2
    if args.keep is not None:
        args.keep.mkdir(parents=True, exist_ok=True)
        return 0 if verify(Run(args.keep, args.uno_python)) else 1
    with tempfile.TemporaryDirectory(prefix="verify-") as temporary:
        return 0 if verify(Run(Path(temporary), args.uno_python)) else 1


if __name__ == "__main__":
    sys.exit(main())
