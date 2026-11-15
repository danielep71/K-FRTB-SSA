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
   PositionID is another row's Issue ID, a maturity that is not a date, a
   changed constant and an option charge.
5. Make two Params values unusable: the run stops and lists both.
6. Run TEST_Suite.Test_RunAll: no assertion fails.

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
import warnings
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any

import openpyxl

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))

import build_prototype  # noqa: E402
import build_workbook  # noqa: E402
import compare_outputs  # noqa: E402
import ssa_reference  # noqa: E402

LO_MACRO = ROOT / "tools" / "devbuild" / "lo_macro.py"
RUN = "M_Main.RunSSA_Silent"
TESTS = "TEST_Suite.Test_RunAll"
# The first row of a position table; its first column is C.
FIRST_ROW = 5
FIRST_COLUMN = 3

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


def append_rows(sheet: Any, first: int, records: Sequence[Sequence[object]]) -> None:
    for offset, record in enumerate(records):
        for column, value in enumerate(record, FIRST_COLUMN):
            sheet.cell(first + offset, column).value = value


def negative(source: Path, target: Path) -> None:
    """Save ``source`` with the negative scenario on its input sheets."""
    workbook = load(source)
    day = datetime.date
    append_rows(workbook["IRPositions"], FIRST_ROW + len(ssa_reference.IR), [
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
        ["IR-FRN", "T", "FRN", "EUR", "Short", 3e6, "floating", None, day(2029, 1, 1),
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
    ])
    append_rows(workbook["EQPositions"], FIRST_ROW + len(ssa_reference.EQ), [
        ["EQ-BAD", "T", "IT", "Basket", "B", "EUR", "Long", 1e6],
        ["EQ-NEG", "T", "FR", "Single name", "BNP", "EUR", "Short", 2e6],
    ])
    append_rows(workbook["COPositions"], FIRST_ROW + len(ssa_reference.CO), [
        ["CO-BAD", "T", "Gold bar", "Long", "ten", "oz", 3500, "EUR"],
        ["CO-007", "T", "Wheat", "Short", 1000, "t", 250, "EUR"],
    ])
    append_rows(workbook["FXPositions"], FIRST_ROW + len(ssa_reference.FXPOS), [
        ["FX-007", "FX", "EUR", "Long", 5e6],
        ["FX-008", "FX", "JPY", "Long", -5],
        ["FX-009", "FX", "CHF", "Long", 4e6],
    ])
    named_cell(workbook, "ScalingIR").value = 1.5
    named_cell(workbook, "OptionsFX").value = 1000
    workbook.save(target)


def unusable(source: Path, target: Path) -> None:
    """Save ``source`` with a reporting date and an option charge that are not usable."""
    workbook = load(source)
    named_cell(workbook, "AsOfDate").value = "not a date"
    named_cell(workbook, "OptionsIR").value = "abc"
    workbook.save(target)


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


def stopped_run(run: Run) -> str:
    workbook, output = run.folder / "unusable.xlsm", run.folder / "vba-unusable.xlsx"
    unusable(run.workbook, workbook)
    result = run.macro(workbook, RUN, output)
    listed = findings(output)
    if not result.startswith("RESULT=STOPPED; errors=2;") or len(listed) != 2:
        raise Failure(f"{result}; findings: {sorted(map(str, listed))}")
    return result


def test_suite(run: Run) -> str:
    result = run.macro(run.workbook, TESTS, run.folder / "vba-tests.xlsx")
    if not result.startswith("PASS=") or not result.endswith("FAIL=0"):
        raise Failure(result)
    return result


STEPS: tuple[tuple[str, Callable[[Run], str]], ...] = (
    ("Build the formula workbook and the development workbook", build),
    ("Recalculate the formula workbook", formula_tests),
    ("RunSSA on the sample data against the formula workbook", sample),
    ("Negative scenario on both workbooks", negative_scenario),
    ("Unusable Params values stop the run", stopped_run),
    ("TEST_Suite.Test_RunAll", test_suite),
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
