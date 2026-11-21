#!/usr/bin/env python3
"""Run the registered test cases on the VBA engine in LibreOffice.

Usage: python3 tools/prototype/run_cases.py [--uno-python PYTHON] [--keep FOLDER] [CASE ...]

For each JSON fixture of tests/fixtures/ whose calculation is one of the
project's, CALC-02 to CALC-09, it writes the fixture's parameters and
records to the input sheets of the development workbook that
tools/devbuild/build_workbook.py builds, runs RunSSA through tools/devbuild/lo_macro.py, and compares every
output of the expected file of tests/expected/ with the cell that holds it:
a number within the output's tolerance, a text exactly. With CASE arguments
it runs those cases only.

It prints one line per case and the outputs that differ, and exits 0 when
every output of every case agrees, 1 when one differs or a run fails, and 2
when LibreOffice is missing.

LibreOffice is not a supported host: this is never Excel evidence, and a
rule is not validated by it. It shows that the expected values, which a
derivation script computes from the calculation contracts, and the engine
agree on the cases.
"""

from __future__ import annotations

import argparse
import datetime
import json
import shutil
import sys
import tempfile
import warnings
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))

import verify  # noqa: E402

ROOT = verify.ROOT
FIXTURES = ROOT / "tests" / "fixtures"
EXPECTED = ROOT / "tests" / "expected"
CALCULATIONS = frozenset(f"CALC-0{number}" for number in range(2, 10))
FIRST_ROW, FIRST_COLUMN, CAPACITY, FX_ROWS = verify.FIRST_ROW, verify.FIRST_COLUMN, 500, 20
# Dataset: its sheet and, in column order, the fixture key of each column
# (None: a column the engine does not read, left empty).
POSITIONS: dict[str, tuple[str, tuple[str | None, ...]]] = {
    "DS-06": ("IRPositions", (None, "instrument_code", "currency_code", "direction_code",
                              "position_value_amount", "rate_type_code", "coupon_rate",
                              "final_maturity_date", "next_repricing_date", "sr_category_code",
                              "rating_grade_code", "issue_id", None)),
    "DS-07": ("EQPositions", (None, "market_code", "type_code", "identifier_code", "currency_code",
                              "direction_code", "market_value_amount", None)),
    "DS-08": ("COPositions", (None, "commodity_code", "direction_code", "quantity_count", None,
                              "spot_price_amount", "currency_code", None, None)),
    "DS-09": ("FXPositions", (None, "currency_code", "direction_code", "net_position_amount", None)),
}
# A quantity of one position row: its calculation sheet and column.
ROW_OUTPUTS = {"ir_position_status": ("IRCalc", 4), "ir_time_band": ("IRCalc", 11),
               "ir_weighted_position": ("IRCalc", 14), "eq_position_status": ("EQCalc", 4),
               "eq_net_position": ("EQCalc", 11), "co_position_status": ("COCalc", 4),
               "fx_position_status": ("FXCalc", 4)}
# A quantity of one currency: the columns of its row of IR_Zones that add up to it.
ZONES_OUTPUTS = {"ir_vertical_disallowance": (5,), "ir_horizontal_disallowance_within_zones": (9, 14, 19),
                 "ir_horizontal_disallowance_between_zones": (28,), "ir_net_position_charge": (29,),
                 "ir_general_market_risk_charge": (30,), "ir_specific_risk_charge": (31,)}
# A quantity of the whole calculation: its cell of Results, by risk class row and column.
RESULTS_ROW = {"IR": 7, "EQ": 8, "CO": 9, "FX": 10}
TOTALS = {"ir_specific_risk_charge": ("IR", 5), "ir_general_market_risk_charge": ("IR", 6),
          "eq_specific_risk_charge": ("EQ", 5), "eq_general_market_risk_charge": ("EQ", 6),
          "eq_index_charge": ("EQ", 7), "co_directional_charge": ("CO", 6),
          "co_gross_charge": ("CO", 7), "fx_charge": ("FX", 6)}
BY_CLASS = {"option_charge": 8, "requirement_before_scaling": 9, "requirement_after_scaling": 11}
NAMED = {"fx_overall_net_open_position": "FX_ONOP", "own_funds_requirement": "SSA_Requirement",
         "risk_weighted_assets": "SSA_RWA"}
# The currencies of the FX block of Buckets: the column of the code and of the net position.
FX_BLOCK_CODE, FX_BLOCK_NET = 4, 8
# A parameter a fixture may set: the workbook name of its cell, and its value when the
# fixture does not set it (the regime's reference value; no total own funds).
PARAMETERS = {"regime_code": ("Regime", "Basel"), "total_own_funds_amount": ("TotalOwnFunds", None)}
# The Eligibility sheet (DS-11): its rows, the fixture keys of its input columns in their
# order, the column of the status of a month and, for each approach, the column of each
# figure; the result block, one row per approach, and the column of each of its outputs.
MONTH_ROWS = 24
MONTH_KEYS = ("month_end_date", "total_assets_amount", "long_325a_amount", "short_325a_amount",
              "long_94_amount", "short_94_amount")
MONTH_STATUS_COLUMN = 9
APPROACHES = ("325a", "94")
MONTH_FIGURES = ("size", "share", "month_result", "run", "count", "event")
BLOCK_ROW, BLOCK_COLUMN = 32, 4
BLOCK_OUTPUTS = {"status": 0, "restart": 6}


def cases(names: list[str]) -> list[dict[str, Any]]:
    """The fixtures to run: those named, or every JSON fixture of the project's calculations."""
    found = []
    for path in sorted(FIXTURES.glob("*.json")):
        fixture = json.loads(path.read_text(encoding="utf-8"))
        if (not names and fixture.get("calculation") in CALCULATIONS) or fixture["id"] in names:
            found.append(fixture)
    return found


def eligibility_cell(quantity: str) -> tuple[int, int] | None:
    """The cell of an output of CALC-09: the column on the row of its month, or the cell of the
    result block (row, column); None for a quantity of another calculation."""
    if not quantity.startswith("eligibility_"):
        return None
    if quantity == "eligibility_month_status":
        return 0, MONTH_STATUS_COLUMN
    name, approach = quantity[len("eligibility_"):].rsplit("_", 1)
    k = APPROACHES.index(approach)
    if name in BLOCK_OUTPUTS:
        return BLOCK_ROW + k, BLOCK_COLUMN + BLOCK_OUTPUTS[name]
    return 0, MONTH_STATUS_COLUMN + 1 + len(MONTH_FIGURES) * k + MONTH_FIGURES.index(name)


def cell_value(key: str, value: object) -> object:
    """A fixture value as the cell holds it: an ISO date as a date."""
    if key.endswith("_date") and isinstance(value, str):
        try:
            return datetime.date.fromisoformat(value)
        except ValueError:
            return value
    return value


def write_inputs(workbook: Any, fixture: dict[str, Any]) -> None:
    """Replace the sample data of the workbook with the records of a fixture."""
    datasets = fixture["datasets"]
    verify.named_cell(workbook, "AsOfDate").value = datetime.date.fromisoformat(fixture["reporting_date"])
    verify.named_cell(workbook, "ReportingCcy").value = fixture["reporting_currency_code"]
    write_parameters(workbook, fixture)
    codes, rates = verify.first_cell(workbook, "FX_Ccy"), verify.first_cell(workbook, "FX_Rate")
    for offset in range(FX_ROWS):
        codes.offset(offset).value = rates.offset(offset).value = None
        codes.offset(offset, 1).value = None
    for offset, record in enumerate(datasets.get("DS-05", [])):
        codes.offset(offset).value = record["id"]
        rates.offset(offset).value = record["inputs"].get("fx_rate")
    for name in RESULTS_ROW:
        verify.named_cell(workbook, f"Options{name}").value = None
    for record in datasets.get("DS-10", []):
        verify.named_cell(workbook, f"Options{record['id']}").value = record["inputs"].get("charge_amount")
    for dataset, (sheet, keys) in POSITIONS.items():
        target = workbook[sheet]
        for row in range(FIRST_ROW, FIRST_ROW + CAPACITY):
            for column in range(FIRST_COLUMN, FIRST_COLUMN + len(keys) + 1):
                target.cell(row, column).value = None
        for offset, record in enumerate(datasets.get(dataset, [])):
            target.cell(FIRST_ROW + offset, FIRST_COLUMN).value = record["id"]
            for column, key in enumerate(keys, FIRST_COLUMN + 1):
                if key is not None:
                    target.cell(FIRST_ROW + offset, column).value = cell_value(
                        key, record["inputs"].get(key))
    write_months(workbook, fixture)


def write_parameters(workbook: Any, fixture: dict[str, Any]) -> None:
    """Set each parameter a fixture may set: its value in the fixture, or its default."""
    parameters = fixture.get("parameters", {})
    for key in parameters.keys() - PARAMETERS.keys():
        raise ValueError(f"{fixture['id']}: no cell for the parameter {key}")
    for key, (name, default) in PARAMETERS.items():
        verify.named_cell(workbook, name).value = parameters.get(key, default)


def write_months(workbook: Any, fixture: dict[str, Any]) -> None:
    """Replace the rows of Eligibility with the records of DS-11, from the first row."""
    months, records = workbook["Eligibility"], fixture["datasets"].get("DS-11", [])
    if len(records) > MONTH_ROWS:
        raise ValueError(f"{fixture['id']}: more than {MONTH_ROWS} months")
    for offset in range(MONTH_ROWS):
        for column, key in enumerate(MONTH_KEYS, FIRST_COLUMN):
            months.cell(FIRST_ROW + offset, column).value = (
                cell_value(key, records[offset]["inputs"].get(key)) if offset < len(records) else None)


def row_of(sheet: Any, column: int, wanted: str, first: int = 1) -> int | None:
    """The first row, from first, whose cell in a column holds a text, without regard to case."""
    for row in range(first, sheet.max_row + 1):
        if str(sheet.cell(row, column).value or "").strip().upper() == wanted.upper():
            return row
    return None


def actual(workbook: Any, fixture: dict[str, Any], output: dict[str, Any]) -> object:
    """The value the workbook holds for an expected output; a text that says so when it has none."""
    quantity, scope = output["quantity"], output["scope"]
    if (cell := eligibility_cell(quantity)) is not None:
        month_row, column = cell
        if scope is not None:
            month_row = FIRST_ROW + [item["id"] for item in fixture["datasets"]["DS-11"]].index(
                scope["record"])
        return workbook["Eligibility"].cell(month_row, column).value
    if scope is None:
        if quantity in NAMED:
            return verify.named_cell(workbook, NAMED[quantity]).value
        name, column = TOTALS[quantity]
        return workbook["Results"].cell(RESULTS_ROW[name], column).value
    dataset, record = scope["dataset"], scope["record"]
    if dataset == "DS-10":
        return workbook["Results"].cell(RESULTS_ROW[record], BY_CLASS[quantity]).value
    if dataset == "DS-05":
        if quantity == "fx_net_position":
            block = verify.named_cell(workbook, "Buckets_FX_Header")
            row = row_of(block.parent, FX_BLOCK_CODE, record, block.row + 1)
            return "no such currency on Buckets" if row is None else \
                block.parent.cell(row, FX_BLOCK_NET).value
        row = row_of(workbook["IR_Zones"], FIRST_COLUMN, record)
        if row is None:
            return "no such currency on IR_Zones"
        return sum(float(workbook["IR_Zones"].cell(row, column).value or 0)
                   for column in ZONES_OUTPUTS[quantity])
    index = [item["id"] for item in fixture["datasets"][dataset]].index(record)
    sheet, column = ROW_OUTPUTS[quantity]
    return workbook[sheet].cell(FIRST_ROW + index, column).value


def agrees(output: dict[str, Any], value: object) -> bool:
    if "text" in output:
        return bool(value == output["text"])
    wanted = float(output["value"])
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        # An output that is zero may be an empty cell.
        return value in (None, "") and wanted == 0
    tolerance = output["tolerance"]
    bound = max(float(tolerance["absolute"]), float(tolerance["relative"]) * abs(wanted))
    return abs(float(value) - wanted) <= bound


def differences(workbook: Any, fixture: dict[str, Any], expected: dict[str, Any]) -> Iterator[str]:
    for output in expected["outputs"]:
        value = actual(workbook, fixture, output)
        if not agrees(output, value):
            scope = "" if output["scope"] is None else f" of {output['scope']['record']}"
            wanted = output.get("text", output.get("value"))
            yield f"{output['quantity']}{scope}: expected {wanted!r}, the workbook has {value!r}"


def run_case(run: verify.Run, fixture: dict[str, Any]) -> list[str]:
    """Run one case on the built workbook; return what differs from its expected file."""
    expected = json.loads((EXPECTED / f"{fixture['id']}.json").read_text(encoding="utf-8"))
    source, output = run.folder / f"{fixture['id']}.xlsm", run.folder / f"{fixture['id']}-values.xlsx"
    workbook = verify.load(run.workbook)
    write_inputs(workbook, fixture)
    workbook.save(source)
    try:
        run.macro(source, verify.RUN, output)
    except verify.Failure as failure:
        return [f"the run failed: {failure}"]
    return list(differences(openpyxl.load_workbook(output, data_only=True), fixture, expected))


def check(run: verify.Run, names: list[str]) -> tuple[int, int, list[str]]:
    """Cases run, outputs compared and the lines to report."""
    selected = cases(names)
    compared = 0
    report = []
    for fixture in selected:
        expected = json.loads((EXPECTED / f"{fixture['id']}.json").read_text(encoding="utf-8"))
        found = run_case(run, fixture)
        compared += len(expected["outputs"])
        report.append(f"{fixture['id']}: {len(expected['outputs'])} output(s), "
                      + ("all agree" if not found else f"{len(found)} differ"))
        report += [f"  {line}" for line in found]
    return len(selected), compared, report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("cases", nargs="*", metavar="CASE", help="a case ID; default: every case")
    parser.add_argument("--uno-python", default=sys.executable,
                        help="the Python that provides the UNO bridge (default: this one)")
    parser.add_argument("--keep", type=Path, metavar="FOLDER",
                        help="keep the workbooks of the run in FOLDER")
    args = parser.parse_args(argv)
    warnings.filterwarnings("ignore", message="Cell .* is marked as a date", category=UserWarning)
    if shutil.which("soffice") is None:
        print("run-cases: SKIP (needs LibreOffice; this is never Excel evidence)")
        return 2
    with tempfile.TemporaryDirectory(prefix="run-cases-") as temporary:
        folder = args.keep if args.keep is not None else Path(temporary)
        folder.mkdir(parents=True, exist_ok=True)
        run = verify.Run(folder, args.uno_python)
        print(verify.build(run))
        count, compared, report = check(run, args.cases)
    print("\n".join(report))
    failed = any(line.startswith("  ") for line in report)
    print(f"\nRUN-CASES: {'FAIL' if failed else 'PASS'} ({count} case(s), {compared} output(s); "
          "LibreOffice; never Excel evidence)")
    return 1 if failed or count == 0 else 0


if __name__ == "__main__":
    sys.exit(main())
