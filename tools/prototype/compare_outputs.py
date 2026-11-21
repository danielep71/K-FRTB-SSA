#!/usr/bin/env python3
"""Compare the output sheets the VBA engine wrote with the formula workbook's.

Usage: python tools/prototype/compare_outputs.py VBA_RUN.xlsx FORMULA.xlsx

VBA_RUN is a workbook saved after RunSSA; FORMULA is the formula workbook of
build_prototype.py, recalculated on the same inputs. Both must hold values,
as tools/devbuild/lo_macro.py saves them. Every output area is compared cell
by cell: numbers within a relative tolerance of 1e-6, dates as their serial
numbers, text exactly, and an empty cell equal to an empty text.

A calculation row is compared where either workbook has a PositionID, and a
Buckets row where either has a bucket name: the formula workbook numbers its
empty rows, and the engine leaves them blank. Every other area is compared
in full.

It prints the number of cells compared and the differences, and exits 0 when
there is none, 1 when there is one, and 2 when a file is missing.
"""

from __future__ import annotations

import argparse
import datetime
import sys
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl

TOLERANCE = 1e-6
# The day before serial 1: Excel's date serials count from here.
EPOCH = datetime.datetime(1899, 12, 30)
# Differences printed; the count is always complete.
SHOWN = 60


@dataclass(frozen=True)
class Area:
    """Cells of one sheet to compare; ``key`` is the column that marks a used row."""
    sheet: str
    rows: range
    columns: tuple[int, ...]
    key: int | None = None


def columns(first: int, count: int) -> tuple[int, ...]:
    return tuple(range(first, first + count))


# One row per input row, from row 5; the last row is found in the workbooks.
CALCULATION_SHEETS = (("IRCalc", 22), ("EQCalc", 12), ("COCalc", 8), ("FXCalc", 7))
FIRST_CALCULATION_ROW = 5
# Buckets: 30 rows and a total row per block; the next block's rows start four
# rows below a total row.
EQUITY_FIRST = 7
COMMODITY_FIRST = EQUITY_FIRST + 30 + 1 + 4
FX_FIRST = COMMODITY_FIRST + 30 + 1 + 4
FIXED_AREAS = (
    Area("IR_Ladder", range(6, 306), columns(3, 13)),
    Area("IR_Zones", range(6, 27), columns(3, 30)),
    Area("Buckets", range(EQUITY_FIRST, EQUITY_FIRST + 31), columns(3, 10), key=4),
    Area("Buckets", range(COMMODITY_FIRST, COMMODITY_FIRST + 31), columns(3, 10), key=4),
    # The currencies, their total row, the net open position and the charge.
    Area("Buckets", range(FX_FIRST, FX_FIRST + 24), columns(3, 9)),
    Area("Results", range(7, 12), columns(5, 11)),
    Area("Results", range(14, 18), (5,)),
    Area("Checks", range(7, 23), (6, 7)),
    Area("Checks", range(4, 5), (4, 6)),
    # The calculated columns I:U of the 24 months and the result block of each approach.
    Area("Eligibility", range(5, 29), columns(9, 13)),
    Area("Eligibility", range(32, 34), columns(4, 7)),
    Area("TestCatalogue", range(7, 26), (9, 13)),
)


def normalized(value: object) -> object:
    """A cell value in the form two workbooks can be compared in."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, datetime.datetime):
        return (value - EPOCH).total_seconds() / 86400.0
    if isinstance(value, (int, float)):
        return float(value)
    return str(value)


def same(first: object, second: object) -> bool:
    left, right = normalized(first), normalized(second)
    if isinstance(left, bool) or isinstance(right, bool):
        return left == right
    if isinstance(left, float) and isinstance(right, float):
        return abs(left - right) <= TOLERANCE * max(1.0, abs(left), abs(right))
    return left == right


def last_row(sheet: Any, column: int) -> int:
    row: int = sheet.max_row
    while row > 1 and sheet.cell(row, column).value in (None, ""):
        row -= 1
    return row


def compare(vba: Any, formula: Any, area: Area) -> tuple[int, list[str]]:
    """Cells compared and the differences of one area."""
    compared = 0
    differences = []
    for row in area.rows:
        if area.key is not None and all(
                normalized(sheet.cell(row, area.key).value) is None for sheet in (vba, formula)):
            continue
        for column in area.columns:
            compared += 1
            cell = vba.cell(row, column)
            other = formula.cell(row, column).value
            if not same(cell.value, other):
                differences.append(f"{vba.title}!{cell.coordinate}: VBA={cell.value!r} "
                                   f"formula={other!r}")
    return compared, differences


def areas(vba: Any, formula: Any) -> Iterable[Area]:
    for name, width in CALCULATION_SHEETS:
        last = max(last_row(vba[name], 3), last_row(formula[name], 3))
        yield Area(name, range(FIRST_CALCULATION_ROW, last + 1), columns(3, width), key=3)
    yield from FIXED_AREAS


def compare_workbooks(vba_run: Path, formula: Path) -> tuple[int, list[str]]:
    """Cells compared and the differences between two workbooks of values."""
    vba = openpyxl.load_workbook(vba_run, data_only=True)
    reference = openpyxl.load_workbook(formula, data_only=True)
    compared = 0
    differences: list[str] = []
    for area in areas(vba, reference):
        count, found = compare(vba[area.sheet], reference[area.sheet], area)
        compared += count
        differences += found
    return compared, differences


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("vba_run", type=Path, help="the workbook saved after RunSSA")
    parser.add_argument("formula", type=Path, help="the formula workbook, recalculated")
    args = parser.parse_args(argv)
    for path in (args.vba_run, args.formula):
        if not path.is_file():
            print(f"compare-outputs: ERROR (no workbook at {path})", file=sys.stderr)
            return 2
    compared, differences = compare_workbooks(args.vba_run, args.formula)
    print(f"cells compared: {compared}; differences: {len(differences)}")
    for line in differences[:SHOWN]:
        print("  " + line)
    return 1 if differences else 0


if __name__ == "__main__":
    sys.exit(main())
