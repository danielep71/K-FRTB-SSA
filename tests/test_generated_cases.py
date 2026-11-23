"""Generated cases: properties that the derivation script must keep on inputs no fixture holds.

The registered cases check the values someone derived. This module checks what a
list of cases cannot: that the derivation script, ssa.py, keeps the relations of
the calculations on inputs it generates from the registered fixtures, and that it
never fails with an exception on an unusable position value, which the contracts'
Failure behaviour answers with a row status.

Every case is generated from a fixed seed with the standard library, so a failure
repeats: the message names the fixture and the draw. No generator or fuzzing
library is used or needed (see docs/methodology/TEST_CASES.md, Generated cases).

Out of scope, on purpose:
- The VBA engine. It runs in Excel or in LibreOffice (tools/prototype/run_cases.py),
  never in this module, so these properties hold for the derivation only.
- The FX rates (DS-05, field fx_rate): the derivation script does not implement the
  stop rules C05 to C07 yet, so a generated rate would only find that gap, which
  issue 119 tracks. Its acceptance criteria ask for the exclusion below to go.
- The option charge amounts (DS-10): the contracts state that the workbook checks none.
- CALC-01 and CALC-09: CALC-01 is an illustrative stand-in, and CALC-09 reads its
  months in order, so a shuffle changes its result by design.
"""

from __future__ import annotations

import copy
import json
import math
import random
import sys
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests" / "derivations"))

import ssa  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures"
CALCULATIONS = frozenset(f"CALC-0{number}" for number in range(2, 9))
SEED = 20261123
DRAWS = 20
# The datasets of the positions. DS-05 (FX rates) and DS-10 (option charges) are out of scope.
POSITIONS = ("DS-06", "DS-07", "DS-08", "DS-09")
# The quantities that are a charge or a requirement of a risk class, in the reporting currency.
CHARGES = frozenset({
    "ir_specific_risk_charge", "ir_general_market_risk_charge", "eq_specific_risk_charge",
    "eq_general_market_risk_charge", "eq_index_charge", "co_gross_charge", "fx_charge",
    "option_charge", "requirement_before_scaling"})
# The inputs that are an amount or a quantity, which a scale factor multiplies.
AMOUNTS = frozenset({
    "position_value_amount", "market_value_amount", "net_position_amount", "charge_amount",
    "quantity_count"})
OPPOSITE = {"Long": "Short", "Short": "Long"}
# Values a cell can hold that no position field accepts, and the edges of the numbers.
UNUSABLE: tuple[object, ...] = (None, "", " ", "abc", "1,5", "2026-02-30", True, [], {}, 0, -1, 0.5, 1e-12, 1e15)


def fixtures() -> list[tuple[str, dict[str, Any]]]:
    """The registered fixtures of the calculations CALC-02 to CALC-08, by file name."""
    found = []
    for path in sorted(FIXTURES.glob("*.json")):
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, dict) and data.get("calculation") in CALCULATIONS and "datasets" in data:
            found.append((path.name, data))
    return found


def charges(fixture: dict[str, Any]) -> dict[tuple[str, str], float]:
    """The charges of a fixture's derivation by quantity and scope."""
    return {(output["quantity"], json.dumps(output.get("scope"), sort_keys=True)): output["value"]
            for output in ssa.derive(fixture)
            if output["quantity"] in CHARGES and isinstance(output.get("value"), (int, float))}


def stopped(fixture: dict[str, Any]) -> bool:
    return any(output["quantity"] == "run_status" and "STOPPED" in str(output.get("text"))
               for output in ssa.derive(fixture))


def agree(first: dict[tuple[str, str], float], second: dict[tuple[str, str], float],
          factor: float = 1.0) -> bool:
    """Whether every value of the first, times the factor, is the second's within 1e-6 relative."""
    return first.keys() == second.keys() and all(
        math.isclose(factor * first[key], second[key], rel_tol=1e-6, abs_tol=1e-6) for key in first)


def usable(name: str) -> bool:
    """Whether a fixture is one the properties apply to: it runs and is under the Basel regime or
    the properties that do not involve the own funds."""
    return "crr" not in name


class GeneratedCases(unittest.TestCase):
    """The properties, on the registered fixtures and on cases generated from them."""

    def setUp(self) -> None:
        self.cases = fixtures()
        self.assertGreaterEqual(len(self.cases), 30, "the registered fixtures were not found")

    def test_the_order_of_the_records_does_not_change_any_output(self) -> None:
        rng = random.Random(SEED)
        for name, case in self.cases:
            expected = sorted(json.dumps(output, sort_keys=True) for output in ssa.derive(case))
            for draw in range(DRAWS):
                shuffled = copy.deepcopy(case)
                for records in shuffled["datasets"].values():
                    rng.shuffle(records)
                with self.subTest(fixture=name, draw=draw):
                    got = sorted(json.dumps(output, sort_keys=True) for output in ssa.derive(shuffled))
                    self.assertEqual(expected, got)

    def test_a_charge_is_finite_and_never_negative(self) -> None:
        for name, case in self.cases:
            for key, value in charges(case).items():
                with self.subTest(fixture=name, quantity=key[0]):
                    self.assertTrue(math.isfinite(value) and value >= 0, value)

    def test_scaling_every_amount_scales_every_charge_under_the_basel_regime(self) -> None:
        for name, case in self.cases:
            if not usable(name) or stopped(case):
                continue
            for factor in (0.5, 3.0, 1000.0):
                scaled = copy.deepcopy(case)
                for records in scaled["datasets"].values():
                    for record in records:
                        for field, value in record["inputs"].items():
                            if field in AMOUNTS and isinstance(value, (int, float)) \
                                    and not isinstance(value, bool):
                                record["inputs"][field] = value * factor
                with self.subTest(fixture=name, factor=factor):
                    self.assertTrue(agree(charges(case), charges(scaled), factor))

    def test_turning_every_direction_does_not_change_a_charge(self) -> None:
        for name, case in self.cases:
            flipped = copy.deepcopy(case)
            for records in flipped["datasets"].values():
                for record in records:
                    direction = record["inputs"].get("direction_code")
                    if direction in OPPOSITE:
                        record["inputs"]["direction_code"] = OPPOSITE[direction]
            with self.subTest(fixture=name):
                self.assertTrue(agree(charges(case), charges(flipped)))

    def test_an_unusable_position_value_gives_a_status_and_never_an_exception(self) -> None:
        rng = random.Random(SEED)
        for name, case in self.cases:
            for draw in range(DRAWS):
                mutated = copy.deepcopy(case)
                for _ in range(rng.randint(1, 3)):
                    dataset = rng.choice([key for key in POSITIONS if mutated["datasets"].get(key)]
                                         or [None])
                    if dataset is None:
                        break
                    record = rng.choice(mutated["datasets"][dataset])
                    if record["inputs"]:
                        record["inputs"][rng.choice(list(record["inputs"]))] = rng.choice(UNUSABLE)
                with self.subTest(fixture=name, draw=draw):
                    outputs = ssa.derive(mutated)
                    for output in outputs:
                        value = output.get("value")
                        if isinstance(value, float):
                            self.assertTrue(math.isfinite(value), output)
                    for key, value in charges(mutated).items():
                        self.assertGreaterEqual(value, 0, key)

    def test_the_comparison_rejects_a_corrupted_charge(self) -> None:
        # The properties above would pass every output if the comparison could not fail.
        name, case = next((name, case) for name, case in self.cases if charges(case))
        honest = charges(case)
        corrupted = {key: value + 1.0 for key, value in honest.items()}
        self.assertTrue(agree(honest, dict(honest)), name)
        self.assertFalse(agree(honest, corrupted))
        self.assertFalse(agree(honest, {}))


if __name__ == "__main__":
    unittest.main()
