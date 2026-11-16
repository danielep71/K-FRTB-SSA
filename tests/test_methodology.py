"""The calculation contracts against the VBA source: the status texts of a
rejected row and the reference values of the parameters."""

from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CONTRACTS = ROOT / "docs" / "methodology" / "CALCULATION_CONTRACTS.md"
CORE = ROOT / "src" / "core"
# The contract whose Failure behaviour lists the statuses of a module's rows.
ROW_MODULES = {"CALC-03": "CORE_InterestRate.bas", "CALC-04": "CORE_Equity.bas",
               "CALC-05": "CORE_Commodity.bas", "CALC-06": "CORE_ForeignExchange.bas"}
PREFIXES = {"ST_INVALID": "INVALID: ", "ST_EXCLUDED": "EXCLUDED: "}
STATUS_IN_SOURCE = re.compile(r'\b(ST_INVALID|ST_EXCLUDED) & "([^"]+)"')
STATUS_IN_CONTRACT = re.compile(r"^- `((?:INVALID|EXCLUDED): [^`]+)`$", re.M)


def vba(name: str) -> str:
    return (CORE / name).read_bytes().decode("cp1252").replace("\r\n", "\n")


def contract(calculation: str) -> str:
    """The text of one contract, from its heading to the next heading."""
    text = CONTRACTS.read_text(encoding="utf-8")
    start = text.index(f"### `{calculation}`")
    following = re.search(r"^#{1,3} ", text[start + 4:], re.M)
    return text[start:] if following is None else text[start:start + 4 + following.start()]


def constant(name: str) -> str:
    """The text of a string constant of CORE_Config, with its continuation lines joined."""
    source = vba("CORE_Config.bas")
    match = re.search(rf'Public Const {name}\s+As String = ((?:_\n\s*)?"[^\n]*(?:\n\s*"[^\n]*)*)',
                      source)
    assert match is not None, name
    return "".join(re.findall(r'"([^"]*)"', match.group(1)))


class StatusTexts(unittest.TestCase):
    def test_each_contract_lists_the_statuses_its_module_gives(self) -> None:
        for calculation, module in ROW_MODULES.items():
            with self.subTest(calculation=calculation):
                given = {PREFIXES[prefix] + text
                         for prefix, text in STATUS_IN_SOURCE.findall(vba(module))}
                listed = set(STATUS_IN_CONTRACT.findall(contract(calculation)))
                self.assertTrue(given)
                self.assertEqual(listed, given)

    def test_the_option_charges_name_their_two_findings(self) -> None:
        source = vba("CORE_Parameters.bas")
        text = contract("CALC-07")
        for finding in ("must be a number", "must be 0 or more"):
            with self.subTest(finding=finding):
                self.assertIn(f'ST_INVALID & "{finding}"', source)
                self.assertIn(f"`INVALID: {finding}`", text)


class Parameters(unittest.TestCase):
    def setUp(self) -> None:
        text = CONTRACTS.read_text(encoding="utf-8")
        start = text.index("## Parameters")
        self.section = text[start:text.index("## Contract template")]
        codes = constant("PRM_NUMERIC_CODES").split("|")
        values = constant("PRM_REFERENCE_VALUES").split("|")
        self.assertEqual(len(codes), len(values))
        self.reference = dict(zip(codes, map(float, values), strict=True))

    def test_every_numeric_parameter_is_traced(self) -> None:
        self.assertEqual(len(self.reference), 19)
        for code in self.reference:
            with self.subTest(code=code):
                self.assertIn(f"`{code}`", self.section)

    def test_the_reference_values_of_the_table_are_those_of_the_source(self) -> None:
        compared = 0
        for line in self.section.splitlines():
            cells = [cell.strip() for cell in line.strip("|").split("|")]
            if len(cells) != 4 or not cells[0].startswith("`"):
                continue
            codes = re.findall(r"`([A-Za-z0-9]+)`", cells[0])
            if not all(code in self.reference for code in codes):
                continue  # A table of Params, or a parameter that is a text.
            values = [float(value) for value in cells[1].split(", ")]
            self.assertEqual(len(values), len(codes), line)
            for code, value in zip(codes, values, strict=True):
                with self.subTest(code=code):
                    self.assertEqual(value, self.reference[code])
                    compared += 1
        self.assertEqual(compared, len(self.reference))

    def test_the_ladder_of_the_source_has_the_weights_the_register_was_checked_against(
            self) -> None:
        # Table 4 of MAR40.26, as percentages, compared with the text on 2026-11-16.
        weights = [0, 0.2, 0.4, 0.7, 1.25, 1.75, 2.25, 2.75, 3.25, 3.75, 4.5, 5.25, 6, 8, 12.5]
        self.assertEqual([round(float(value) * 100, 6)
                          for value in constant("LADDER_WEIGHTS").split("|")], weights)
        self.assertEqual(constant("LADDER_ZONES"), "1|1|1|1|2|2|2|3|3|3|3|3|3|3|3")
        self.assertEqual(constant("LADDER_UPPER_LOW").split("|")[4:],
                         ["1.9", "2.8", "3.6", "4.3", "5.7", "7.3", "9.3", "10.6", "12", "20", ""])
        self.assertEqual(constant("LADDER_UPPER_HIGH").split("|")[4:],
                         ["2", "3", "4", "5", "7", "10", "15", "20", "", "", ""])


if __name__ == "__main__":
    unittest.main()
