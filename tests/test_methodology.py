"""The methodology guides against the VBA source, the template and the test
cases: the status texts of a rejected row, the cases each calculation contract
names, the reference values of the parameters, and the datasets and fields of
the data contract."""

from __future__ import annotations

import json
import re
import unittest
from pathlib import Path
from typing import Any

try:
    import openpyxl
except ImportError:  # pragma: no cover - installed from .github/requirements.txt
    openpyxl = None

ROOT = Path(__file__).resolve().parents[1]
CONTRACTS = ROOT / "docs" / "methodology" / "CALCULATION_CONTRACTS.md"
DATA_CONTRACT = ROOT / "docs" / "methodology" / "data-contract.json"
TEMPLATE = ROOT / "src" / "workbook" / "Template.xlsx"
FIXTURES = ROOT / "tests" / "fixtures"
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
                # A text that the source continues on the next line is one text.
                source = re.sub(r'"\s*&\s*_\s*\n\s*"', "", vba(module))
                given = {PREFIXES[prefix] + text
                         for prefix, text in STATUS_IN_SOURCE.findall(source)}
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


class ValidationParagraphs(unittest.TestCase):
    def test_each_contract_names_the_cases_whose_fixture_cites_it(self) -> None:
        cited: dict[str, set[str]] = {}
        for path in sorted(FIXTURES.glob("*.json")):
            fixture = json.loads(path.read_text(encoding="utf-8"))
            cited.setdefault(fixture["calculation"], set()).add(fixture["id"])
        for number in range(2, 9):
            calculation = f"CALC-0{number}"
            with self.subTest(calculation=calculation):
                text = contract(calculation)
                paragraph = text[text.index("**Validation:**"):].split("\n\n")[0]
                named = set(re.findall(r"`([a-z0-9]+(?:-[a-z0-9]+)+)`", paragraph))
                self.assertTrue(cited[calculation])
                self.assertEqual(named, cited[calculation])
                self.assertNotIn("no test case", paragraph)


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
        self.assertEqual(len(self.reference), 21)
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


def workbook_name(field: dict[str, Any]) -> str:
    """The workbook name a field of the parameters is found by, from its notes."""
    match = re.match(r"Name (\w+)", field["notes"])
    assert match is not None, field["name"]
    return match.group(1)


@unittest.skipIf(openpyxl is None, "needs openpyxl, which .github/requirements.txt pins")
class DataContract(unittest.TestCase):
    """The data contract against the template and the columns the engine reads."""

    # Dataset, sheet and the constant of CORE_Config with the columns read.
    POSITIONS = (("DS-06", "IRPositions", "IR_INPUT_WIDTH"),
                 ("DS-07", "EQPositions", "EQ_INPUT_WIDTH"),
                 ("DS-08", "COPositions", "CO_INPUT_WIDTH"),
                 ("DS-09", "FXPositions", "FX_INPUT_WIDTH"))
    datasets: dict[str, dict[str, Any]]
    template: Any

    @classmethod
    def setUpClass(cls) -> None:
        assert openpyxl is not None
        contract = json.loads(DATA_CONTRACT.read_text(encoding="utf-8"))
        cls.datasets = {dataset["id"]: dataset for dataset in contract["datasets"]}
        cls.template = openpyxl.load_workbook(TEMPLATE)

    def headers(self, sheet: str, row: int = 4) -> list[str]:
        cells = self.template[sheet][row][2:]
        return [str(cell.value) for cell in cells if cell.value is not None]

    def test_each_input_sheet_has_a_dataset(self) -> None:
        sources = {dataset["source"] for dataset in self.datasets.values()
                   if not dataset["illustrative"]}
        self.assertEqual(sources, {"Params", "IRPositions", "EQPositions", "COPositions",
                                   "FXPositions", "OptionCharges"})

    def test_the_fields_of_a_position_dataset_are_the_columns_of_its_sheet(self) -> None:
        config = vba("CORE_Config.bas")
        for dataset, sheet, width in (*self.POSITIONS, ("DS-10", "OptionCharges", None)):
            with self.subTest(dataset=dataset):
                fields = self.datasets[dataset]["fields"]
                headers = self.headers(sheet)
                self.assertEqual(len(fields), len(headers))
                for field, header in zip(fields, headers, strict=True):
                    self.assertRegex(field["notes"], rf"^Column {re.escape(header)}(\.|,|$)",
                                     field["name"])
                if width is not None:
                    read = re.search(rf"Public Const {width}\s+As Long = (\d+)", config)
                    assert read is not None
                    self.assertEqual(int(read.group(1)), len(fields))

    def test_the_parameters_are_those_of_the_template_in_its_order(self) -> None:
        sheet = self.template["Params"]
        codes = []
        row = 6
        while sheet.cell(row, 3).value is not None:
            codes.append(str(sheet.cell(row, 3).value))
            row += 1
        fields = self.datasets["DS-02"]["fields"]
        self.assertEqual([workbook_name(field) for field in fields], codes)
        numeric = constant("PRM_NUMERIC_CODES").split("|")
        self.assertEqual(set(codes), {"AsOfDate", "ReportingCcy", "GoldCode", "Regime",
                                      "TransitionalEndDate", "TotalOwnFunds", *numeric})
        # A rate, by the engine's own test, lies from 0 to 1.
        domains = dict(zip(numeric, constant("PRM_DOMAINS"), strict=True))
        by_code = {workbook_name(field): field for field in fields}
        for code, domain in domains.items():
            with self.subTest(code=code):
                expected = {"min": 0, "max": 1} if domain == "R" else {"min": 0}
                self.assertEqual(by_code[code]["allowed"], expected)

    def test_the_lists_of_the_template_are_the_allowed_values_of_their_fields(self) -> None:
        fields = {field["name"]: field for field in self.datasets["DS-06"]["fields"]}
        lists = {str(rule.sqref).split(":")[0]: str(rule.formula1).strip('"').split(",")
                 for rule in self.template["IRPositions"].data_validations.dataValidation
                 if rule.type == "list" and str(rule.formula1).startswith('"')}
        self.assertEqual(lists["E5"], fields["instrument"]["allowed"]["values"])
        self.assertEqual(lists["G5"], fields["direction"]["allowed"]["values"])
        self.assertEqual(lists["I5"], fields["rate_type"]["allowed"]["values"])
        self.assertEqual(lists["M5"], fields["sr_category"]["allowed"]["values"])

    def test_the_tables_of_params_have_the_rows_the_contract_states(self) -> None:
        names = self.template.defined_names

        def rows(name: str) -> int:
            first, last = names[name].attr_text.split("!")[1].replace("$", "").split(":")
            return int(re.sub(r"\D", "", last)) - int(re.sub(r"\D", "", first)) + 1

        self.assertEqual((rows("SR_Key"), rows("Ladder_RW"), rows("FX_Ccy")), (31, 15, 20))
        for dataset, count in (("DS-03", "31 rows"), ("DS-04", "15 rows"), ("DS-05", "20 rows")):
            with self.subTest(dataset=dataset):
                self.assertIn(count, self.datasets[dataset]["granularity"])


if __name__ == "__main__":
    unittest.main()
