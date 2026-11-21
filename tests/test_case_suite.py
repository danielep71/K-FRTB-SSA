"""tools/prototype/case_suite.py and the module it writes, tests/modules/TEST_Cases.bas:
the tracked module is what the test cases give, every expected output has one
assertion, and the cells it reads are those the template labels."""

from __future__ import annotations

import contextlib
import io
import json
import re
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any, ClassVar
from unittest import mock

try:
    import openpyxl
except ImportError:  # pragma: no cover - installed from .github/requirements.txt
    openpyxl = None

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "prototype"))

NAME = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")


@unittest.skipIf(openpyxl is None, "needs openpyxl, which .github/requirements.txt pins")
class CaseSuite(unittest.TestCase):
    case_suite: ClassVar[Any]
    run_cases: ClassVar[Any]
    text: ClassVar[str]
    bodies: ClassVar[dict[str, str]]

    @classmethod
    def setUpClass(cls) -> None:
        import case_suite
        import run_cases
        cls.case_suite, cls.run_cases = case_suite, run_cases
        cls.text = case_suite.module().replace("\r\n", "\n")
        pattern = re.compile(r"^Private Sub (\w+)\(.*?^End Sub$", re.M | re.S)
        cls.bodies = {match.group(1): match.group(0) for match in pattern.finditer(cls.text)}

    def main(self, *arguments: str) -> tuple[int, str]:
        with contextlib.redirect_stdout(io.StringIO()) as output, \
                contextlib.redirect_stderr(output):
            code = self.case_suite.main(list(arguments))
        return code, output.getvalue()

    def test_the_tracked_module_is_what_the_test_cases_give(self) -> None:
        code, output = self.main()
        self.assertEqual((code, output), (0, "case-suite: PASS (tests/modules/TEST_Cases.bas "
                                             "is what the test cases give)\n"))
        data = self.case_suite.MODULE.read_bytes()
        # A VBA source file: ASCII, CRLF line ends and nothing else.
        self.assertTrue(data.isascii())
        self.assertNotIn(b"\n", data.replace(b"\r\n", b""))

    def test_a_module_that_differs_or_is_missing_fails_and_write_replaces_it(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "tests" / "modules" / "TEST_Cases.bas"
            path.parent.mkdir(parents=True)
            with mock.patch.object(self.case_suite, "MODULE", path), \
                    mock.patch.object(self.case_suite, "ROOT", Path(directory)):
                self.assertEqual(self.main()[0], 1)
                path.write_bytes(b"edited\r\n")
                code, output = self.main()
                self.assertEqual(code, 1)
                self.assertIn("run python tools/prototype/case_suite.py --write", output)
                self.assertEqual(self.main("--write")[0], 0)
                self.assertEqual(self.main()[0], 0)
                self.assertEqual(path.read_bytes(), (ROOT / "tests" / "modules"
                                                     / "TEST_Cases.bas").read_bytes())

    def test_a_case_the_module_cannot_hold_is_an_error(self) -> None:
        pairs = self.case_suite.loaded()
        fixture = json.loads(json.dumps(pairs[0][0]))
        fixture["reporting_currency_code"] = "€"
        with mock.patch.object(self.case_suite, "loaded", return_value=[(fixture, pairs[0][1])]):
            code, output = self.main()
        self.assertEqual(code, 2)
        self.assertIn("case-suite: ERROR (a text is not printable ASCII", output)

    def test_each_case_asserts_every_expected_output_once_and_records_each_number(self) -> None:
        pairs = self.case_suite.loaded()
        self.assertEqual([fixture["id"] for fixture, _ in pairs],
                         [fixture["id"] for fixture in self.run_cases.cases([])])
        total = 0
        for fixture, expected in pairs:
            with self.subTest(case=fixture["id"]):
                body = self.bodies[f"Case{self.case_suite.pascal(fixture['id'])}"]
                self.assertIn(f'TEST_Harness.BeginCase "{fixture["id"]}"', body)
                names = re.findall(r'TEST_Harness\.Assert\w+ "([^"]+)"', body)
                outputs = expected["outputs"]
                # One per expected output, and the restoration of the sheets, last.
                self.assertEqual(names.pop(), "inputs-restored")
                self.assertEqual(len(names), len(outputs))
                self.assertEqual(len(set(names)), len(names))
                for name in names:
                    self.assertRegex(name, NAME)
                texts = [output for output in outputs if "text" in output]
                self.assertEqual(body.count("TEST_Harness.AssertEqualText "), len(texts))
                self.assertEqual(body.count("TEST_Harness.AssertNear "), len(outputs) - len(texts))
                # Each number is compared and recorded from one reading of the cell.
                self.assertEqual(body.count("TEST_Harness.RecordResult "),
                                 len(outputs) - len(texts))
                self.assertEqual(body.count("        Actual = "), len(outputs) - len(texts))
                # The run is the workbook's entry point, once, after the case's inputs.
                self.assertEqual(body.count("RunSSA_Silent()"), 1)
                self.assertLess(body.index(f"Inputs{self.case_suite.pascal(fixture['id'])}\n"),
                                body.index("RunSSA_Silent()"))
                total += len(outputs)
        constants = dict(re.findall(r"Private Const (EXPECTED_\w+)\s+As Long = (\d+)", self.text))
        self.assertEqual(int(constants["EXPECTED_CASES"]), len(pairs))
        self.assertEqual(int(constants["EXPECTED_ASSERTIONS"]), total + len(pairs))

    def test_the_expected_values_and_tolerances_are_those_of_the_expected_file(self) -> None:
        fixture, expected = next(pair for pair in self.case_suite.loaded()
                                 if pair[0]["id"] == "mar40-61-table-9")
        body = self.bodies["CaseMar4061Table9"]
        self.assertIn('TEST_Harness.AssertNear "fx-overall-net-open-position", 335, Actual, '
                      "0.5, 0\n", body)
        self.assertIn('TEST_Harness.AssertNear "fx-charge", 26.8, Actual, 0.05, 0\n', body)
        self.assertIn("Actual = NumberOf(NamedCell(NM_FX_ONOP))\n", body)
        self.assertIn('TEST_Harness.RecordResult "mar40-61-table-9", vbNullString, '
                      '"fx_charge", vbNullString, _\n', body)
        self.assertIn('TEST_Harness.AssertEqualText "fx-position-status-of-t9-jpy", "OK", _\n'
                      "            PlainText(RowCell(NM_FX_CALC_HEADER, 1, 1))\n", body)
        inputs = self.bodies["InputsMar4061Table9"]
        self.assertIn('SetNamed "ReportingCcy", "CHF"\n', inputs)
        self.assertIn('SetNamed "AsOfDate", DateSerial(2026, 9, 30)\n', inputs)
        self.assertIn('PutRate 7, "XAU", 1\n', inputs)
        self.assertIn('PutRow NM_FX_POS_HEADER, 6, Array("T9-XAU", Empty, "XAU", "Short", 35)\n',
                      inputs)
        self.assertEqual(len(fixture["datasets"]["DS-09"]), inputs.count("PutRow "))
        self.assertEqual(len(expected["outputs"]), 14)

    def test_a_zero_may_be_an_empty_cell_and_no_other_number_may(self) -> None:
        for fixture, expected in self.case_suite.loaded():
            lines, _, _ = self.case_suite.case_lines(fixture, expected)
            joined = " ".join(line.strip().rstrip("_").strip() for line in lines)
            wanted = [float(output["value"]) for output in expected["outputs"]
                      if "value" in output]
            readings = re.findall(r"Actual = (\w+)\(", joined)
            self.assertEqual(len(readings), len(wanted), fixture["id"])
            for value, function in zip(wanted, readings, strict=True):
                if function != "ZonesTotal":  # A sum, which is a number already.
                    self.assertEqual(function, "ZeroIfEmpty" if value == 0 else "NumberOf")

    def test_names_that_would_be_equal_get_the_position_of_their_output(self) -> None:
        outputs: list[dict[str, Any]] = [
            {"quantity": "ir_position_status", "scope": {"record": record}}
            for record in ("X-DUP", "x-dup", "X-CCY")]
        outputs.append({"quantity": "fx_charge", "scope": None})
        self.assertEqual(
            self.case_suite.assertion_names({"fixture": "a-case", "outputs": outputs}),
            ["ir-position-status-of-x-dup-output-1", "ir-position-status-of-x-dup-output-2",
             "ir-position-status-of-x-ccy", "fx-charge"])

    def test_a_number_is_written_as_the_editor_keeps_it(self) -> None:
        literal = self.case_suite.number_literal
        for value, text in ((0, "0"), (335.0, "335"), (-35, "-35"), (26.8, "26.8"),
                            (48267441, "48267441"), (603343012.5, "603343012.5"),
                            (3000000000, "3000000000#"), (1e-06, "0.000001"),
                            (1e-09, "0.000000001"), (0.1 + 0.2, "0.3"),
                            (10.499999999999998, "10.5"), (-0.045, "-0.045")):
            with self.subTest(value=value):
                self.assertEqual(literal(value), text)
        with self.assertRaises(self.case_suite.SuiteError):
            literal(float("nan"))

    def test_a_cell_value_is_typed_as_the_python_runner_types_it(self) -> None:
        literal = self.case_suite.cell_literal
        self.assertEqual(literal("final_maturity_date", "2031-10-01"), "DateSerial(2031, 10, 1)")
        # A date that is none stays the text it is, and so does a date under another key.
        self.assertEqual(literal("final_maturity_date", "2031-13-01"), '"2031-13-01"')
        self.assertEqual(literal("issue_id", "2031-10-01"), '"2031-10-01"')
        self.assertEqual(literal("issue_id", 'a "quoted" id'), '"a ""quoted"" id"')
        self.assertEqual(literal("coupon_rate", None), "Empty")
        self.assertEqual(literal("coupon_rate", 0.04), "0.04")
        self.assertEqual(literal("coupon_rate", True), "True")
        with self.assertRaises(self.case_suite.SuiteError):
            literal("coupon_rate", [1])

    def test_a_text_is_typed_as_a_text(self) -> None:
        # Assigned as it is, "1000" or "00123" would become a number, and "2029-06-26" a
        # date. Every write of a test case goes through PutValue.
        self.assertIn("PutValue Target.Offset(0, c - LBound(Values)), Values(c)",
                      self.bodies["PutRow"])
        self.assertIn("PutValue ThisWorkbook.Names(NM_FX_CCY).RefersToRange.Cells(Index, 1), Code",
                      self.bodies["PutRate"])
        self.assertIn("PutValue ThisWorkbook.Names(NM_FX_RATE).RefersToRange.Cells(Index, 1), "
                      "Rate", self.bodies["PutRate"])
        self.assertIn("PutValue ThisWorkbook.Names(CellName).RefersToRange, Value",
                      self.bodies["SetNamed"])
        for name in ("PutRow", "PutRate", "SetNamed"):
            self.assertNotIn(".Value =", self.bodies[name], name)
        put = re.search(r"^Private Sub PutValue\(.*?^End Sub$", self.text, re.M | re.S)
        assert put is not None
        # The cell is formatted as text for the write and gets its format back, on the
        # error path too; an apostrophe typed before the text would stay in a cell that is
        # formatted as text already, as the template's code columns are.
        steps = [line.strip() for line in put.group(0).split("\n")]
        # The write of a text that needs the format follows the format; an earlier
        # "Target.Value = Value" is the write into a cell that is formatted as text already.
        order = [steps.index(step, steps.index("On Error GoTo Failed")) for step in (
            'Target.NumberFormat = "@"', "Target.Value = Value", "Restored:",
            "On Error GoTo 0", "Target.NumberFormat = Saved",
            "If ErrNumber <> 0 Then Err.Raise ErrNumber, ErrSource, ErrDescription",
            "Failed:", "Resume Restored")]
        self.assertEqual(order, sorted(order))
        self.assertNotIn("'\" &", put.group(0))
        self.assertNotIn("Typed", self.text)
        # The cases hold such a text: a number typed as a text.
        self.assertTrue(any(re.search(r', "\d[\d.]*"[,)]', body)
                            for name, body in self.bodies.items() if name.startswith("Inputs")))

    def test_a_long_statement_is_broken_after_a_comma_outside_a_text(self) -> None:
        statement = ('PutRow NAME, 1, Array("a text, with a comma, that is long enough to pass '
                     'the width of the line", 1, 2, 3, "another, one", 4, 5, 6, 7, 8, 9, 10)')
        lines = self.case_suite.wrapped(statement)
        self.assertGreater(len(lines), 1)
        self.assertEqual(" ".join(line.strip().removesuffix(" _") for line in lines), statement)
        for line in lines[:-1]:
            self.assertTrue(line.endswith(", _"))
            self.assertEqual(line.count('"') % 2, 0)
        for line in self.text.split("\n"):
            self.assertLessEqual(len(line), 100, line)
        with self.assertRaises(self.case_suite.SuiteError):
            self.case_suite.wrapped('X "' + "a" * 200 + '"')

    def test_each_case_keeps_the_inputs_and_restores_the_sheets_on_both_paths(self) -> None:
        cases = {name: body for name, body in self.bodies.items() if name.startswith("Case")}
        self.assertEqual(len(cases), len(self.case_suite.loaded()))
        for name, body in cases.items():
            with self.subTest(case=name):
                statements = [line.strip() for line in body.split("\n")]
                order = [statements.index(step) for step in (
                    "On Error GoTo CaseFailed", "KeepInputs", "RestoreDue = True",
                    f"Inputs{name[4:]}", "ResultText = RunSSA_Silent()", "CleanUp:",
                    "On Error GoTo 0", "If RestoreDue Then", "Detail = RestoreInputs()",
                    'TEST_Harness.AssertTrue "inputs-restored", Len(Detail) = 0, Detail',
                    "Exit Sub", "CaseFailed:", "TEST_Harness.RecordCaseError",
                    "Resume CleanUp")]
                self.assertEqual(order, sorted(order))
                # One statement asserts the restoration, whichever path reaches it.
                self.assertEqual(body.count('"inputs-restored"'), 1)
        restore = re.search(r"^Private Function RestoreInputs\(.*?^End Function$", self.text,
                            re.M | re.S)
        assert restore is not None
        # It writes the inputs back, checks them, clears the outputs and keeps its error.
        for step in ("On Error GoTo Failed", "InputArea(i).Value = mKept(i)",
                     "Differences = InputDifferences()", "Cleared = ClearOutputs_Silent()"):
            self.assertIn(step, restore.group(0))
        # A formula an input cell held is written back, and the check compares it: a
        # run that left its result as a constant would otherwise pass.
        self.assertIn("InputArea(i).Cells(r, c).Formula = mFormulas(i)(r, c)",
                      restore.group(0))
        self.assertLess(restore.group(0).index("InputArea(i).Value = mKept(i)"),
                        restore.group(0).index(".Formula = mFormulas(i)(r, c)"))
        for name, step in (("KeepInputs", "mFormulas(i) = AsTable(InputArea(i).Formula)"),
                           ("InputDifferences",
                            "CurrentFormulas = AsTable(InputArea(i).Formula)")):
            kept = re.search(rf"^Private (?:Sub|Function) {name}\(.*?^End (?:Sub|Function)$",
                             self.text, re.M | re.S)
            assert kept is not None
            self.assertIn(step, kept.group(0))
        differences = re.search(r"^Private Function InputDifferences\(.*?^End Function$",
                                self.text, re.M | re.S)
        assert differences is not None
        self.assertIn("IsFormula(CurrentFormulas(r, c)) <> IsFormula(mFormulas(i)(r, c))",
                      differences.group(0))
        self.assertIn("CurrentFormulas(r, c) <> mFormulas(i)(r, c)", differences.group(0))
        for name, body in self.bodies.items():
            if name.startswith("Inputs"):
                self.assertRegex(body, r"\n    '[^\n]*\n\s+ClearInputs\n", name)

    def test_a_text_is_put_back_as_a_text_and_every_row_of_a_table_is_kept(self) -> None:
        restore = re.search(r"^Private Function RestoreInputs\(.*?^End Function$", self.text,
                            re.M | re.S)
        assert restore is not None
        # After the values, a kept text that is not a formula is written again as a text.
        self.assertRegex(restore.group(0), r"If VarType\(AsTable\(mKept\(i\)\)\(r, c\)\) = vbString "
                         r"Then\n\s+If Not IsFormula\(mFormulas\(i\)\(r, c\)\) Then\n\s+"
                         r"PutValue InputArea\(i\)\.Cells\(r, c\), AsTable\(mKept\(i\)\)\(r, c\)")
        # A position sheet is kept, cleared and restored through the rows of its table,
        # as M_Main reads it, and not through a fixed number of rows.
        keep = re.search(r"^Private Sub KeepInputs\(.*?^End Sub$", self.text, re.M | re.S)
        rows = re.search(r"^Private Function TableRows\(.*?^End Function$", self.text,
                         re.M | re.S)
        area = re.search(r"^Private Function InputArea\(.*?^End Function$", self.text,
                         re.M | re.S)
        assert keep is not None and rows is not None and area is not None
        self.assertIn("mRows(i) = TableRows(i)", keep.group(0))
        self.assertIn("Sheet.Cells(Sheet.Rows.Count, Header.Column).End(xlUp).Row - Header.Row",
                      rows.group(0))
        self.assertIn("If Last < INPUT_ROWS Then Last = INPUT_ROWS", rows.group(0))
        for index, width in enumerate(("IR", "EQ", "CO", "FX"), 1):
            self.assertIn(f"Resize(PositionRows({index}), {width}_INPUT_WIDTH)", area.group(0))
        self.assertNotIn("Resize(INPUT_ROWS", area.group(0))
        # The same rule as M_Main.TableValues: the last row that holds a PositionID.
        main = (ROOT / "src" / "workbook" / "M_Main.bas").read_bytes().decode("cp1252")
        self.assertIn("LastRow = Sheet.Cells(Sheet.Rows.Count, Header.Column).End(xlUp).Row", main)

    def test_the_cells_read_are_those_the_template_labels(self) -> None:
        assert openpyxl is not None
        book = openpyxl.load_workbook(ROOT / "src" / "workbook" / "Template.xlsx")

        def label(name: str, offset: int) -> str:
            sheet, _, cell = book.defined_names[name].attr_text.rpartition("!")
            header = book[sheet.strip("'")][cell.replace("$", "")]
            return str(header.offset(0, offset).value)

        first = self.run_cases.FIRST_COLUMN
        zones = {quantity: [label("IR_Zones_Header", column - first) for column in columns]
                 for quantity, columns in self.run_cases.ZONES_OUTPUTS.items()}
        self.assertEqual(zones, {
            "ir_vertical_disallowance": ["Vertical disallowance"],
            "ir_horizontal_disallowance_within_zones":
                ["Z1 charge (40%)", "Z2 charge (30%)", "Z3 charge (30%)"],
            "ir_horizontal_disallowance_between_zones": ["Between-zone charge"],
            "ir_net_position_charge": ["Net position charge (100%)"],
            "ir_general_market_risk_charge": ["General market risk"],
            "ir_specific_risk_charge": ["Specific risk"]})
        results = {quantity: label("Results_Header", column - first)
                   for quantity, (_, column) in self.run_cases.TOTALS.items()}
        self.assertEqual({results[name] for name in ("ir_specific_risk_charge",
                                                     "eq_specific_risk_charge")},
                         {"Specific risk"})
        self.assertEqual({results[name] for name in (
            "ir_general_market_risk_charge", "eq_general_market_risk_charge",
            "co_directional_charge", "fx_charge")}, {"General market risk"})
        self.assertEqual({results[name] for name in ("eq_index_charge", "co_gross_charge")},
                         {"Other charge"})
        self.assertEqual({quantity: label("Results_Header", column - first)
                          for quantity, column in self.run_cases.BY_CLASS.items()},
                         {"option_charge": "Non-delta options",
                          "requirement_before_scaling": "Requirement before scaling",
                          "requirement_after_scaling": "Requirement after scaling"})
        self.assertEqual((label("Buckets_FX_Header", self.run_cases.FX_BLOCK_CODE - first),
                          label("Buckets_FX_Header", self.run_cases.FX_BLOCK_NET - first)),
                         ("Currency", "Net position (rep ccy)"))
        sheet, _, cell = book.defined_names["Results_Header"].attr_text.rpartition("!")
        header = book[sheet][cell.replace("$", "")]
        self.assertEqual([str(header.offset(row, 0).value) for row in range(1, 5)],
                         ["Interest rate", "Equity", "Commodity", "Foreign exchange"])
        self.assertEqual(header.row, self.run_cases.RESULTS_ROW["IR"] - 1)
        # The module reads the FX charge four rows under the header, three columns
        # to the right of the risk class.
        self.assertIn("ResultsCell(4, 3)", self.bodies["CaseIllustrativeSampleFx"])


if __name__ == "__main__":
    unittest.main()
