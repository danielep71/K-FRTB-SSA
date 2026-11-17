"""The project's test cases: the derivation script, the fixtures against the
sample data, and the tool that runs the cases on the engine."""

from __future__ import annotations

import contextlib
import io
import json
import sys
import tempfile
import unittest
import unittest.mock
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "prototype"))
sys.path.insert(0, str(ROOT / "tests" / "derivations"))

import ssa  # noqa: E402
import ssa_reference  # noqa: E402

try:
    import openpyxl
except ImportError:  # pragma: no cover - installed from .github/requirements.txt
    openpyxl = None

FIXTURES = ROOT / "tests" / "fixtures"
EXPECTED = ROOT / "tests" / "expected"
TEMPLATE = ROOT / "src" / "workbook" / "Template.xlsx"
PROJECT = frozenset(f"CALC-0{number}" for number in range(2, 9))


def fixture(case: str) -> dict[str, Any]:
    data: dict[str, Any] = json.loads((FIXTURES / f"{case}.json").read_text(encoding="utf-8"))
    return data


def derived(case: dict[str, Any]) -> dict[tuple[str, str | None], Any]:
    """The outputs of a fixture by quantity and record."""
    return {(output["quantity"], None if output["scope"] is None else output["scope"]["record"]):
            output.get("value", output.get("text")) for output in ssa.derive(case)}


def ir_case(**inputs: object) -> dict[str, Any]:
    """A fixture with one interest rate row, valid unless inputs change it."""
    row = {"instrument_code": "Bond", "currency_code": "EUR", "direction_code": "Long",
           "position_value_amount": 1000000, "rate_type_code": "Fixed", "coupon_rate": 0.05,
           "final_maturity_date": "2030-12-31", "next_repricing_date": None,
           "sr_category_code": "NONE", "rating_grade_code": None, "issue_id": None, **inputs}
    return {"calculation": "CALC-03", "reporting_date": "2026-12-31", "reporting_currency_code": "EUR",
            "datasets": {"DS-05": [{"id": "EUR", "inputs": {"fx_rate": 1.0}},
                                   {"id": "XAU", "inputs": {"fx_rate": 3500.0}}],
                         "DS-06": [{"id": "P1", "inputs": row}]}}


class Derivation(unittest.TestCase):
    """tests/derivations/ssa.py, on values that can be worked out by hand."""

    def test_the_example_of_table_9(self) -> None:
        outputs = derived(fixture("mar40-61-table-9"))
        self.assertEqual(outputs[("fx_overall_net_open_position", None)], 335)
        self.assertAlmostEqual(outputs[("fx_charge", None)], 26.8, places=9)

    def test_the_sample_requirement_is_that_of_the_independent_recalculation(self) -> None:
        outputs = derived(fixture("illustrative-sample-requirement"))
        reference = ssa_reference.calc_all()
        self.assertAlmostEqual(outputs[("own_funds_requirement", None)], reference["total"], places=6)
        self.assertAlmostEqual(outputs[("own_funds_requirement", None)], 48267441, places=6)
        self.assertAlmostEqual(outputs[("risk_weighted_assets", None)], 603343012.5, places=6)
        for name, key in (("IR", "k_ir"), ("EQ", "k_eq"), ("CO", "k_co"), ("FX", "k_fx")):
            with self.subTest(risk_class=name):
                self.assertAlmostEqual(outputs[("requirement_before_scaling", name)],
                                       reference[key], places=6)  # type: ignore[literal-required]

    def test_a_term_on_the_limit_of_a_band_belongs_to_it(self) -> None:
        for term, high, band in ((1.0, True, 4), (1.0001, True, 5), (4.0, True, 7), (4.0, False, 8),
                                 (20.0, True, 12), (20.0001, True, 13), (500.0, True, 13),
                                 (12.0, False, 13), (20.0, False, 14), (20.0001, False, 15),
                                 (0.0, True, 1), (1 / 12, False, 1)):
            with self.subTest(term=term, high=high):
                self.assertEqual(ssa.band_of(term, high), band)

    def test_the_status_is_the_first_test_that_fails(self) -> None:
        cases: tuple[tuple[dict[str, object], str], ...] = (
            ({}, "OK"),
            # A row that fails several tests reports the first of the contract's list.
            ({"currency_code": "ZAR", "direction_code": "Buy"}, "INVALID: currency not in FX_Table"),
            ({"currency_code": " eur "}, "OK"),
            ({"currency_code": "XAU"}, "INVALID: gold is FX risk; enter it on FXPositions"),
            ({"direction_code": " long "}, "OK"),
            ({"position_value_amount": True}, "INVALID: Position value must be a number"),
            ({"final_maturity_date": "2026-12-30"}, "INVALID: final maturity before AsOfDate"),
            ({"final_maturity_date": "2026-12-31"}, "OK"),
            ({"final_maturity_date": "31/12/2030"}, "INVALID: Final maturity date missing"),
            ({"final_maturity_date": 3000000}, "INVALID: Final maturity date is not a date"),
            ({"coupon_rate": None, "sr_category_code": "CORP"},
             "INVALID: Coupon required for a fixed-rate position"),
            ({"rate_type_code": "Floating", "coupon_rate": None, "next_repricing_date": "2026-12-31"},
             "OK"),
            ({"rate_type_code": "Floating", "next_repricing_date": "2031-01-01"},
             "INVALID: repricing date outside AsOfDate to final maturity"),
            ({"sr_category_code": "qual", "rating_grade_code": "a+ to bbb-"}, "OK"),
            ({"sr_category_code": "QUAL", "rating_grade_code": "BB+ to BB-"},
             "INVALID: no MAR40.6 charge for this category and rating grade"),
        )
        for inputs, status in cases:
            with self.subTest(inputs=inputs):
                self.assertEqual(derived(ir_case(**inputs))[("ir_position_status", "P1")], status)

    def test_an_entered_issue_never_nets_with_a_position_without_one(self) -> None:
        case = ir_case(sr_category_code="OTHER", rating_grade_code="BB+ to BB-")
        first = case["datasets"]["DS-06"][0]
        first["id"] = "R1"
        second = {"id": "P2", "inputs": {**first["inputs"], "direction_code": "Short",
                                         "issue_id": "POS:R1"}}
        case["datasets"]["DS-06"].append(second)
        case["calculation"] = "CALC-02"
        # Two issues of 1,000,000 at 8% each, not one issue that nets to zero.
        self.assertAlmostEqual(derived(case)[("ir_specific_risk_charge", None)], 160000, places=6)

    def test_an_unknown_calculation_is_refused(self) -> None:
        with self.assertRaises(SystemExit):
            ssa.derive({**ir_case(), "calculation": "CALC-99"})


class Registered(unittest.TestCase):
    """The fixtures and expected files of the project's cases."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.cases = {path.stem: json.loads(path.read_text(encoding="utf-8"))
                     for path in sorted(FIXTURES.glob("*.json"))}
        cls.cases = {name: case for name, case in cls.cases.items() if case["calculation"] in PROJECT}

    cases: dict[str, dict[str, Any]]

    def test_each_calculation_has_a_nominal_and_a_boundary_case(self) -> None:
        self.assertEqual(len(self.cases), 18)
        for calculation in sorted(PROJECT):
            categories = {case["category"] for case in self.cases.values()
                          if case["calculation"] == calculation}
            with self.subTest(calculation=calculation):
                self.assertIn("nominal", categories)
                # The option charges are taken as entered: they have no limit of their own.
                if calculation != "CALC-07":
                    self.assertIn("boundary", categories)
        invalid = {case["calculation"] for case in self.cases.values() if case["category"] == "invalid"}
        self.assertEqual(invalid, {"CALC-03", "CALC-04", "CALC-05", "CALC-06"})

    def test_the_sample_cases_hold_the_sample_data_of_the_template(self) -> None:
        datasets = self.cases["illustrative-sample-requirement"]["datasets"]
        self.assertEqual({record["id"]: record["inputs"]["fx_rate"] for record in datasets["DS-05"]},
                         {code: rate for code, rate, _ in ssa_reference.FX})
        for dataset, rows, key, column in (("DS-06", ssa_reference.IR, "position_value_amount", 5),
                                           ("DS-07", ssa_reference.EQ, "market_value_amount", 7),
                                           ("DS-08", ssa_reference.CO, "quantity_count", 4),
                                           ("DS-09", ssa_reference.FXPOS, "net_position_amount", 4)):
            with self.subTest(dataset=dataset):
                self.assertEqual([(record["id"], record["inputs"][key]) for record in datasets[dataset]],
                                 [(row[0], row[column]) for row in rows])
        self.assertEqual(self.cases["illustrative-sample-requirement"]["reporting_date"],
                         ssa_reference.AS_OF.isoformat())
        # The cases of one calculation hold the same records as the whole portfolio.
        for case, dataset in (("illustrative-sample-ir-general-market-risk", "DS-06"),
                              ("illustrative-sample-equity", "DS-07"),
                              ("illustrative-sample-commodities", "DS-08"),
                              ("illustrative-sample-fx", "DS-09")):
            with self.subTest(case=case):
                self.assertEqual(self.cases[case]["datasets"][dataset], datasets[dataset])

    def test_only_the_values_printed_in_the_source_are_published(self) -> None:
        published = []
        for name in self.cases:
            expected = json.loads((EXPECTED / f"{name}.json").read_text(encoding="utf-8"))
            for output in expected["outputs"]:
                reference = output["reference"]
                self.assertEqual(reference["script"], "tests/derivations/ssa.py")
                self.assertIn(reference["class"], ("published", "illustrative"))
                if reference["class"] == "published":
                    published.append((name, output["quantity"], output["value"], reference["source"],
                                      reference["locator"], output["tolerance"]["absolute"]))
        self.assertEqual(published, [
            ("mar40-61-table-9", "fx_overall_net_open_position", 335.0, "SRC-02", "MAR40.61, table 9",
             0.5),
            ("mar40-61-table-9", "fx_charge", 26.8, "SRC-02", "MAR40.61, table 9", 0.05)])


@unittest.skipIf(openpyxl is None, "needs openpyxl, which .github/requirements.txt pins")
class Runner(unittest.TestCase):
    """tools/prototype/run_cases.py without LibreOffice: what it writes and what it reads."""

    @classmethod
    def setUpClass(cls) -> None:
        import run_cases
        cls.tool = run_cases

    tool: Any

    def test_it_runs_the_cases_of_the_project_and_not_the_illustrative_examples(self) -> None:
        names = [case["id"] for case in self.tool.cases([])]
        self.assertEqual(len(names), 18)
        self.assertNotIn("illustrative-format-example", names)
        self.assertEqual([case["id"] for case in self.tool.cases(["mar40-61-table-9"])],
                         ["mar40-61-table-9"])

    def test_a_fixture_replaces_the_sample_data_of_the_workbook(self) -> None:
        assert openpyxl is not None
        workbook = openpyxl.load_workbook(TEMPLATE)
        case = fixture("illustrative-ir-time-band-limits")
        self.tool.write_inputs(workbook, case)
        verify = self.tool.verify
        self.assertEqual(verify.named_cell(workbook, "AsOfDate").value.isoformat(), "2026-12-31")
        codes = verify.first_cell(workbook, "FX_Ccy")
        self.assertEqual([codes.offset(offset).value for offset in range(3)], ["EUR", None, None])
        self.assertEqual(verify.first_cell(workbook, "FX_Rate").value, 1.0)
        positions = workbook["IRPositions"]
        records = case["datasets"]["DS-06"]
        self.assertEqual([positions.cell(5 + offset, 3).value for offset in range(len(records) + 1)],
                         [record["id"] for record in records] + [None])
        # Instrument, currency, value, coupon and maturity of the first row, in their columns.
        self.assertEqual([positions.cell(5, column).value for column in (5, 6, 8, 10)],
                         ["Bond", "EUR", 1000000, 0.05])
        self.assertEqual(positions.cell(5, 11).value.isoformat(), "2027-01-30")
        # The other position sheets hold no row of the sample.
        for sheet in ("EQPositions", "COPositions", "FXPositions"):
            self.assertIsNone(workbook[sheet].cell(5, 3).value)

    def test_a_value_that_is_not_what_a_cell_can_hold_is_written_as_it_is(self) -> None:
        self.assertEqual(self.tool.cell_value("final_maturity_date", 20311231), 20311231)
        self.assertEqual(self.tool.cell_value("final_maturity_date", "31/12/2030"), "31/12/2030")
        self.assertEqual(self.tool.cell_value("position_value_amount", "ten"), "ten")

    def test_an_output_is_read_from_its_cell(self) -> None:
        assert openpyxl is not None
        workbook = openpyxl.load_workbook(TEMPLATE)
        case = fixture("illustrative-sample-requirement")
        workbook["Results"]["K11"].value = 48267441.0
        workbook["Results"]["I8"].value = 2792000.0
        workbook["IRCalc"]["D6"].value = "OK"
        workbook["IR_Zones"]["C6"].value = "EUR"
        for cell, value in (("I6", 10.0), ("N6", 20.0), ("S6", 30.0)):
            workbook["IR_Zones"][cell].value = value

        def read(quantity: str, dataset: str | None = None, record: str | None = None) -> object:
            scope = None if dataset is None else {"dataset": dataset, "record": record}
            return self.tool.actual(workbook, case, {"quantity": quantity, "scope": scope})

        self.assertEqual(read("own_funds_requirement"), 48267441.0)
        self.assertEqual(read("requirement_before_scaling", "DS-10", "EQ"), 2792000.0)
        self.assertEqual(read("ir_position_status", "DS-06", case["datasets"]["DS-06"][1]["id"]), "OK")
        self.assertEqual(read("ir_horizontal_disallowance_within_zones", "DS-05", "EUR"), 60.0)
        self.assertEqual(read("ir_vertical_disallowance", "DS-05", "ZAR"), "no such currency on IR_Zones")

    def test_a_number_agrees_within_its_tolerance_and_a_text_exactly(self) -> None:
        number = {"value": 100.0, "tolerance": {"absolute": 0.5, "relative": 0}}
        self.assertTrue(self.tool.agrees(number, 100.5))
        self.assertFalse(self.tool.agrees(number, 100.6))
        self.assertFalse(self.tool.agrees(number, None))
        self.assertFalse(self.tool.agrees(number, "100"))
        self.assertTrue(self.tool.agrees({"value": 0, "tolerance": {"absolute": 0, "relative": 0}}, None))
        self.assertTrue(self.tool.agrees({"text": "OK"}, "OK"))
        self.assertFalse(self.tool.agrees({"text": "OK"}, "ok"))

    def test_a_difference_names_the_quantity_the_record_and_both_values(self) -> None:
        with unittest.mock.patch.object(self.tool, "run_case",
                                        return_value=["fx_charge: expected 26.8, the workbook has 27"]):
            count, compared, report = self.tool.check(unittest.mock.Mock(), ["mar40-61-table-9"])
        self.assertEqual((count, compared), (1, 14))
        self.assertEqual(report, ["mar40-61-table-9: 14 output(s), 1 differ",
                                  "  fx_charge: expected 26.8, the workbook has 27"])

    def test_without_libreoffice_main_exits_2(self) -> None:
        with unittest.mock.patch.object(self.tool.shutil, "which", return_value=None), \
                contextlib.redirect_stdout(io.StringIO()) as output, \
                tempfile.TemporaryDirectory():
            self.assertEqual(self.tool.main([]), 2)
        self.assertIn("SKIP", output.getvalue())


if __name__ == "__main__":
    unittest.main()
