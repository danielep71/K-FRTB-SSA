"""The project's test cases: the derivation script, the fixtures against the
sample data, and the tool that runs the cases on the engine."""

from __future__ import annotations

import calendar
import contextlib
import datetime
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


INSTRUMENT = ("INVALID: Instrument must be Bond, FRN, Swap leg, Future leg, FRA leg, Forward leg "
              "or Other")


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
            # An instrument is one of the seven values of DS-06, whatever its case and blanks;
            # the test comes before the currency's.
            ({"instrument_code": " swap LEG "}, "OK"),
            ({"instrument_code": "Swap  leg"}, INSTRUMENT),
            ({"instrument_code": "Securitisation"}, INSTRUMENT),
            ({"instrument_code": None}, INSTRUMENT),
            ({"instrument_code": "Equity", "currency_code": "ZAR"}, INSTRUMENT),
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

    def test_a_commodity_that_is_gold_is_rejected_before_its_other_fields_are_read(self) -> None:
        records = [{"id": f"C{number}", "inputs": {
            "commodity_code": commodity, "direction_code": "Sell", "quantity_count": 1,
            "spot_price_amount": 1, "currency_code": "USD"}}
            for number, commodity in enumerate((" xau ", "Copper", None))]
        case = {"calculation": "CALC-05", "reporting_date": "2026-12-31",
                "reporting_currency_code": "EUR",
                "datasets": {"DS-05": [{"id": "EUR", "inputs": {"fx_rate": 1.0}},
                                       {"id": "USD", "inputs": {"fx_rate": 0.9}}],
                             "DS-08": records}}
        outputs = derived(case)
        self.assertEqual(outputs[("co_position_status", "C0")],
                         "INVALID: gold is FX risk; enter it on FXPositions")
        self.assertEqual(outputs[("co_position_status", "C1")],
                         "INVALID: Direction must be Long or Short")
        self.assertEqual(outputs[("co_position_status", "C2")], "INVALID: Commodity missing")

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


def month_ends(count: int, year: int = 2025, month: int = 1) -> list[datetime.date]:
    """count month ends, the first in the month of year and month."""
    ends = []
    for _ in range(count):
        ends.append(datetime.date(year, month, calendar.monthrange(year, month)[1]))
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    return ends


def sheet_rows(pattern: str, total: float = 8_000_000_000, **sizes: float) -> list[tuple[Any, ...]]:
    """One row per letter of the pattern: N, a month that misses both conditions, or M, one that
    meets them. The size of Article 325a is 600 million in an N and 300 million in an M, split
    60% long and 40% short with a minus sign; that of Article 94 is 60 million and 30 million,
    split 70% and 30%."""
    rows = []
    for end, flag in zip(month_ends(len(pattern)), pattern, strict=True):
        size_325a, size_94 = (600_000_000, 60_000_000) if flag == "N" else (300_000_000, 30_000_000)
        rows.append((end, total, 0.6 * size_325a, -0.4 * size_325a, 0.7 * size_94, 0.3 * size_94))
    return rows


def at_last(result: dict[str, Any], approach: str = "325a") -> tuple[Any, ...]:
    """The status, run, count and restart condition of an approach."""
    return (result["status"][approach], result["run"][approach], result["count"][approach],
            result["restart"][approach])


class Eligibility(unittest.TestCase):
    """tests/derivations/ssa.py, CALC-09: the eligibility under the CRR."""

    def test_the_size_at_the_share_or_at_the_amount_meets_the_conditions_and_one_above_does_not(self) -> None:
        end = datetime.date(2026, 1, 31)
        # With total assets of 1 billion, 10% is 100 million, below the amount of 500 million,
        # and 5% is 50 million, which is the amount; with 6 billion, the amounts are the limits.
        for total, size_325a, size_94, status_325a, status_94 in (
                (1_000_000_000, 100_000_000, 50_000_000, "MET", "MET"),
                (1_000_000_000, 100_000_001, 50_000_001, "NOT MET", "NOT MET"),
                (6_000_000_000, 500_000_000, 50_000_000, "MET", "MET"),
                (6_000_000_000, 500_000_001, 50_000_001, "NOT MET", "NOT MET"),
                (6_000_000_000, 500_000_000, 50_000_001, "MET", "NOT MET")):
            with self.subTest(total=total, size_325a=size_325a, size_94=size_94):
                result = ssa.eligibility([(end, total, size_325a, 0, size_94, 0)], "CRR", {})
                self.assertEqual(result["status"], {"325a": status_325a, "94": status_94})
                self.assertEqual(result["size"], {"325a": size_325a, "94": size_94})
                self.assertEqual(result["share"], {"325a": size_325a / total, "94": size_94 / total})

    def test_the_size_adds_the_absolute_values_of_the_long_and_the_short_total(self) -> None:
        for long, short in ((300, -200), (300, 200), (-300, -200), (-300, 200), (500, 0), (0, -500)):
            with self.subTest(long=long, short=short):
                result = ssa.eligibility([(datetime.date(2026, 1, 31), 1e12, long * 1e6, short * 1e6, 0, 0)],
                                         "CRR", {})
                self.assertEqual(result["size"]["325a"], abs(long) * 1e6 + abs(short) * 1e6)

    def test_a_parameter_that_is_given_replaces_its_reference_value(self) -> None:
        rows = [(datetime.date(2026, 1, 31), 1_000_000_000, 150_000_000, 0, 60_000_000, 0)]
        self.assertEqual(ssa.eligibility(rows, "CRR", {})["status"], {"325a": "NOT MET", "94": "NOT MET"})
        given = {"EligibilityShare325a": 0.2, "EligibilityCap325a": 150_000_000,
                 "EligibilityShare94": 0.1, "EligibilityCap94": 60_000_000}
        self.assertEqual(ssa.eligibility(rows, "CRR", given)["status"], {"325a": "MET", "94": "MET"})
        self.assertEqual(ssa.eligibility(rows, "CRR", {"EligibilityCap325a": 149_999_999})["status"]["325a"],
                         "NOT MET")

    def test_a_run_of_two_months_is_not_met_and_a_run_of_three_ceases(self) -> None:
        for pattern, expected in (("MMNN", ("NOT MET", 2, 2, "NO")), ("MMNNN", ("CEASE", 3, 3, "NO")),
                                  ("NNNM", ("MET", 0, 3, "NO")), ("MNNMN", ("NOT MET", 1, 3, "NO")),
                                  ("NNNNN", ("CEASE", 5, 5, "NO")), ("M", ("MET", 0, 0, "NO"))):
            with self.subTest(pattern=pattern):
                self.assertEqual(at_last(ssa.eligibility(sheet_rows(pattern), "CRR", {})), expected)

    def test_a_run_of_three_stays_visible_in_its_row_after_the_months_that_meet_the_conditions(self) -> None:
        result = ssa.eligibility(sheet_rows("MNNNMM"), "CRR", {})
        events = [item["325a"]["event"] for item in result["rows"]]
        self.assertEqual(events, ["NO", "NO", "NO", "YES", "NO", "NO"])
        self.assertEqual(result["status"]["325a"], "MET")

    def test_a_count_of_seven_in_twelve_months_ceases_and_a_count_of_six_does_not(self) -> None:
        for pattern, expected in (("NMNMNMNMNMNM", ("MET", 0, 6, "NO")),
                                  ("NMNMNMNMNMMN", ("NOT MET", 1, 6, "NO")),
                                  ("NMNMNMNMNMNN", ("CEASE", 2, 7, "NO")),
                                  ("NNMNNMNMNMNM", ("CEASE", 0, 7, "NO")),
                                  ("NNMNNMNMNMNMM", ("MET", 0, 6, "NO"))):
            with self.subTest(pattern=pattern):
                self.assertEqual(at_last(ssa.eligibility(sheet_rows(pattern), "CRR", {})), expected)

    def test_the_count_goes_over_the_last_twelve_months_and_over_all_of_them_before_the_twelfth(self) -> None:
        # Seven months that miss the conditions in a row of seven, then five that meet them.
        self.assertEqual(ssa.eligibility(sheet_rows("NMNMNMN"), "CRR", {})["count"]["325a"], 4)
        counts = [item["325a"]["count"] for item in ssa.eligibility(sheet_rows("NNMMMMMMMMMMMMM"), "CRR", {})["rows"]]
        # The first month leaves the window at the 13th row.
        self.assertEqual(counts, [1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 0, 0])

    def test_the_restart_condition_needs_twelve_months_that_meet_the_conditions(self) -> None:
        for pattern, restart in (("M" * 11, "NO"), ("M" * 12, "YES"), ("N" + "M" * 11, "NO"),
                                 ("N" + "M" * 12, "YES"), ("M" * 11 + "N", "NO"), ("M" * 30, "YES")):
            with self.subTest(pattern=pattern):
                # The sheet holds 24 rows, but the derivation takes as many as it is given.
                self.assertEqual(ssa.eligibility(sheet_rows(pattern), "CRR", {})["restart"],
                                 {"325a": restart, "94": restart})

    def test_the_two_approaches_are_assessed_apart(self) -> None:
        rows = [(end, 8e9, 300e6, 0, 60e6 if flag == "N" else 30e6, 0)
                for end, flag in zip(month_ends(3), "MNN", strict=True)]
        result = ssa.eligibility(rows, "CRR", {})
        self.assertEqual(at_last(result, "325a"), ("MET", 0, 0, "NO"))
        self.assertEqual(at_last(result, "94"), ("NOT MET", 2, 2, "NO"))
        self.assertEqual(result["last_month"], datetime.date(2025, 3, 31))

    def test_nothing_is_assessed_under_the_basel_regime_and_the_sheet_is_not_read(self) -> None:
        rows = [*sheet_rows("MMM"), ("not a date", 0, None, None, None, None)]
        for regime in ("Basel", "BASEL", "", None, "Other"):
            with self.subTest(regime=regime):
                result = ssa.eligibility(rows, regime, {})
                reading = "NOT ASSESSED: Basel regime"
                self.assertEqual(result["status"], {"325a": reading, "94": reading})
                self.assertEqual(result["rows"], [None] * 4)
                self.assertIsNone(result["last_month"])
                for key in ("restart", "run", "count", "size", "share"):
                    self.assertEqual(result[key], {"325a": None, "94": None})

    def test_the_regime_is_read_without_regard_to_case_and_blanks(self) -> None:
        for regime in ("CRR", "crr", " Crr "):
            with self.subTest(regime=regime):
                self.assertEqual(ssa.eligibility(sheet_rows("M"), regime, {})["status"],
                                 {"325a": "MET", "94": "MET"})

    def test_no_month_entered_is_not_assessed(self) -> None:
        for rows in ([], [(None, None, None, None, None, None)], [(None, 8e9, 1, 2, 3, 4), ("", 8e9, 1, 2, 3, 4)]):
            with self.subTest(rows=rows):
                result = ssa.eligibility(rows, "CRR", {})
                reading = "NOT ASSESSED: no month entered"
                self.assertEqual(result["status"], {"325a": reading, "94": reading})
                self.assertEqual(result["rows"], [None] * len(rows))
                self.assertIsNone(result["last_month"])

    def test_each_reason_that_rejects_a_month_and_its_order(self) -> None:
        end = datetime.date(2026, 1, 31)
        cases: tuple[tuple[tuple[Any, ...], str], ...] = (
            (("abc", 1e9, 1, 1, 1, 1), "INVALID: MonthEnd must be a date"),
            ((True, 1e9, 1, 1, 1, 1), "INVALID: MonthEnd must be a date"),
            ((0, 1e9, 1, 1, 1, 1), "INVALID: MonthEnd must be a date"),
            ((3_000_000, 1e9, 1, 1, 1, 1), "INVALID: MonthEnd must be a date"),
            ((datetime.date(2026, 1, 30), 1e9, 1, 1, 1, 1), "INVALID: MonthEnd must be the last day of a month"),
            ((datetime.date(2026, 2, 27), 1e9, 1, 1, 1, 1), "INVALID: MonthEnd must be the last day of a month"),
            ((46053.5, 1e9, 1, 1, 1, 1), "INVALID: MonthEnd must be the last day of a month"),
            ((end, 0, 1, 1, 1, 1), "INVALID: TotalAssets must be a number greater than 0"),
            ((end, -5, 1, 1, 1, 1), "INVALID: TotalAssets must be a number greater than 0"),
            ((end, None, 1, 1, 1, 1), "INVALID: TotalAssets must be a number greater than 0"),
            ((end, "1e9", 1, 1, 1, 1), "INVALID: TotalAssets must be a number greater than 0"),
            ((end, True, 1, 1, 1, 1), "INVALID: TotalAssets must be a number greater than 0"),
            ((end, 1e9, None, 1, 1, 1), "INVALID: a long or short total must be a number"),
            ((end, 1e9, 1, "x", 1, 1), "INVALID: a long or short total must be a number"),
            ((end, 1e9, 1, 1, False, 1), "INVALID: a long or short total must be a number"),
            ((end, 1e9, 1, 1, 1, None), "INVALID: a long or short total must be a number"),
            # A row that fails several tests gives the first reason of the list.
            ((datetime.date(2026, 1, 30), 0, None, 1, 1, 1), "INVALID: MonthEnd must be the last day of a month"),
            (("abc", 0, None, 1, 1, 1), "INVALID: MonthEnd must be a date"),
            ((end, 0, None, 1, 1, 1), "INVALID: TotalAssets must be a number greater than 0"))
        for values, reason in cases:
            with self.subTest(values=values):
                result = ssa.eligibility([values], "CRR", {})
                self.assertEqual(result["rows"][0], {"status": reason, "325a": None, "94": None})
                reading = "NOT ASSESSED: a row is invalid"
                self.assertEqual(result["status"], {"325a": reading, "94": reading})
                self.assertIsNone(result["last_month"])
                self.assertEqual(result["restart"], {"325a": None, "94": None})

    def test_a_month_follows_the_row_entered_before_it(self) -> None:
        first, second, third = (datetime.date(2026, 1, 31), datetime.date(2026, 2, 28), datetime.date(2026, 3, 31))
        sequence = "INVALID: MonthEnd must be one month after the previous row"

        def status(*ends: Any) -> list[str]:
            rows = [(end, 8e9, 1, 1, 1, 1) for end in ends]
            return [item["status"] for item in ssa.eligibility(rows, "CRR", {})["rows"]]

        self.assertEqual(status(first, second, third), ["OK", "OK", "OK"])
        self.assertEqual(status(first, third), ["OK", sequence])
        self.assertEqual(status(second, first), ["OK", sequence])
        self.assertEqual(status(first, first), ["OK", sequence])
        self.assertEqual(status(datetime.date(2025, 12, 31), first), ["OK", "OK"])
        self.assertEqual(status(datetime.date(2024, 2, 29), datetime.date(2024, 3, 31)), ["OK", "OK"])
        # The first row has no previous row. A row after one that is not a date is not tested.
        self.assertEqual(status("abc", third), ["INVALID: MonthEnd must be a date", "OK"])
        # A previous row that is a date but not a month end is the reference all the same.
        self.assertEqual(status(datetime.date(2026, 1, 30), second),
                         ["INVALID: MonthEnd must be the last day of a month", "OK"])
        # The last day of 9999 has no month after it.
        self.assertEqual(status(2958465, 2958496), ["OK", "INVALID: MonthEnd must be a date"])

    def test_a_row_without_a_month_end_is_skipped_and_the_row_entered_before_it_is_the_reference(self) -> None:
        full = sheet_rows("MMNNN")
        rows: list[tuple[Any, ...]] = [full[0], (None, 8e9, 1, 1, 1, 1), full[1], ("", None, None, None, None, None),
                                       (None, None, None, None, None, None), full[2], full[3], (None, 1, 2, 3, 4, 5), full[4]]
        result = ssa.eligibility(rows, "CRR", {})
        self.assertEqual([item is None for item in result["rows"]], [False, True, False, True, True, False, False, True, False])
        # The run and the count are those of the rows entered, as if the blank rows were not there.
        dense = ssa.eligibility(full, "CRR", {})
        self.assertEqual([item for item in result["rows"] if item is not None], dense["rows"])
        self.assertEqual({key: result[key] for key in ("status", "run", "count", "restart", "last_month")},
                         {key: dense[key] for key in ("status", "run", "count", "restart", "last_month")})
        self.assertEqual(at_last(result), ("CEASE", 3, 3, "NO"))

    def test_a_rejected_row_leaves_the_others_their_figures_and_stops_the_assessment(self) -> None:
        rows = sheet_rows("NNMN")
        rows[1] = (rows[1][0], 0, *rows[1][2:])
        result = ssa.eligibility(rows, "CRR", {})
        self.assertEqual([item["status"] for item in result["rows"]],
                         ["OK", "INVALID: TotalAssets must be a number greater than 0", "OK", "OK"])
        self.assertIsNone(result["rows"][1]["325a"])
        # The run and the count go over the valid rows, in the order of the sheet: the rejected month is passed over.
        self.assertEqual([(item["325a"]["run"], item["325a"]["count"]) for item in result["rows"] if item["325a"]],
                         [(1, 1), (0, 1), (1, 2)])
        self.assertEqual(result["status"], {"325a": "NOT ASSESSED: a row is invalid",
                                            "94": "NOT ASSESSED: a row is invalid"})

    def test_a_month_end_is_a_date_a_datetime_at_midnight_or_a_serial_number(self) -> None:
        january = datetime.date(2026, 1, 31)
        serial = (january - datetime.date(1899, 12, 30)).days
        for value in (january, datetime.datetime(2026, 1, 31), serial, float(serial)):
            with self.subTest(value=value):
                result = ssa.eligibility([(value, 8e9, 1, 1, 1, 1)], "CRR", {})
                self.assertEqual(result["rows"][0]["status"], "OK")
                self.assertEqual(result["last_month"], value)
        late = ssa.eligibility([(datetime.datetime(2026, 1, 31, 12), 8e9, 1, 1, 1, 1)], "CRR", {})
        self.assertEqual(late["rows"][0]["status"], "INVALID: MonthEnd must be the last day of a month")

    def test_the_sample_series_of_the_workbook(self) -> None:
        rows = ssa_reference.ELIGIBILITY
        self.assertEqual(len(rows), 24)
        self.assertEqual(rows[-1][0], ssa_reference.AS_OF)
        result = ssa.eligibility(rows, "CRR", {})
        self.assertEqual([item["status"] for item in result["rows"]], ["OK"] * 24)
        # The series meets and misses the conditions of Article 325a, and has a month at the amount.
        self.assertEqual({item["325a"]["result"] for item in result["rows"]}, {"MET", "NOT MET"})
        self.assertEqual(result["rows"][18]["325a"]["size"], 500_000_000)
        self.assertEqual(result["rows"][18]["325a"]["result"], "MET")
        self.assertIn("YES", {item["325a"]["event"] for item in result["rows"]})
        self.assertEqual(result["status"], {"325a": "NOT MET", "94": "MET"})
        self.assertEqual(result["restart"], {"325a": "NO", "94": "YES"})
        self.assertEqual(result["last_month"], ssa_reference.AS_OF)
        self.assertEqual(ssa.eligibility(rows, "Basel", {})["status"]["325a"], "NOT ASSESSED: Basel regime")


class Registered(unittest.TestCase):
    """The fixtures and expected files of the project's cases."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.cases = {path.stem: json.loads(path.read_text(encoding="utf-8"))
                     for path in sorted(FIXTURES.glob("*.json"))}
        cls.cases = {name: case for name, case in cls.cases.items() if case["calculation"] in PROJECT}

    cases: dict[str, dict[str, Any]]

    def test_each_calculation_has_a_nominal_and_a_boundary_case(self) -> None:
        self.assertEqual(len(self.cases), 20)
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
        # What each source prints: table 10 and paragraphs 2 and 3 of example C.2 of SRC-03,
        # whose positions are printed to 0.01 mn (tolerance 500); MAR40.27, which prints whole
        # millions (tolerance 500,000); table 9 of MAR40.61.
        c2 = "bcbs119-c2-maturity-method"
        v27 = "mar40-27-vertical-disallowance"
        self.assertEqual(published, [
            (c2, "ir_weighted_position", 500000.0, "SRC-03", "Part C.2, table 10, 7 to 10 years", 500),
            (c2, "ir_weighted_position", 150000.0, "SRC-03", "Part C.2, table 10, 1 to 3 months", 500),
            (c2, "ir_weighted_position", 1050000.0, "SRC-03", "Part C.2, table 10, 6 to 12 months", 500),
            (c2, "ir_weighted_position", -5625000.0, "SRC-03", "Part C.2, table 10, 7 to 10 years", 500),
            (c2, "ir_weighted_position", 1125000.0, "SRC-03",
             "Part C.2, table 10, the position of the future in zone 2", 500),
            (c2, "ir_weighted_position", -200000.0, "SRC-03", "Part C.2, table 10, 3 to 6 months", 500),
            (c2, "ir_vertical_disallowance", 50000.0, "SRC-03", "Part C.2, paragraph 2(a)", 500),
            (c2, "ir_horizontal_disallowance_within_zones", 80000.0, "SRC-03",
             "Part C.2, paragraph 2(b)", 500),
            (c2, "ir_net_position_charge", 3000000.0, "SRC-03", "Part C.2, paragraph 2(e)", 500),
            (c2, "ir_general_market_risk_charge", 4580000.0, "SRC-03", "Part C.2, paragraph 3", 500),
            (c2, "ir_general_market_risk_charge", 4580000.0, "SRC-03", "Part C.2, paragraph 3", 500),
            (v27, "ir_weighted_position", 100000000.0, "SRC-02", "MAR40.27", 500000),
            (v27, "ir_weighted_position", -90000000.0, "SRC-02", "MAR40.27", 500000),
            (v27, "ir_vertical_disallowance", 9000000.0, "SRC-02", "MAR40.27", 500000),
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
        self.assertEqual(len(names), 20)
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
        workbook["Results"]["K12"].value = 48267441.0
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
