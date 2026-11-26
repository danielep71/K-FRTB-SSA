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
PROJECT = frozenset(f"CALC-0{number}" for number in range(2, 10))


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


def under_crr(case: dict[str, Any], own_funds: float | None = 10_000) -> dict[str, Any]:
    """A fixture under the CRR regime, with total own funds unless None."""
    parameters: dict[str, object] = {"regime_code": "CRR"}
    if own_funds is not None:
        parameters["total_own_funds_amount"] = own_funds
    return {**case, "parameters": parameters}


def fx_case(amount: float, reporting_date: str = "2026-12-31") -> dict[str, Any]:
    """A fixture of CALC-06 with one long USD position of amount, at rate 1."""
    return {"calculation": "CALC-06", "reporting_date": reporting_date, "reporting_currency_code": "EUR",
            "datasets": {"DS-05": [{"id": "EUR", "inputs": {"fx_rate": 1.0}},
                                   {"id": "USD", "inputs": {"fx_rate": 1.0}}],
                         "DS-09": [{"id": "F1", "inputs": {"currency_code": "USD", "direction_code": "Long",
                                                           "net_position_amount": amount}}]}}


class CrrDerivation(unittest.TestCase):
    """tests/derivations/ssa.py under the CRR regime, on values that can be worked out by hand."""

    def test_the_category_of_a_row_is_crr_or_none_under_the_crr_and_never_crr_under_basel(self) -> None:
        cases: tuple[tuple[dict[str, object], str, str], ...] = (
            ({"sr_category_code": "crr", "rating_grade_code": " rw 20% "}, "OK",
             "INVALID: SR category must be GOV, QUAL, OTHER or NONE"),
            ({"sr_category_code": "NONE"}, "OK", "OK"),
            ({"sr_category_code": "GOV", "rating_grade_code": "AAA to AA-"},
             "INVALID: SR category must be CRR or NONE under the CRR", "OK"),
            ({"sr_category_code": "CRR", "rating_grade_code": "RW 35%"},
             "INVALID: no Article 336 charge for this risk weight",
             "INVALID: SR category must be GOV, QUAL, OTHER or NONE"),
            ({"sr_category_code": "CRR", "rating_grade_code": "AAA to AA-"},
             "INVALID: no Article 336 charge for this risk weight",
             "INVALID: SR category must be GOV, QUAL, OTHER or NONE"),
        )
        for inputs, crr, basel in cases:
            with self.subTest(inputs=inputs):
                self.assertEqual(derived(under_crr(ir_case(**inputs)))[("ir_position_status", "P1")], crr)
                self.assertEqual(derived(ir_case(**inputs))[("ir_position_status", "P1")], basel)

    def test_the_rates_of_article_336_by_risk_weight_and_residual_term(self) -> None:
        # 1,000,000 at each risk weight and final maturity: the rate of table 1 of Article 336(1),
        # half of it at 10% (Article 336(3)). 2028-12-30 is 730 days, two years, after the
        # reporting date: the limit of the middle range, which it belongs to.
        cases = (("RW 0%", "2040-12-31", 0.0), ("RW 10%", "2027-06-30", 1250.0),
                 ("RW 10%", "2028-12-30", 5000.0), ("RW 10%", "2029-01-01", 8000.0),
                 ("RW 20%", "2027-06-30", 2500.0), ("RW 50%", "2028-12-30", 10000.0),
                 ("RW 50%", "2040-12-31", 16000.0), ("RW 75%", "2027-01-31", 80000.0),
                 ("RW 100%", "2040-12-31", 80000.0), ("RW 150%", "2027-01-31", 120000.0))
        for grade, maturity, charge in cases:
            with self.subTest(grade=grade, maturity=maturity):
                case = under_crr(ir_case(sr_category_code="CRR", rating_grade_code=grade,
                                         final_maturity_date=maturity))
                case["calculation"] = "CALC-02"
                self.assertAlmostEqual(derived(case)[("ir_specific_risk_charge", None)], charge, places=6)

    def test_the_offset_between_zones_1_and_3_is_charged_at_150_percent_under_the_crr(self) -> None:
        # 700 weighted long in band 4 (zone 1) against 700 weighted short in band 11 (zone 3).
        weighted = {4: [700.0, 0.0], 11: [0.0, 700.0]}
        self.assertAlmostEqual(ssa.ladder_charges(weighted, ssa.ZONES_1_AND_3)[
            "ir_horizontal_disallowance_between_zones"], 700.0, places=9)
        self.assertAlmostEqual(ssa.ladder_charges(weighted, ssa.ZONES_1_AND_3_CRR)[
            "ir_horizontal_disallowance_between_zones"], 1050.0, places=9)

    def test_an_index_carries_no_charge_of_its_own_under_the_crr(self) -> None:
        case = {"calculation": "CALC-04", "reporting_date": "2026-12-31", "reporting_currency_code": "EUR",
                "datasets": {"DS-05": [{"id": "EUR", "inputs": {"fx_rate": 1.0}}],
                             "DS-07": [{"id": "E1", "inputs": {
                                 "market_code": "EA", "type_code": "Index", "identifier_code": "IDX",
                                 "currency_code": "EUR", "direction_code": "Long",
                                 "market_value_amount": 1000000}}]}}
        for regime, index in ((under_crr(case), 0.0), (case, 20000.0)):
            outputs = derived(regime)
            with self.subTest(parameters=regime.get("parameters")):
                self.assertAlmostEqual(outputs[("eq_index_charge", None)], index, places=6)
                self.assertAlmostEqual(outputs[("eq_general_market_risk_charge", None)], 80000.0, places=6)

    def test_no_fx_requirement_up_to_two_percent_of_the_total_own_funds_and_all_of_it_above(self) -> None:
        # 2% of 10,000 is 200: a position of 200 carries nothing, one of 200.01 its whole 8%.
        for amount, charge in ((200.0, 0.0), (200.01, 16.0008), (0.0, 0.0)):
            with self.subTest(amount=amount):
                outputs = derived(under_crr(fx_case(amount)))
                self.assertAlmostEqual(outputs[("fx_overall_net_open_position", None)], amount, places=9)
                self.assertAlmostEqual(outputs[("fx_charge", None)], charge, places=9)
        self.assertAlmostEqual(derived(fx_case(200.0))[("fx_charge", None)], 16.0, places=9)

    def test_an_unusable_regime_or_total_own_funds_stops_the_run(self) -> None:
        stopped: list[dict[str, object]] = [{
            "quantity": "run_status", "scope": None,
            "text": "STOPPED: 1 unusable input(s); see Checks. No results were calculated."}]
        # Under the CRR the total own funds must be a number greater than 0; an unusable regime is
        # read as Basel, whose run does not read them, so the two never count twice.
        cases: tuple[tuple[dict[str, object], list[dict[str, object]] | None], ...] = (
            ({"regime_code": "CRR"}, stopped),
            ({"regime_code": "CRR", "total_own_funds_amount": 0}, stopped),
            ({"regime_code": "CRR", "total_own_funds_amount": -10_000}, stopped),
            ({"regime_code": "CRR", "total_own_funds_amount": "10000"}, stopped),
            ({"regime_code": "EU"}, stopped),
            ({"regime_code": "", "total_own_funds_amount": 0}, stopped),
            ({"regime_code": " crr ", "total_own_funds_amount": 0.01}, None),
            ({"regime_code": "basel"}, None),
            ({}, None))
        for parameters, outputs in cases:
            with self.subTest(parameters=parameters):
                result = ssa.derive({**fx_case(200.0), "parameters": parameters})
                if outputs is None:
                    self.assertIn("fx_charge", [output["quantity"] for output in result])
                else:
                    self.assertEqual(result, outputs)
        # The de minimis of foreign_exchange() keeps its own guard for a caller that skips derive().
        with self.assertRaises(AssertionError):
            ssa.foreign_exchange(ssa.Case(under_crr(fx_case(200.0), None)), [])

    def test_the_transitional_factor_applies_under_the_crr_up_to_its_last_date(self) -> None:
        options = [{"id": name, "inputs": {"currency_code": "EUR", "charge_amount": 0}}
                   for name in ("IR", "EQ", "CO")]
        for date, crr_factor in (("2029-12-31", 0.9), ("2030-01-01", 1.0), ("2026-12-31", 0.9)):
            case = {**fx_case(1000.0, date), "calculation": "CALC-08"}
            case["datasets"] = {**case["datasets"], "DS-10": [
                *options, {"id": "FX", "inputs": {"currency_code": "EUR", "charge_amount": 0}}]}
            for regime, factor in ((under_crr(case), crr_factor), (case, 1.0)):
                outputs = derived(regime)
                with self.subTest(date=date, parameters=regime.get("parameters")):
                    self.assertEqual(outputs[("transitional_factor", None)], factor)
                    # 8% of 1,000 scaled by 1.20, then the factor; the risk-weighted assets follow it.
                    self.assertAlmostEqual(outputs[("own_funds_requirement", None)], 96.0 * factor, places=9)
                    self.assertAlmostEqual(outputs[("risk_weighted_assets", None)], 1200.0 * factor,
                                           places=9)


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
        self.assertEqual(len(self.cases), 47)
        for calculation in sorted(PROJECT):
            categories = {case["category"] for case in self.cases.values()
                          if case["calculation"] == calculation}
            with self.subTest(calculation=calculation):
                self.assertIn("nominal", categories)
                # The option charges are taken as entered: they have no limit of their own.
                if calculation != "CALC-07":
                    self.assertIn("boundary", categories)
        invalid = {case["calculation"] for case in self.cases.values() if case["category"] == "invalid"}
        self.assertEqual(invalid, {"CALC-03", "CALC-04", "CALC-05", "CALC-06", "CALC-08", "CALC-09"})

    def test_each_parameter_that_stops_the_run_has_a_negative_and_a_boundary_case(self) -> None:
        # The regime: one that is neither value, and the CRR written in lower case, which is read.
        # The total own funds under the CRR: none, and 0, the largest value that is not greater
        # than 0. Each stopped case expects the status line alone; the CRR in lower case expects
        # what the CRR in capitals does.
        stopped = "STOPPED: 1 unusable input(s); see Checks. No results were calculated."
        for name, regime, own_funds in (("illustrative-invalid-regime", "EU", 16750),
                                        ("illustrative-invalid-crr-own-funds", "CRR", None),
                                        ("illustrative-crr-zero-own-funds", "CRR", 0)):
            with self.subTest(case=name):
                parameters = self.cases[name]["parameters"]
                self.assertEqual((parameters["regime_code"], parameters.get("total_own_funds_amount")),
                                 (regime, own_funds))
                self.assertEqual(self.cases[name]["datasets"], self.cases["mar40-61-table-9"]["datasets"])
                self.assertEqual(self.expected(name), {("run_status", None): stopped})
        # The FX table, rules C05 to C07: the positions of table 9 with one unusable row each.
        table = {record["id"]: record["inputs"]["fx_rate"]
                 for record in self.cases["mar40-61-table-9"]["datasets"]["DS-05"]}
        for name, rows in (("illustrative-invalid-fx-rate-not-positive", {**table, "USD": 0}),
                           ("illustrative-invalid-fx-rate-not-a-number", {**table, "USD": "abc"}),
                           ("illustrative-invalid-fx-currency-listed-twice", None),
                           ("illustrative-invalid-fx-reporting-currency-missing",
                            {code: rate for code, rate in table.items() if code != "CHF"}),
                           ("illustrative-invalid-fx-reporting-currency-rate", {**table, "CHF": 2.0})):
            with self.subTest(case=name):
                case = self.cases[name]
                self.assertEqual(case["category"], "invalid")
                self.assertNotIn("parameters", case)
                self.assertEqual(case["datasets"]["DS-09"], self.cases["mar40-61-table-9"]["datasets"]["DS-09"])
                listed = [(record["id"], record["inputs"]["fx_rate"]) for record in case["datasets"]["DS-05"]]
                if rows is None:
                    self.assertEqual(sorted(listed), sorted([*table.items(), ("JPY", 1.0)]))
                else:
                    self.assertEqual(dict(listed), rows)
                    self.assertEqual(len(listed), len(rows))
                self.assertEqual(self.expected(name), {("run_status", None): stopped})
        lower, upper = "illustrative-crr-regime-in-lower-case", "illustrative-crr-fx-de-minimis-limit"
        self.assertEqual(self.cases[lower]["parameters"],
                         {**self.cases[upper]["parameters"], "regime_code": "crr"})
        self.assertEqual(self.cases[lower]["datasets"], self.cases[upper]["datasets"])
        self.assertEqual(self.expected(lower), self.expected(upper))

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
        # The sample months of Eligibility, in the order of their rows.
        months = self.cases["illustrative-eligibility-sample-months"]
        self.assertEqual(months["parameters"]["regime_code"], "CRR")
        self.assertEqual([tuple(record["inputs"][key] for key in ssa.MONTH_KEYS)
                          for record in months["datasets"]["DS-11"]],
                         [(row[0].isoformat(), *row[1:]) for row in ssa_reference.ELIGIBILITY])
        # The cases of one calculation hold the same records as the whole portfolio.
        for case, dataset in (("illustrative-sample-ir-general-market-risk", "DS-06"),
                              ("illustrative-sample-equity", "DS-07"),
                              ("illustrative-sample-commodities", "DS-08"),
                              ("illustrative-sample-fx", "DS-09")):
            with self.subTest(case=case):
                self.assertEqual(self.cases[case]["datasets"][dataset], datasets[dataset])

    def expected(self, name: str) -> dict[tuple[str, str | None], Any]:
        """The expected outputs of a registered case by quantity and record."""
        data = json.loads((EXPECTED / f"{name}.json").read_text(encoding="utf-8"))
        return {(output["quantity"], None if output["scope"] is None else output["scope"]["record"]):
                output.get("value", output.get("text")) for output in data["outputs"]}

    def assert_same_positions(self, crr: str, basel: str) -> None:
        """The CRR case holds the records of its Basel case, but the category and the risk weight
        of Article 336 that a debt security has in place of its Basel category and grade."""
        crr_case, basel_case = self.cases[crr], self.cases[basel]
        self.assertEqual(crr_case["parameters"]["regime_code"], "CRR")
        self.assertNotIn("parameters", basel_case)
        self.assertEqual(crr_case["datasets"].keys(), basel_case["datasets"].keys())
        for dataset, records in basel_case["datasets"].items():
            for mine, theirs in zip(crr_case["datasets"][dataset], records, strict=True):
                with self.subTest(case=crr, record=theirs["id"]):
                    self.assertEqual(mine["id"], theirs["id"])
                    category = theirs["inputs"].get("sr_category_code")
                    if dataset == "DS-06" and category not in (None, "NONE"):
                        self.assertEqual(mine["inputs"]["sr_category_code"], "CRR")
                        rates = {rate for (cat, grade, _), rate in ssa.SPECIFIC.items()
                                 if (cat, grade) == (category.upper(), theirs["inputs"]["rating_grade_code"].upper())}
                        weight = mine["inputs"]["rating_grade_code"].upper()
                        self.assertEqual({rate for (cat, grade, _), rate in ssa.SPECIFIC.items()
                                          if (cat, grade) == ("CRR", weight)}, rates)
                        mine = {**mine, "inputs": {**mine["inputs"], "sr_category_code": category,
                                                   "rating_grade_code": theirs["inputs"]["rating_grade_code"]}}
                    self.assertEqual(mine, theirs)

    def assert_differ_only_in(self, crr: str, basel: str, changed: dict[tuple[str, str | None], float]) -> None:
        """The expected values of the CRR case equal those the derivation gives the Basel case, but
        those of changed, which the CRR case has. The derivation stands for the Basel case because a
        published value of it is the source's, rounded as the source prints it."""
        mine, theirs = self.expected(crr), derived(self.cases[basel])
        common = mine.keys() & theirs.keys()
        self.assertTrue(changed.keys() <= common)
        for key in sorted(common, key=str):
            with self.subTest(case=crr, output=key):
                if key in changed:
                    self.assertNotAlmostEqual(mine[key], theirs[key], places=6)
                    self.assertAlmostEqual(mine[key], changed[key], places=6)
                elif isinstance(theirs[key], str):
                    self.assertEqual(mine[key], theirs[key])
                else:
                    self.assertAlmostEqual(mine[key], theirs[key], places=6)

    def test_the_crr_cases_hold_the_positions_of_a_basel_case_and_differ_by_a_rule_of_the_crr(self) -> None:
        # Article 339(9)(f): example C.2 offsets 1,000,000 between zones 1 and 3, charged at 150% in
        # place of 100%, 500,000 more between the zones and in the charge.
        c2 = derived(self.cases["bcbs119-c2-maturity-method"])
        self.assert_same_positions("illustrative-crr-maturity-method-c2", "bcbs119-c2-maturity-method")
        self.assert_differ_only_in("illustrative-crr-maturity-method-c2", "bcbs119-c2-maturity-method", {
            (quantity, scope): c2[(quantity, scope)] + 500000
            for quantity in ("ir_horizontal_disallowance_between_zones", "ir_general_market_risk_charge")
            for scope in ("USD", None) if (quantity, scope) in c2})
        # Article 344: no charge on an index.
        self.assert_same_positions("illustrative-crr-sample-equity", "illustrative-sample-equity")
        self.assert_differ_only_in("illustrative-crr-sample-equity", "illustrative-sample-equity",
                                   {("eq_index_charge", None): 0.0})
        # Article 351: no requirement where the position of 335 is 2% of the total own funds; all of it
        # where it is above.
        for crr, waived in (("illustrative-crr-fx-de-minimis-limit", True),
                            ("illustrative-crr-fx-above-de-minimis", False)):
            self.assert_same_positions(crr, "mar40-61-table-9")
            self.assertAlmostEqual(0.02 * self.cases[crr]["parameters"]["total_own_funds_amount"],
                                   335.0 if waived else 334.999, places=9)
            self.assert_differ_only_in(crr, "mar40-61-table-9", {("fx_charge", None): 0.0} if waived else {})
        # Article 495s(2): the factor 0.9 up to 2029-12-31, then 1; the total own funds leave the FX
        # requirement charged.
        only = derived(self.cases["illustrative-requirement-fx-and-options-only"])
        for crr, factor in (("illustrative-crr-transitional-end-date", 0.9),
                            ("illustrative-crr-after-transitional-end-date", 1.0)):
            self.assert_same_positions(crr, "illustrative-requirement-fx-and-options-only")
            changes: dict[tuple[str, str | None], float] = {} if factor == 1.0 else {
                ("transitional_factor", None): factor,
                ("own_funds_requirement", None): factor * only[("own_funds_requirement", None)],
                ("risk_weighted_assets", None): factor * only[("risk_weighted_assets", None)]}
            self.assert_differ_only_in(crr, "illustrative-requirement-fx-and-options-only", changes)

    def test_the_sample_under_both_regimes_differs_by_the_index_charge_and_the_transitional_factor(self) -> None:
        crr, basel = "illustrative-crr-sample-requirement", "illustrative-sample-requirement"
        self.assert_same_positions(crr, basel)
        sample = derived(self.cases[basel])
        index = derived(self.cases["illustrative-sample-equity"])[("eq_index_charge", None)]
        equity = sample[("requirement_before_scaling", "EQ")] - index
        scaled = sample[("own_funds_requirement", None)] - ssa.SCALING["EQ"] * index
        # The sample offsets nothing between zones 1 and 3, and its position of 335 million is above
        # 2% of the 10 billion of total own funds: the interest rate and FX requirements do not change.
        self.assert_differ_only_in(crr, basel, {
            ("requirement_before_scaling", "EQ"): equity,
            ("requirement_after_scaling", "EQ"): ssa.SCALING["EQ"] * equity,
            ("transitional_factor", None): 0.9,
            ("own_funds_requirement", None): 0.9 * scaled,
            ("risk_weighted_assets", None): 0.9 * ssa.RWA_MULTIPLIER * scaled})

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
        self.assertEqual(len(names), 47)
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

    def test_a_fixture_sets_its_parameters_and_its_months(self) -> None:
        assert openpyxl is not None
        workbook = openpyxl.load_workbook(TEMPLATE)
        verify = self.tool.verify
        case = fixture("illustrative-invalid-eligibility-months")
        self.tool.write_inputs(workbook, case)
        self.assertEqual(verify.named_cell(workbook, "Regime").value, "CRR")
        self.assertEqual(verify.named_cell(workbook, "TotalOwnFunds").value, 10000000000)
        sheet = workbook["Eligibility"]
        # Row 5 is the first record; a month end that is no ISO date is written as it is.
        self.assertEqual(sheet["C5"].value.isoformat(), "2025-10-31")
        self.assertEqual([sheet[f"C{row}"].value for row in (6, 7, 12)], ["31.11.2025", 20251231, None])
        self.assertEqual([sheet.cell(7, column).value for column in range(4, 9)],
                         [8000000000, 100000000, 0, 10000000, 0])
        self.assertIsNone(sheet["H14"].value)
        self.assertIsNone(sheet["C15"].value)
        # A case that sets no parameter has the reference regime and no own funds.
        self.tool.write_inputs(workbook, fixture("illustrative-sample-fx"))
        self.assertEqual(verify.named_cell(workbook, "Regime").value, "Basel")
        self.assertIsNone(verify.named_cell(workbook, "TotalOwnFunds").value)
        self.assertIsNone(sheet["C5"].value)
        with self.assertRaises(ValueError):
            self.tool.write_inputs(workbook, {**case, "parameters": {"scaling_ir_factor": 1.0}})

    def test_the_book_of_a_record_is_typed_in_the_book_column(self) -> None:
        assert openpyxl is not None
        workbook = openpyxl.load_workbook(TEMPLATE)
        # The Book column is the one before Comment: P, K, L and H.
        for case, sheet, column in (("illustrative-ir-book", "IRPositions", 16),
                                    ("illustrative-eq-book", "EQPositions", 11),
                                    ("illustrative-co-book", "COPositions", 12),
                                    ("illustrative-fx-book", "FXPositions", 8)):
            with self.subTest(case=case):
                self.tool.write_inputs(workbook, fixture(case))
                positions = workbook[sheet]
                self.assertEqual(positions.cell(4, column).value, "Book")
                self.assertEqual([positions.cell(row, column).value for row in (5, 6, 7)],
                                 ["Trading", None, "Non-trading"])

    def test_a_value_that_is_not_what_a_cell_can_hold_is_written_as_it_is(self) -> None:
        self.assertEqual(self.tool.cell_value("final_maturity_date", 20311231), 20311231)
        self.assertEqual(self.tool.cell_value("final_maturity_date", "31/12/2030"), "31/12/2030")
        self.assertEqual(self.tool.cell_value("position_value_amount", "ten"), "ten")

    def test_an_output_is_read_from_its_cell(self) -> None:
        assert openpyxl is not None
        workbook = openpyxl.load_workbook(TEMPLATE)
        case = fixture("illustrative-sample-requirement")
        workbook["Results"]["K12"].value = 48267441.0
        workbook["Results"]["J12"].value = 0.9
        workbook["Results"]["I8"].value = 2792000.0
        workbook["IRCalc"]["D6"].value = "OK"
        workbook["IR_Zones"]["C6"].value = "EUR"
        for cell, value in (("I6", 10.0), ("N6", 20.0), ("S6", 30.0)):
            workbook["IR_Zones"][cell].value = value

        def read(quantity: str, dataset: str | None = None, record: str | None = None) -> object:
            scope = None if dataset is None else {"dataset": dataset, "record": record}
            return self.tool.actual(workbook, case, {"quantity": quantity, "scope": scope})

        self.assertEqual(read("own_funds_requirement"), 48267441.0)
        # The status line without the time of the run that ends it.
        status = self.tool.verify.named_cell(workbook, "Results_Status")
        status.value = "STOPPED: 1 unusable input(s); see Checks. No results were calculated. Run 2026-11-21 10:00."
        self.assertEqual(read("run_status"), "STOPPED: 1 unusable input(s); see Checks. No results were calculated.")
        status.value = None
        self.assertEqual(read("run_status"), "")
        self.assertEqual(self.tool.run_line(" OK: Run of 3 positions. Run 10:00. "), "OK: Run of 3 positions.")
        self.assertEqual(read("transitional_factor"), 0.9)
        self.assertEqual(read("requirement_before_scaling", "DS-10", "EQ"), 2792000.0)
        self.assertEqual(read("ir_position_status", "DS-06", case["datasets"]["DS-06"][1]["id"]), "OK")
        self.assertEqual(read("ir_horizontal_disallowance_within_zones", "DS-05", "EUR"), 60.0)
        self.assertEqual(read("ir_vertical_disallowance", "DS-05", "ZAR"), "no such currency on IR_Zones")
        # An output of CALC-09: a cell of the row of its month, or of the result block.
        months = fixture("illustrative-eligibility-limits")
        workbook["Eligibility"]["P6"].value = 50000001.0
        workbook["Eligibility"]["J33"].value = "NO"
        self.assertEqual(self.tool.actual(workbook, months, {
            "quantity": "eligibility_size_94", "scope": {"dataset": "DS-11", "record": "2026-05"}}),
            50000001.0)
        self.assertEqual(self.tool.actual(workbook, months, {
            "quantity": "eligibility_restart_94", "scope": None}), "NO")

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
