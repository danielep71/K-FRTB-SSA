"""Derivation of the expected values of the simplified standardised approach.

Usage: python -I tests/derivations/ssa.py tests/fixtures/<case-id>.json

It reads one fixture and prints, in the derivation output format, the outputs
of the fixture's calculation, CALC-02 to CALC-09, as
docs/methodology/CALCULATION_CONTRACTS.md states them: the status of each
position row, and the quantities of the rows that are used. A fixture gives
the reporting date and currency, the FX rates (DS-05), the positions (DS-06
to DS-09), the option charges (DS-10) and the months of the size of the
business (DS-11). Every parameter has its reference value, which the
Parameters table of the contracts traces to SRC-02, or to SRC-04 and SRC-06
for those of the CRR regime, but those a fixture sets under "parameters": the
regime and the total own funds. Under the CRR regime it applies the rules of
the contracts that name it: the categories of Article 336 (CALC-02), 150%
between zones 1 and 3 (CALC-03), no charge on an index (CALC-04), the de
minimis of Article 351 (CALC-06) and the transitional multiplier of Article
495s(2) (CALC-08).

eligibility() derives CALC-09, the eligibility under the CRR, from the rows of
the Eligibility sheet and the regime; derive() calls it for a fixture of
CALC-09, and a test may call it with rows that no fixture can hold.

It is written from the contracts and shares no code with the VBA engine. Its
values are illustrative all the same, because the project derived them.
"""

from __future__ import annotations

import calendar
import datetime
import json
import sys
from collections.abc import Mapping, Sequence
from typing import Any

# ------------------------------------------------------------- parameters
DAYS_PER_YEAR = 365.0
COUPON_THRESHOLD = 0.03
SCALING = {"IR": 1.30, "EQ": 3.50, "CO": 1.90, "FX": 1.20}
RWA_MULTIPLIER = 12.5
VERTICAL = 0.10
WITHIN_ZONE = {1: 0.40, 2: 0.30, 3: 0.30}
ADJACENT_ZONES = 0.40
ZONES_1_AND_3 = 1.00
EQ_SPECIFIC, EQ_INDEX, EQ_GENERAL = 0.08, 0.02, 0.08
FX_RATE = 0.08
CO_DIRECTIONAL, CO_GROSS = 0.15, 0.03
GOLD = "XAU"
# The parameters of the CRR regime: SRC-04, Articles 339(9)(f), 344 and 351, and SRC-06,
# Article 495s(2), whose multiplier applies on and before its last date.
ZONES_1_AND_3_CRR = 1.50
EQ_INDEX_CRR = 0.0
FX_DE_MINIMIS = 0.02
TRANSITIONAL_MULTIPLIER, TRANSITIONAL_END = 0.9, "2029-12-31"
# The instruments of an interest rate row that DS-06 allows, as codes.
INSTRUMENTS = frozenset(("BOND", "FRN", "SWAP LEG", "FUTURE LEG", "FRA LEG", "FORWARD LEG", "OTHER"))
# Table 4 of MAR40.26: zone, upper limit in years for a coupon of 3% or more,
# upper limit for a lower coupon (None: no upper limit, or no such band), and
# risk weight. A term equal to an upper limit belongs to the band.
LADDER: list[tuple[int, float | None, float | None, float]] = [
    (1, 1 / 12, 1 / 12, 0.0000), (1, 0.25, 0.25, 0.0020), (1, 0.5, 0.5, 0.0040),
    (1, 1.0, 1.0, 0.0070), (2, 2.0, 1.9, 0.0125), (2, 3.0, 2.8, 0.0175),
    (2, 4.0, 3.6, 0.0225), (3, 5.0, 4.3, 0.0275), (3, 7.0, 5.7, 0.0325),
    (3, 10.0, 7.3, 0.0375), (3, 15.0, 9.3, 0.0450), (3, 20.0, 10.6, 0.0525),
    (3, None, 12.0, 0.0600), (3, None, 20.0, 0.0800), (3, None, None, 0.1250),
]
# Table 1 of MAR40.6: (category, rating grade, range of residual maturity).
SHORT, MEDIUM, LONG, ANY = "<=6m", "6m-24m", ">24m", "Any"
BY_MATURITY = {SHORT: 0.0025, MEDIUM: 0.01, LONG: 0.016}
SPECIFIC: dict[tuple[str, str, str], float] = {
    ("GOV", "AAA TO AA-", ANY): 0.0,
    **{("GOV", "A+ TO BBB-", band): rate for band, rate in BY_MATURITY.items()},
    ("GOV", "BB+ TO BB-", ANY): 0.08, ("GOV", "B+ TO B-", ANY): 0.08,
    ("GOV", "BELOW B-", ANY): 0.12, ("GOV", "UNRATED", ANY): 0.08,
    **{("QUAL", "AAA TO AA-", band): rate for band, rate in BY_MATURITY.items()},
    **{("QUAL", "A+ TO BBB-", band): rate for band, rate in BY_MATURITY.items()},
    ("OTHER", "BB+ TO BB-", ANY): 0.08, ("OTHER", "B+ TO B-", ANY): 0.12,
    ("OTHER", "BELOW B-", ANY): 0.12, ("OTHER", "UNRATED", ANY): 0.08,
    # Table 1 of Article 336(1) of the CRR, by the risk weight of the standardised approach
    # for credit risk, and half of the 20% and 50% rates for the covered bonds of Article 336(3).
    ("CRR", "RW 0%", ANY): 0.0,
    **{("CRR", "RW 10%", band): rate / 2 for band, rate in BY_MATURITY.items()},
    **{("CRR", "RW 20%", band): rate for band, rate in BY_MATURITY.items()},
    **{("CRR", "RW 50%", band): rate for band, rate in BY_MATURITY.items()},
    ("CRR", "RW 75%", ANY): 0.08, ("CRR", "RW 100%", ANY): 0.08, ("CRR", "RW 150%", ANY): 0.12,
}
# The serials Excel holds as dates: 1 January 1900 to 31 December 9999.
SERIAL_EPOCH = datetime.date(1899, 12, 30)
SERIAL_MIN, SERIAL_MAX = 1, 2958465
OK, INVALID, EXCLUDED = "OK", "INVALID: ", "EXCLUDED: "
NOT_ASSESSED = "NOT ASSESSED"
RISK_CLASSES = ("IR", "EQ", "CO", "FX")

# Articles 325a(1) and 94(1) of the CRR: the share of the total assets and the amount in
# euro that the size of the business may not exceed, by approach, under the names of
# the parameters, which eligibility() takes from its params or, missing there, here.
ELIGIBILITY_REFERENCE = {"EligibilityShare325a": 0.10, "EligibilityCap325a": 500_000_000.0,
                         "EligibilityShare94": 0.05, "EligibilityCap94": 50_000_000.0}
APPROACHES = ("325a", "94")
# The months of the run and of the count of an approach's conditions that are not met:
# a run of this many months, or a count over the limit among the last WINDOW months, is an event.
EVENT_RUN, EVENT_COUNT, WINDOW = 3, 6, 12
MONTH_REASONS = ("INVALID: MonthEnd must be a date",
                 "INVALID: MonthEnd must be the last day of a month",
                 "INVALID: MonthEnd must be one month after the previous row",
                 "INVALID: TotalAssets must be a number greater than 0",
                 "INVALID: a long or short total must be a number")

Record = dict[str, Any]
Output = dict[str, Any]


# ---------------------------------------------------------------- reading
def text(value: object) -> str:
    """A cell as text, without the blanks at its ends."""
    return "" if value is None else str(value).strip()


def code(value: object) -> str:
    """A code, compared without regard to case."""
    return text(value).upper()


def is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def serial(value: object) -> float | None:
    """A date as its serial number; None when the value is neither a date nor a number."""
    if is_number(value):
        assert isinstance(value, (int, float))
        return float(value)
    if isinstance(value, str):
        try:
            return float((datetime.date.fromisoformat(value) - SERIAL_EPOCH).days)
        except ValueError:
            return None
    return None


def is_date(number: float) -> bool:
    return SERIAL_MIN <= number <= SERIAL_MAX


def sign(direction: object) -> float:
    return {"LONG": 1.0, "SHORT": -1.0}.get(code(direction), 0.0)


def book_status(inputs: Record, in_scope: bool) -> str | None:
    """The status the book of a position gives it, or None when the other tests decide.

    A blank book is the trading book. A non-trading book position is calculated as a
    trading book one where its class is in the scope (foreign exchange, commodities),
    and excluded where it is not (interest rate, equity)."""
    book = code(inputs.get("book_code"))
    if book not in ("", "TRADING", "NON-TRADING"):
        return INVALID + "Book must be Trading or Non-trading"
    if book == "NON-TRADING" and not in_scope:
        return EXCLUDED + "non-trading book position outside the market risk scope"
    return None


def duplicates(records: list[Record]) -> set[str]:
    """The identifiers that more than one record holds, without regard to case."""
    seen: dict[str, int] = {}
    for record in records:
        seen[code(record["id"])] = seen.get(code(record["id"]), 0) + 1
    return {identifier for identifier, count in seen.items() if count > 1}


class Case:
    """One fixture: its dates, its FX rates and its records."""

    def __init__(self, fixture: dict[str, Any]) -> None:
        self.calculation = str(fixture["calculation"])
        self.reporting = code(fixture["reporting_currency_code"])
        as_of = serial(fixture["reporting_date"])
        assert as_of is not None
        self.as_of = as_of
        self.datasets: dict[str, list[Record]] = fixture["datasets"]
        parameters = fixture.get("parameters", {})
        self.regime = parameters.get("regime_code", "Basel")
        self.crr = code(self.regime) == "CRR"
        self.own_funds = parameters.get("total_own_funds_amount")
        self.rates = {code(record["id"]): record["inputs"].get("fx_rate")
                      for record in self.records("DS-05")}

    def records(self, dataset: str) -> list[Record]:
        return self.datasets.get(dataset, [])

    def years(self, date: float) -> float:
        return (date - self.as_of) / DAYS_PER_YEAR

    def no_currency(self, currency: str, gold_allowed: bool = False) -> bool:
        """Whether a currency is missing from the FX rates, or is gold where gold is not a currency."""
        return currency == "" or currency not in self.rates or (currency == GOLD and not gold_allowed)


# --------------------------------------------------------- interest rates
def band_of(maturity: float, high_coupon: bool) -> int:
    limits = [row[1] if high_coupon else row[2] for row in LADDER]
    return 1 + sum(1 for limit in limits if limit is not None and limit < maturity)


def specific_key(case: Case, inputs: Record) -> tuple[str, str, str] | None:
    """The key of table 1 for a row, or None when the table has none for it."""
    category, grade = code(inputs.get("sr_category_code")), code(inputs.get("rating_grade_code"))
    maturity = serial(inputs.get("final_maturity_date"))
    if maturity is None:
        return None
    term = case.years(maturity)
    band = SHORT if term <= 0.5 else MEDIUM if term <= 2 else LONG
    for key in ((category, grade, band), (category, grade, ANY)):
        if key in SPECIFIC:
            return key
    return None


def ir_value_reason(case: Case, record: Record, repeated: set[str]) -> str | None:
    """Why the identifier, currency, direction, value or rate type of a row rejects it."""
    inputs = record["inputs"]
    currency = code(inputs.get("currency_code"))
    value = inputs.get("position_value_amount")
    if code(record["id"]) in repeated:
        return "duplicate PositionID"
    if code(inputs.get("instrument_code")) not in INSTRUMENTS:
        return "Instrument must be Bond, FRN, Swap leg, Future leg, FRA leg, Forward leg or Other"
    if currency == "" or currency not in case.rates:
        return "currency not in FX_Table"
    if currency == GOLD:
        return "gold is FX risk; enter it on FXPositions"
    if sign(inputs.get("direction_code")) == 0:
        return "Direction must be Long or Short"
    if not is_number(value):
        return "Position value must be a number"
    if value < 0:
        return "Position value must be >= 0 (Direction gives the sign)"
    if code(inputs.get("rate_type_code")) not in ("FIXED", "FLOATING"):
        return "Rate type must be Fixed or Floating"
    return None


def ir_date_reason(case: Case, inputs: Record) -> str | None:
    """Why the dates or the coupon of a row reject it."""
    maturity, repricing = serial(inputs.get("final_maturity_date")), serial(inputs.get("next_repricing_date"))
    floating = code(inputs.get("rate_type_code")) == "FLOATING"
    if maturity is None:
        return "Final maturity date missing"
    if not is_date(maturity):
        return "Final maturity date is not a date"
    if maturity < case.as_of:
        return "final maturity before AsOfDate"
    if not floating:
        return None if is_number(inputs.get("coupon_rate")) else "Coupon required for a fixed-rate position"
    if repricing is None:
        return "Next repricing date required for a floating-rate position"
    if not is_date(repricing):
        return "Next repricing date is not a date"
    if not case.as_of <= repricing <= maturity:
        return "repricing date outside AsOfDate to final maturity"
    return None


def ir_category_reason(case: Case, inputs: Record) -> str | None:
    """Why the category or the rating grade of a row rejects it: under the CRR regime the
    category is CRR or NONE, under the Basel regime one of MAR40.6."""
    category = code(inputs.get("sr_category_code"))
    if case.crr and category not in ("CRR", "NONE"):
        return "SR category must be CRR or NONE under the CRR"
    if not case.crr and category not in ("GOV", "QUAL", "OTHER", "NONE"):
        return "SR category must be GOV, QUAL, OTHER or NONE"
    if category != "NONE" and specific_key(case, inputs) is None:
        return ("no Article 336 charge for this risk weight" if case.crr
                else "no MAR40.6 charge for this category and rating grade")
    return None


def ir_status(case: Case, record: Record, repeated: set[str]) -> str:
    """OK, or the first test of the contract's list that the row fails."""
    if code(record["id"]) not in repeated and (book := book_status(record["inputs"], False)) is not None:
        return book
    reason = (ir_value_reason(case, record, repeated) or ir_date_reason(case, record["inputs"])
              or ir_category_reason(case, record["inputs"]))
    return OK if reason is None else INVALID + reason


def offset(first: float, second: float) -> tuple[float, float, float]:
    """The matched amount of two net positions of opposite sign, and what is left of each."""
    if first * second >= 0:
        return 0.0, first, second
    matched = min(abs(first), abs(second))
    return (matched, first - (matched if first > 0 else -matched),
            second - (matched if second > 0 else -matched))


def ladder_charges(weighted: dict[int, list[float]], zones_1_and_3: float) -> dict[str, float]:
    """The charges of one currency's ladder, from its weighted longs and shorts per band, with
    the disallowance between zones 1 and 3 of the regime."""
    vertical = 0.0
    longs = {1: 0.0, 2: 0.0, 3: 0.0}
    shorts = {1: 0.0, 2: 0.0, 3: 0.0}
    for band, (long, short) in weighted.items():
        vertical += VERTICAL * min(long, short)
        zone = LADDER[band - 1][0]
        if long >= short:
            longs[zone] += long - short
        else:
            shorts[zone] += short - long
    within = sum(WITHIN_ZONE[zone] * min(longs[zone], shorts[zone]) for zone in (1, 2, 3))
    first, second, third = (longs[zone] - shorts[zone] for zone in (1, 2, 3))
    matched_12, first, second = offset(first, second)
    matched_23, second, third = offset(second, third)
    matched_13, first, third = offset(first, third)
    between = ADJACENT_ZONES * (matched_12 + matched_23) + zones_1_and_3 * matched_13
    net = abs(first + second + third)
    return {"ir_vertical_disallowance": vertical,
            "ir_horizontal_disallowance_within_zones": within,
            "ir_horizontal_disallowance_between_zones": between,
            "ir_net_position_charge": net,
            "ir_general_market_risk_charge": vertical + within + between + net}


def interest_rate(case: Case, outputs: list[Output]) -> tuple[float, float]:
    """Add the outputs of CALC-02 and CALC-03; return the two requirements."""
    records = case.records("DS-06")
    repeated = duplicates(records)
    ladders: dict[str, dict[int, list[float]]] = {}
    issues: dict[str, tuple[float, float, str]] = {}   # key -> net value, rate, currency
    for record in records:
        if text(record["id"]) == "":
            continue
        status = ir_status(case, record, repeated)
        outputs.append(row("ir_position_status", "DS-06", record, text=status))
        if status != OK:
            continue
        inputs = record["inputs"]
        currency = code(inputs["currency_code"])
        value = sign(inputs["direction_code"]) * inputs["position_value_amount"] * case.rates[currency]
        floating = code(inputs["rate_type_code"]) == "FLOATING"
        date = serial(inputs["next_repricing_date"] if floating else inputs["final_maturity_date"])
        assert date is not None
        band = band_of(case.years(date), floating or inputs["coupon_rate"] >= COUPON_THRESHOLD)
        weighted = value * LADDER[band - 1][3]
        cell = ladders.setdefault(currency, {}).setdefault(band, [0.0, 0.0])
        cell[0 if weighted > 0 else 1] += abs(weighted)
        outputs.append(row("ir_time_band", "DS-06", record, value=band))
        outputs.append(row("ir_weighted_position", "DS-06", record, value=weighted))
        if code(inputs["sr_category_code"]) != "NONE":
            key = specific_key(case, inputs)
            assert key is not None
            # An entered issue and a position without one never net.
            issue = text(inputs.get("issue_id"))
            name = "ISSUE:" + issue.upper() if issue else "POS:" + code(record["id"])
            net, rate, first_currency = issues.get(name, (0.0, SPECIFIC[key], currency))
            issues[name] = (net + value, rate, first_currency)
    specific_total = general_total = 0.0
    for currency in sorted(case.rates):
        specific = sum(abs(net) * rate for net, rate, issue_currency in issues.values()
                       if issue_currency == currency)
        charges = ladder_charges(ladders.get(currency, {}), ZONES_1_AND_3_CRR if case.crr else ZONES_1_AND_3)
        specific_total += specific
        general_total += charges["ir_general_market_risk_charge"]
        if currency == GOLD:
            continue
        outputs.append(by_currency("ir_specific_risk_charge", currency, specific))
        outputs.extend(by_currency(quantity, currency, value) for quantity, value in charges.items())
    outputs.append(total("ir_specific_risk_charge", specific_total))
    outputs.append(total("ir_general_market_risk_charge", general_total))
    return specific_total, general_total


# ----------------------------------------------------------------- equity
def eq_status(case: Case, record: Record, repeated: set[str]) -> str:
    inputs = record["inputs"]
    value = inputs.get("market_value_amount")
    if code(record["id"]) in repeated:
        reason = "duplicate PositionID"
    elif (book := book_status(inputs, False)) is not None:
        return book
    elif code(inputs.get("market_code")) == "":
        reason = "Market missing"
    elif code(inputs.get("type_code")) not in ("SINGLE NAME", "INDEX"):
        reason = "Type must be Single name or Index"
    elif code(inputs.get("identifier_code")) == "":
        reason = "stock or index identifier missing"
    elif case.no_currency(code(inputs.get("currency_code"))):
        reason = "currency not in FX_Table"
    elif sign(inputs.get("direction_code")) == 0:
        reason = "Direction must be Long or Short"
    elif not is_number(value):
        reason = "Market value must be a number"
    elif value < 0:
        reason = "Market value must be >= 0 (Direction gives the sign)"
    else:
        return OK
    return INVALID + reason


def equity(case: Case, outputs: list[Output]) -> tuple[float, float, float]:
    """Add the outputs of CALC-04; return specific risk, the index charge and general market risk."""
    records = case.records("DS-07")
    repeated = duplicates(records)
    nets: dict[tuple[str, str, str], float] = {}
    markets: dict[str, float] = {}
    used = []
    for record in records:
        if text(record["id"]) == "":
            continue
        status = eq_status(case, record, repeated)
        outputs.append(row("eq_position_status", "DS-07", record, text=status))
        if status != OK:
            continue
        inputs = record["inputs"]
        value = (sign(inputs["direction_code"]) * inputs["market_value_amount"]
                 * case.rates[code(inputs["currency_code"])])
        key = (code(inputs["market_code"]), code(inputs["type_code"]), code(inputs["identifier_code"]))
        nets[key] = nets.get(key, 0.0) + value
        markets[key[0]] = markets.get(key[0], 0.0) + value
        used.append((record, key))
    outputs.extend(row("eq_net_position", "DS-07", record, value=nets[key]) for record, key in used)
    specific = sum(EQ_SPECIFIC * abs(net) for (_, kind, _), net in nets.items() if kind == "SINGLE NAME")
    rate = EQ_INDEX_CRR if case.crr else EQ_INDEX
    index = sum(rate * abs(net) for (_, kind, _), net in nets.items() if kind == "INDEX")
    general = sum(EQ_GENERAL * abs(net) for net in markets.values())
    outputs.append(total("eq_specific_risk_charge", specific))
    outputs.append(total("eq_index_charge", index))
    outputs.append(total("eq_general_market_risk_charge", general))
    return specific, index, general


# ------------------------------------------------------------ commodities
def co_status(case: Case, record: Record, repeated: set[str]) -> str:
    inputs = record["inputs"]
    quantity = inputs.get("quantity_count")
    if code(record["id"]) in repeated:
        reason = "duplicate PositionID"
    elif (book := book_status(inputs, True)) is not None:
        return book
    elif code(inputs.get("commodity_code")) == "":
        reason = "Commodity missing"
    elif code(inputs.get("commodity_code")) == GOLD:
        reason = "gold is FX risk; enter it on FXPositions"
    elif sign(inputs.get("direction_code")) == 0:
        reason = "Direction must be Long or Short"
    elif not is_number(quantity):
        reason = "Quantity must be a number"
    elif quantity < 0:
        reason = "Quantity must be >= 0 (Direction gives the sign)"
    elif not is_number(inputs.get("spot_price_amount")):
        reason = "Spot price must be a number"
    elif case.no_currency(code(inputs.get("currency_code"))):
        reason = "currency not in FX_Table"
    else:
        return OK
    return INVALID + reason


def commodities(case: Case, outputs: list[Output]) -> tuple[float, float]:
    """Add the outputs of CALC-05; return the directional charge and the charge on gross positions."""
    records = case.records("DS-08")
    repeated = duplicates(records)
    net: dict[str, float] = {}
    gross: dict[str, float] = {}
    for record in records:
        if text(record["id"]) == "":
            continue
        status = co_status(case, record, repeated)
        outputs.append(row("co_position_status", "DS-08", record, text=status))
        if status != OK:
            continue
        inputs = record["inputs"]
        value = (sign(inputs["direction_code"]) * inputs["quantity_count"] * inputs["spot_price_amount"]
                 * case.rates[code(inputs["currency_code"])])
        name = code(inputs["commodity_code"])
        net[name] = net.get(name, 0.0) + value
        gross[name] = gross.get(name, 0.0) + abs(value)
    directional = sum(CO_DIRECTIONAL * abs(value) for value in net.values())
    on_gross = sum(CO_GROSS * value for value in gross.values())
    outputs.append(total("co_directional_charge", directional))
    outputs.append(total("co_gross_charge", on_gross))
    return directional, on_gross


# ------------------------------------------------------- foreign exchange
def fx_status(case: Case, record: Record, repeated: set[str]) -> str:
    inputs = record["inputs"]
    currency = code(inputs.get("currency_code"))
    amount = inputs.get("net_position_amount")
    if code(record["id"]) in repeated:
        reason = "duplicate PositionID"
    elif (book := book_status(inputs, True)) is not None:
        return book
    elif case.no_currency(currency, gold_allowed=True):
        reason = "currency not in FX_Table"
    elif sign(inputs.get("direction_code")) == 0:
        reason = "Direction must be Long or Short"
    elif not is_number(amount):
        reason = "Amount must be a number"
    elif amount < 0:
        reason = "Amount must be >= 0 (Direction gives the sign)"
    elif currency == case.reporting:
        return EXCLUDED + "reporting currency carries no FX risk"
    else:
        return OK
    return INVALID + reason


def foreign_exchange(case: Case, outputs: list[Output]) -> float:
    """Add the outputs of CALC-06; return the requirement."""
    records = case.records("DS-09")
    repeated = duplicates(records)
    net: dict[str, float] = {}
    for record in records:
        if text(record["id"]) == "":
            continue
        status = fx_status(case, record, repeated)
        outputs.append(row("fx_position_status", "DS-09", record, text=status))
        if status != OK:
            continue
        inputs = record["inputs"]
        currency = code(inputs["currency_code"])
        net[currency] = (net.get(currency, 0.0) + sign(inputs["direction_code"])
                         * inputs["net_position_amount"] * case.rates[currency])
    outputs.extend(by_currency("fx_net_position", currency, value)
                   for currency, value in sorted(net.items()))
    longs = sum(value for currency, value in net.items() if currency != GOLD and value > 0)
    shorts = sum(-value for currency, value in net.items() if currency != GOLD and value < 0)
    position = max(longs, shorts) + abs(net.get(GOLD, 0.0))
    charge = FX_RATE * position
    if case.crr:
        # Article 351: no requirement while the position does not exceed 2% of the total own funds.
        assert is_number(case.own_funds) and case.own_funds > 0, "the CRR regime needs total own funds"
        if position <= FX_DE_MINIMIS * case.own_funds:
            charge = 0.0
    outputs.append(total("fx_overall_net_open_position", position))
    outputs.append(total("fx_charge", charge))
    return charge


# ------------------------------------------------- options and requirement
def option_charges(case: Case, outputs: list[Output]) -> dict[str, float]:
    """Add the outputs of CALC-07; return the charge of each risk class."""
    charges = dict.fromkeys(RISK_CLASSES, 0.0)
    for record in case.records("DS-10"):
        value = record["inputs"].get("charge_amount")
        charge = 0.0 if value is None else float(value)
        charges[code(record["id"])] = charge
        outputs.append({"quantity": "option_charge", "value": charge,
                        "scope": {"dataset": "DS-10", "record": record["id"]}})
    return charges


def requirement(case: Case, outputs: list[Output]) -> None:
    """Add the outputs of CALC-08."""
    scratch: list[Output] = []
    options = option_charges(case, scratch)
    before = {"IR": sum(interest_rate(case, scratch)) + options["IR"],
              "EQ": sum(equity(case, scratch)) + options["EQ"],
              "CO": sum(commodities(case, scratch)) + options["CO"],
              "FX": foreign_exchange(case, scratch) + options["FX"]}
    after = {name: SCALING[name] * value for name, value in before.items()}
    for record in case.records("DS-10"):
        name = code(record["id"])
        scope = {"dataset": "DS-10", "record": record["id"]}
        outputs.append({"quantity": "requirement_before_scaling", "scope": scope, "value": before[name]})
        outputs.append({"quantity": "requirement_after_scaling", "scope": scope, "value": after[name]})
    end = serial(TRANSITIONAL_END)
    assert end is not None
    factor = TRANSITIONAL_MULTIPLIER if case.crr and case.as_of <= end else 1.0
    result = factor * sum(after.values())
    outputs.append(total("transitional_factor", factor))
    outputs.append(total("own_funds_requirement", result))
    outputs.append(total("risk_weighted_assets", RWA_MULTIPLIER * result))


# ------------------------------------------------------------ eligibility
def month_serial(value: object) -> float | None:
    """A month end as a date serial number; None when the value is neither a date nor a number."""
    if isinstance(value, datetime.datetime):
        return (value - datetime.datetime.combine(SERIAL_EPOCH, datetime.time())).total_seconds() / 86400
    if isinstance(value, datetime.date):
        return float((value - SERIAL_EPOCH).days)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return None


def month_end_serial(serial_number: float, months: int) -> float | None:
    """The serial of the last day of the month that lies months after the month of a date serial
    number; None when it is beyond 31 December 9999."""
    day = SERIAL_EPOCH + datetime.timedelta(days=int(serial_number))
    year, month = divmod(day.year * 12 + day.month - 1 + months, 12)
    if year > 9999:
        return None
    last = datetime.date(year, month + 1, calendar.monthrange(year, month + 1)[1])
    return float((last - SERIAL_EPOCH).days)


def month_status(values: Sequence[Any], previous: float | None) -> str:
    """OK, or the first test of CALC-09 that a month fails. previous is the serial of the month end
    of the row entered before it, when that was a date."""
    serial_number = month_serial(values[0])
    if serial_number is None or not is_date(serial_number):
        return MONTH_REASONS[0]
    if serial_number != month_end_serial(serial_number, 0):
        return MONTH_REASONS[1]
    if previous is not None and serial_number != month_end_serial(previous, 1):
        return MONTH_REASONS[2]
    if not is_number(values[1]) or values[1] <= 0:
        return MONTH_REASONS[3]
    if not all(is_number(value) for value in values[2:6]):
        return MONTH_REASONS[4]
    return OK


def eligibility_month(values: Sequence[Any], approach: str, params: Mapping[str, float],
                      missed: list[bool]) -> Record:
    """The figures of a valid month for one approach; missed holds the months before it that did not
    meet the conditions, valid months only, and takes this one."""
    long, short = values[2:4] if approach == "325a" else values[4:6]
    size = abs(long) + abs(short)
    share = params.get(f"EligibilityShare{approach}", ELIGIBILITY_REFERENCE[f"EligibilityShare{approach}"])
    cap = params.get(f"EligibilityCap{approach}", ELIGIBILITY_REFERENCE[f"EligibilityCap{approach}"])
    missed.append(not (size <= share * values[1] and size <= cap))
    run = 0
    for flag in reversed(missed):
        if not flag:
            break
        run += 1
    count = sum(missed[-WINDOW:])
    return {"size": size, "share": size / values[1], "result": "NOT MET" if missed[-1] else "MET",
            "run": run, "count": count, "event": "YES" if run >= EVENT_RUN or count > EVENT_COUNT else "NO"}


def eligibility(rows: Sequence[Sequence[Any]], regime: object, params: Mapping[str, float]) -> Record:
    """CALC-09: the eligibility for the simplified standardised approach (Article 325a) and the
    small trading book derogation (Article 94) of the CRR, from the rows of the Eligibility sheet.

    rows: one tuple (month end, total assets, long 325a, short 325a, long 94, short 94) per row,
    as the sheet holds it. A month end that is None, or an empty text, skips the row. regime is the
    value of the Regime parameter, read without regard to case and blanks: anything but CRR is
    not assessed. params holds EligibilityShare325a, EligibilityCap325a, EligibilityShare94 and
    EligibilityCap94; one that is missing takes its reference value.

    Returns the dict "rows", one item per row: None for a skipped row and under the Basel regime;
    otherwise {"status": OK or an INVALID text, "325a": ..., "94": ...}, where each approach is
    {"size", "share", "result" (MET or NOT MET), "run", "count", "event" (YES or NO)} for a valid
    row and None for an invalid one. The run and the count of a row go over the valid rows, in
    the order of the sheet. Then "status", "restart", "run", "count", "size" and "share", each by
    approach, and "last_month": the month end of the last row entered. Where the status is NOT
    ASSESSED, every item but "status" is None."""
    assessed = code(regime) == "CRR"
    items: list[Record | None] = [None] * len(rows)
    missed: dict[str, list[bool]] = {approach: [] for approach in APPROACHES}
    entered = 0
    rejected = False
    previous: float | None = None
    last: Record = {}
    last_month: object = None
    for index, values in enumerate(rows):
        if not assessed or values[0] is None or text(values[0]) == "":
            continue
        entered += 1
        status = month_status(values, previous)
        serial_number = month_serial(values[0])
        previous = serial_number if serial_number is not None and is_date(serial_number) else None
        items[index] = {"status": status, "325a": None, "94": None}
        last_month = values[0]
        if status != OK:
            rejected = True
            continue
        for approach in APPROACHES:
            figures = eligibility_month(values, approach, params, missed[approach])
            items[index][approach] = figures  # type: ignore[index]
            last[approach] = figures
    result: Record = {"rows": items, "status": {}, "last_month": None}
    for key in ("restart", "run", "count", "size", "share"):
        result[key] = dict.fromkeys(APPROACHES)
    if not assessed or entered == 0 or rejected:
        reason = "Basel regime" if not assessed else "no month entered" if entered == 0 else "a row is invalid"
        result["status"] = dict.fromkeys(APPROACHES, f"{NOT_ASSESSED}: {reason}")
        return result
    result["last_month"] = last_month
    for approach in APPROACHES:
        figures = last[approach]
        result["status"][approach] = ("CEASE" if figures["event"] == "YES"
                                      else "NOT MET" if figures["result"] == "NOT MET" else "MET")
        for key in ("size", "share", "run", "count"):
            result[key][approach] = figures[key]
        result["restart"][approach] = "YES" if len(missed[approach]) >= WINDOW and figures["count"] == 0 else "NO"
    return result


# The keys of a record of DS-11, in the order of the columns of the Eligibility sheet.
MONTH_KEYS = ("month_end_date", "total_assets_amount", "long_325a_amount", "short_325a_amount",
              "long_94_amount", "short_94_amount")


def month_cell(key: str, value: object) -> object:
    """A value of a record of DS-11 as the sheet holds it: an ISO month end as a date."""
    if key == "month_end_date" and isinstance(value, str):
        try:
            return datetime.date.fromisoformat(value)
        except ValueError:
            return value
    return value


def eligibility_outputs(case: Case, outputs: list[Output]) -> None:
    """The outputs of CALC-09: the status and the figures of each month entered, and the status
    and the restart condition of each approach."""
    records = case.records("DS-11")
    rows = [tuple(month_cell(key, record["inputs"].get(key)) for key in MONTH_KEYS)
            for record in records]
    result = eligibility(rows, case.regime, {})
    for record, item in zip(records, result["rows"], strict=True):
        if item is None:
            continue
        outputs.append(row("eligibility_month_status", "DS-11", record, text=item["status"]))
        for approach in APPROACHES:
            figures = item[approach]
            if figures is None:
                continue
            for key in ("size", "share", "run", "count"):
                outputs.append(row(f"eligibility_{key}_{approach}", "DS-11", record,
                                   value=float(figures[key])))
            outputs.append(row(f"eligibility_month_result_{approach}", "DS-11", record,
                               text=figures["result"]))
            outputs.append(row(f"eligibility_event_{approach}", "DS-11", record,
                               text=figures["event"]))
    for approach in APPROACHES:
        outputs.append({"quantity": f"eligibility_status_{approach}", "scope": None,
                        "text": result["status"][approach]})
        if result["restart"][approach] is not None:
            outputs.append({"quantity": f"eligibility_restart_{approach}", "scope": None,
                            "text": result["restart"][approach]})


# ---------------------------------------------------------------- outputs
def row(quantity: str, dataset: str, record: Record, **result: object) -> Output:
    return {"quantity": quantity, "scope": {"dataset": dataset, "record": record["id"]}, **result}


def by_currency(quantity: str, currency: str, value: float) -> Output:
    return {"quantity": quantity, "scope": {"dataset": "DS-05", "record": currency}, "value": value}


def total(quantity: str, value: float) -> Output:
    return {"quantity": quantity, "scope": None, "value": value}


def derive(fixture: dict[str, Any]) -> list[Output]:
    """The outputs of a fixture's calculation."""
    case = Case(fixture)
    outputs: list[Output] = []
    if case.calculation in ("CALC-02", "CALC-03"):
        interest_rate(case, outputs)
    elif case.calculation == "CALC-04":
        equity(case, outputs)
    elif case.calculation == "CALC-05":
        commodities(case, outputs)
    elif case.calculation == "CALC-06":
        foreign_exchange(case, outputs)
    elif case.calculation == "CALC-07":
        option_charges(case, outputs)
    elif case.calculation == "CALC-08":
        requirement(case, outputs)
    elif case.calculation == "CALC-09":
        eligibility_outputs(case, outputs)
    else:
        raise SystemExit(f"ssa.py: no derivation for calculation {case.calculation!r}")
    return outputs


def main(path: str) -> None:
    with open(path, encoding="utf-8") as handle:
        fixture = json.load(handle)
    print(json.dumps({"outputs": derive(fixture)}))


if __name__ == "__main__":
    main(sys.argv[1])
