#!/usr/bin/env python3
"""Sample data of the workbook and a Python re-calculation of them.

The template generator, build_prototype.py, takes the sample positions, the
FX rates and the MAR40 tables from here, and the expected values of the
template's TestCatalogue sheet from calc_all. The re-calculation follows the
MAR40 simplified standardised approach as the workbook's README sheet states
it, and is written apart from the VBA engine: it shares no code with it.

It is not a derivation script in the sense of docs/methodology/TEST_CASES.md:
it reads no fixture, and its values are illustrative, because the project
derived them.

Usage: python tools/prototype/ssa_reference.py
Prints the re-calculation of the sample data as JSON.
"""

from __future__ import annotations

import json
import sys
from datetime import date
from typing import TypedDict

# (ID, portfolio, instrument, currency, direction, value, rate type, coupon,
#  final maturity, next repricing, SR category, rating grade, issue ID, comment)
IrRow = tuple[str, str, str, str, str, float, str, float | None, date, date | None,
              str, str, str, str]
# (ID, portfolio, market, type, identifier, currency, direction, value, comment)
EqRow = tuple[str, str, str, str, str, str, str, float, str]
# (ID, portfolio, commodity, direction, quantity, unit, spot, currency, maturity, comment)
CoRow = tuple[str, str, str, str, float, str, float, str, date, str]
# (ID, portfolio, currency, direction, amount, comment)
FxRow = tuple[str, str, str, str, float, str]


class IrCurrency(TypedDict):
    """Interest rate charges of one currency, in the reporting currency."""
    vd: float
    within: float
    between: float
    net: float
    gmr: float
    sr: float


class Equity(TypedDict):
    specific: float
    general: float
    index: float


class Commodity(TypedDict):
    directional: float
    gross: float


class ForeignExchange(TypedDict):
    longs: float
    shorts: float
    gold: float
    onop: float
    charge: float


class Requirement(TypedDict):
    """Every class, its requirement before scaling, the total and the RWA."""
    ir: dict[str, IrCurrency]
    eq: Equity
    co: Commodity
    fx: ForeignExchange
    k_ir: float
    k_eq: float
    k_co: float
    k_fx: float
    total: float
    rwa: float


AS_OF = date(2026, 9, 30)
REPORTING_CCY = "EUR"
DAYS_PER_YEAR = 365.0

# Units of reporting currency (EUR) per 1 unit of currency. Illustrative.
# USD, GBP, CAD, SEK, NOK and XAU are set so that the FX sample reproduces
# the MAR40.61 Table 9 example exactly (in EUR millions).
FX: list[tuple[str, float, str]] = [
    ("EUR", 1.0, "Euro (reporting currency)"),
    ("USD", 0.9, "US dollar"),
    ("GBP", 1.2, "Pound sterling"),
    ("CHF", 1.07, "Swiss franc"),
    ("JPY", 0.0058, "Japanese yen"),
    ("CAD", 0.625, "Canadian dollar"),
    ("SEK", 0.1, "Swedish krona"),
    ("NOK", 0.08, "Norwegian krone"),
    ("DKK", 0.134, "Danish krone"),
    ("AUD", 0.56, "Australian dollar"),
    ("PLN", 0.235, "Polish zloty"),
    ("CZK", 0.041, "Czech koruna"),
    ("XAU", 3500.0, "Gold, per troy ounce (FX risk, MAR40.60)"),
]
FX_RATE: dict[str, float] = {c: r for c, r, _ in FX}

# ---- Interest rate positions -------------------------------------------
# (ID, portfolio, instrument, ccy, dir, value, rate type, coupon,
#  final maturity, next repricing, SR category, rating grade, issue ID, comment)
IR: list[IrRow] = [
    ("IR-001", "Govies", "Bond", "EUR", "Long", 25_000_000, "Fixed", 0.040, date(2034, 3, 1), None, "GOV", "A+ to BBB-", "BTP 4.00% 03/34", "Italian BTP, BBB+"),
    ("IR-002", "Govies", "Forward leg", "EUR", "Short", 5_000_000, "Fixed", 0.040, date(2034, 3, 1), None, "GOV", "A+ to BBB-", "BTP 4.00% 03/34", "Forward sale of the same BTP: nets specific risk"),
    ("IR-003", "Govies", "Bond", "EUR", "Long", 15_000_000, "Fixed", 0.025, date(2030, 2, 15), None, "GOV", "AAA to AA-", "DBR 2.50% 02/30", "Bund, low coupon (< 3%)"),
    ("IR-004", "Rates", "Swap leg", "EUR", "Short", 20_000_000, "Fixed", 0.028, date(2036, 9, 30), None, "NONE", "", "", "10y payer IRS, fixed leg"),
    ("IR-005", "Rates", "Swap leg", "EUR", "Long", 20_000_000, "Floating", None, date(2036, 9, 30), date(2027, 3, 31), "NONE", "", "", "10y payer IRS, floating leg (6M Euribor)"),
    ("IR-006", "Credit", "Bond", "EUR", "Long", 8_000_000, "Fixed", 0.045, date(2029, 6, 15), None, "QUAL", "A+ to BBB-", "XS-CORP-A 4.50% 06/29", "Investment-grade corporate"),
    ("IR-007", "Credit", "Bond", "EUR", "Long", 3_000_000, "Fixed", 0.0625, date(2031, 5, 15), None, "OTHER", "BB+ to BB-", "XS-CORP-HY 6.25% 05/31", "High-yield corporate"),
    ("IR-008", "Credit", "FRN", "EUR", "Long", 5_000_000, "Floating", None, date(2028, 12, 15), date(2026, 12, 15), "QUAL", "AAA to AA-", "XS-SSA-FRN 12/28", "Supranational FRN, quarterly reset"),
    ("IR-009", "Govies", "Bond", "USD", "Long", 12_000_000, "Fixed", 0.0475, date(2055, 5, 15), None, "GOV", "AAA to AA-", "UST 4.75% 05/55", "30y US Treasury"),
    ("IR-010", "Govies", "Bond", "USD", "Short", 18_000_000, "Fixed", 0.03875, date(2028, 9, 30), None, "GOV", "AAA to AA-", "UST 3.875% 09/28", "2y US Treasury, short"),
    ("IR-011", "Credit", "Bond", "USD", "Long", 2_000_000, "Fixed", 0.075, date(2030, 4, 1), None, "OTHER", "B+ to B-", "US-CORP-B 7.50% 04/30", "Single-B corporate"),
    ("IR-012", "Govies", "Bond", "GBP", "Long", 6_000_000, "Fixed", 0.0125, date(2041, 10, 22), None, "GOV", "AAA to AA-", "UKT 1.25% 10/41", "Gilt, low coupon"),
]

# ---- Equity positions ----------------------------------------------------
# (ID, portfolio, market, type, identifier, ccy, dir, value, comment)
EQ: list[EqRow] = [
    ("EQ-001", "Equity IT", "IT", "Single name", "ENI", "EUR", "Long", 4_000_000, ""),
    ("EQ-002", "Equity IT", "IT", "Single name", "ISP", "EUR", "Long", 6_000_000, ""),
    ("EQ-003", "Equity IT", "IT", "Single name", "UCG", "EUR", "Short", 2_000_000, ""),
    ("EQ-004", "Equity IT", "IT", "Index", "FTSE MIB", "EUR", "Short", 5_000_000, "Index future, delta-equivalent"),
    ("EQ-005", "Equity DE", "DE", "Single name", "SAP", "EUR", "Long", 5_000_000, ""),
    ("EQ-006", "Equity DE", "DE", "Single name", "SIE", "EUR", "Long", 3_000_000, ""),
    ("EQ-007", "Equity DE", "DE", "Index", "DAX", "EUR", "Short", 4_000_000, "Index future"),
    ("EQ-008", "Equity US", "US", "Single name", "AAPL", "USD", "Long", 3_000_000, ""),
    ("EQ-009", "Equity US", "US", "Single name", "MSFT", "USD", "Short", 1_500_000, ""),
    ("EQ-010", "Equity US", "US", "Index", "S&P 500", "USD", "Long", 2_000_000, "Index future"),
    ("EQ-011", "Equity IT", "IT", "Single name", "ENI", "EUR", "Short", 1_000_000, "Same stock as EQ-001: nets"),
]

# ---- Commodity positions -------------------------------------------------
# (ID, portfolio, commodity, dir, quantity, unit, spot, ccy, maturity, comment)
CO: list[CoRow] = [
    ("CO-001", "Energy", "Brent crude", "Long", 50_000, "bbl", 72.0, "USD", date(2026, 12, 31), "Futures, front month"),
    ("CO-002", "Energy", "Brent crude", "Short", 20_000, "bbl", 72.0, "USD", date(2027, 6, 30), ""),
    ("CO-003", "Energy", "WTI crude", "Long", 10_000, "bbl", 68.0, "USD", date(2026, 12, 31), "Separate commodity from Brent"),
    ("CO-004", "Metals", "Copper", "Long", 200, "t", 9_800.0, "USD", date(2027, 3, 31), "LME"),
    ("CO-005", "Metals", "Copper", "Short", 50, "t", 9_800.0, "USD", date(2026, 12, 31), ""),
    ("CO-006", "Energy", "TTF natural gas", "Short", 100_000, "MWh", 35.0, "EUR", date(2027, 1, 31), ""),
]

# ---- FX positions (reproduce MAR40.61 Table 9 in EUR millions) ----------
# (ID, portfolio, ccy, dir, amount, comment)
FXPOS: list[FxRow] = [
    ("FX-001", "FX", "CAD", "Long", 80_000_000, "+50m EUR (Table 9: +50)"),
    ("FX-002", "FX", "SEK", "Long", 1_000_000_000, "+100m EUR (Table 9: +100)"),
    ("FX-003", "FX", "GBP", "Long", 125_000_000, "+150m EUR (Table 9: +150)"),
    ("FX-004", "FX", "NOK", "Short", 250_000_000, "-20m EUR (Table 9: -20)"),
    ("FX-005", "FX", "USD", "Short", 200_000_000, "-180m EUR (Table 9: -180)"),
    ("FX-006", "FX", "XAU", "Short", 10_000, "-35m EUR (Table 9: gold -35)"),
]

# ---- Regulatory constants -------------------------------------------------
SCALING: dict[str, float] = {"IR": 1.30, "EQ": 3.50, "CO": 1.90, "FX": 1.20}
# Band, zone, upper bound in years for a coupon of 3% or more, upper bound
# for a lower coupon, risk weight.
LADDER: list[tuple[int, int, float | None, float | None, float]] = [
    (1, 1, 1 / 12, 1 / 12, 0.0000),
    (2, 1, 0.25, 0.25, 0.0020),
    (3, 1, 0.5, 0.5, 0.0040),
    (4, 1, 1.0, 1.0, 0.0070),
    (5, 2, 2.0, 1.9, 0.0125),
    (6, 2, 3.0, 2.8, 0.0175),
    (7, 2, 4.0, 3.6, 0.0225),
    (8, 3, 5.0, 4.3, 0.0275),
    (9, 3, 7.0, 5.7, 0.0325),
    (10, 3, 10.0, 7.3, 0.0375),
    (11, 3, 15.0, 9.3, 0.0450),
    (12, 3, 20.0, 10.6, 0.0525),
    (13, 3, None, 12.0, 0.0600),
    (14, 3, None, 20.0, 0.0800),
    (15, 3, None, None, 0.1250),
]
# (category, grade, maturity band or 'Any') -> charge
SR: dict[tuple[str, str, str], float] = {
    ("GOV", "AAA to AA-", "Any"): 0.0,
    ("GOV", "A+ to BBB-", "<=6m"): 0.0025,
    ("GOV", "A+ to BBB-", "6m-24m"): 0.01,
    ("GOV", "A+ to BBB-", ">24m"): 0.016,
    ("GOV", "BB+ to BB-", "Any"): 0.08,
    ("GOV", "B+ to B-", "Any"): 0.08,
    ("GOV", "Below B-", "Any"): 0.12,
    ("GOV", "Unrated", "Any"): 0.08,
    ("QUAL", "AAA to AA-", "<=6m"): 0.0025,
    ("QUAL", "AAA to AA-", "6m-24m"): 0.01,
    ("QUAL", "AAA to AA-", ">24m"): 0.016,
    ("QUAL", "A+ to BBB-", "<=6m"): 0.0025,
    ("QUAL", "A+ to BBB-", "6m-24m"): 0.01,
    ("QUAL", "A+ to BBB-", ">24m"): 0.016,
    ("OTHER", "BB+ to BB-", "Any"): 0.08,
    ("OTHER", "B+ to B-", "Any"): 0.12,
    ("OTHER", "Below B-", "Any"): 0.12,
    ("OTHER", "Unrated", "Any"): 0.08,
    # CRR Article 336, table 1, by the risk weight of the credit risk standardised approach. A weight of 10%
    # (Article 336(3)) is half of the second category; a weight of 75%, which the table does not name, is read
    # as the category of 100%, by decision (#79).
    ("CRR", "RW 0%", "Any"): 0.0,
    ("CRR", "RW 10%", "<=6m"): 0.00125,
    ("CRR", "RW 10%", "6m-24m"): 0.005,
    ("CRR", "RW 10%", ">24m"): 0.008,
    ("CRR", "RW 20%", "<=6m"): 0.0025,
    ("CRR", "RW 20%", "6m-24m"): 0.01,
    ("CRR", "RW 20%", ">24m"): 0.016,
    ("CRR", "RW 50%", "<=6m"): 0.0025,
    ("CRR", "RW 50%", "6m-24m"): 0.01,
    ("CRR", "RW 50%", ">24m"): 0.016,
    ("CRR", "RW 75%", "Any"): 0.08,
    ("CRR", "RW 100%", "Any"): 0.08,
    ("CRR", "RW 150%", "Any"): 0.12,
}


# Netting keys of specific risk: an entered issue ID and a position without
# one are in separate namespaces, as in the engine, so they can never net.
KEY_ISSUE = "ISSUE:"
KEY_POSITION = "POS:"


def sign(direction: str) -> int:
    return 1 if direction == "Long" else -1


def years(day: date) -> float:
    return (day - AS_OF).days / DAYS_PER_YEAR


def band_of(maturity: float, high: bool) -> int:
    """The time band of a residual maturity in years; upper bounds are inclusive."""
    column = 2 if high else 3
    bounds = [row[column] for row in LADDER]
    return 1 + sum(1 for bound in bounds if bound is not None and bound < maturity)


def offset(first: float, second: float) -> tuple[float, float, float]:
    """The matched amount of two zone positions of opposite sign, and what is left of each."""
    if first * second < 0:
        matched = min(abs(first), abs(second))
        return (matched, first - (matched if first > 0 else -matched),
                second - (matched if second > 0 else -matched))
    return 0.0, first, second


def issue_key(position_id: str, issue: str) -> str:
    return KEY_ISSUE + issue.upper() if issue else KEY_POSITION + position_id


def ladder_charges(weighted: dict[int, list[float]], zone: dict[int, int],
                   zone13: float = 1.00) -> dict[str, float]:
    """Vertical and horizontal disallowances and the net position of one currency's ladder.

    ``zone13`` is the disallowance between zones 1 and 3: 1.00 under MAR40.28, 1.50 under the CRR."""
    vertical = 0.0
    longs = {1: 0.0, 2: 0.0, 3: 0.0}
    shorts = {1: 0.0, 2: 0.0, 3: 0.0}
    for band in range(1, 16):
        long_, short = weighted.get(band, [0.0, 0.0])
        vertical += 0.10 * min(long_, short)
        net = long_ - short
        if net > 0:
            longs[zone[band]] += net
        else:
            shorts[zone[band]] += -net
    rates = {1: 0.40, 2: 0.30, 3: 0.30}
    within = sum(rates[z] * min(longs[z], shorts[z]) for z in (1, 2, 3))
    zone1, zone2, zone3 = (longs[z] - shorts[z] for z in (1, 2, 3))
    matched12, left1, left2 = offset(zone1, zone2)
    matched23, _, left3 = offset(left2, zone3)
    matched13, _, _ = offset(left1, left3)
    between = 0.40 * (matched12 + matched23) + zone13 * matched13
    net_position = abs(zone1 + zone2 + zone3)
    return {"vd": vertical, "within": within, "between": between, "net": net_position,
            "gmr": vertical + within + between + net_position}


def calc_ir(zone13: float = 1.00) -> dict[str, IrCurrency]:
    weight = {band: w for band, _, _, _, w in LADDER}
    zone = {band: z for band, z, _, _, _ in LADDER}
    ladder: dict[str, dict[int, list[float]]] = {}   # currency -> band -> [long, short]
    issues: dict[str, tuple[float, float, str]] = {}  # key -> (net, rate, currency)
    for (pid, _, _, ccy, direction, value, rate_type, coupon, maturity, repricing,
         category, grade, issue, _) in IR:
        signed = sign(direction) * value * FX_RATE[ccy]
        floating = rate_type == "Floating"
        slot = repricing if floating and repricing is not None else maturity
        high = floating or (coupon is not None and coupon >= 0.03)
        band = band_of(years(slot), high)
        weighted = signed * weight[band]
        cell = ladder.setdefault(ccy, {}).setdefault(band, [0.0, 0.0])
        if weighted > 0:
            cell[0] += weighted
        else:
            cell[1] += -weighted
        if category != "NONE":
            to_maturity = years(maturity)
            sr_band = "<=6m" if to_maturity <= 0.5 else ("6m-24m" if to_maturity <= 2 else ">24m")
            rate = SR.get((category, grade, sr_band), SR.get((category, grade, "Any"), 0.0))
            key = issue_key(pid, issue)
            net, first_rate, first_ccy = issues.get(key, (0.0, rate, ccy))
            issues[key] = (net + signed, first_rate, first_ccy)
    out: dict[str, IrCurrency] = {}
    for ccy in sorted(ladder):
        charges = ladder_charges(ladder[ccy], zone, zone13)
        specific = sum(abs(net) * rate for net, rate, issue_ccy in issues.values()
                       if issue_ccy == ccy)
        out[ccy] = {"vd": charges["vd"], "within": charges["within"],
                    "between": charges["between"], "net": charges["net"],
                    "gmr": charges["gmr"], "sr": specific}
    return out


def calc_eq(index_rate: float = 0.02) -> Equity:
    nets: dict[tuple[str, str, str], float] = {}
    markets: dict[str, float] = {}
    for (_, _, market, kind, identifier, ccy, direction, value, _) in EQ:
        signed = sign(direction) * value * FX_RATE[ccy]
        key = (market, kind, identifier.upper())
        nets[key] = nets.get(key, 0.0) + signed
        markets[market] = markets.get(market, 0.0) + signed
    specific = sum(abs(net) * 0.08 for (_, kind, _), net in nets.items() if kind == "Single name")
    index = sum(abs(net) * index_rate for (_, kind, _), net in nets.items() if kind == "Index")
    general = sum(0.08 * abs(net) for net in markets.values())
    return {"specific": specific, "general": general, "index": index}


def calc_co() -> Commodity:
    net: dict[str, float] = {}
    gross: dict[str, float] = {}
    for (_, _, commodity, direction, quantity, _, price, ccy, _, _) in CO:
        signed = sign(direction) * quantity * price * FX_RATE[ccy]
        name = commodity.upper()
        net[name] = net.get(name, 0.0) + signed
        gross[name] = gross.get(name, 0.0) + abs(signed)
    return {"directional": sum(0.15 * abs(value) for value in net.values()),
            "gross": sum(0.03 * value for value in gross.values())}


def calc_fx() -> ForeignExchange:
    net: dict[str, float] = {}
    for (_, _, ccy, direction, amount, _) in FXPOS:
        if ccy == REPORTING_CCY:
            continue
        net[ccy] = net.get(ccy, 0.0) + sign(direction) * amount * FX_RATE[ccy]
    longs = sum(value for ccy, value in net.items() if ccy != "XAU" and value > 0)
    shorts = sum(-value for ccy, value in net.items() if ccy != "XAU" and value < 0)
    gold = abs(net.get("XAU", 0.0))
    onop = max(longs, shorts) + gold
    return {"longs": longs, "shorts": shorts, "gold": gold, "onop": onop, "charge": 0.08 * onop}


def calc_all(zone13: float = 1.00, index_rate: float = 0.02) -> Requirement:
    ir = calc_ir(zone13)
    eq = calc_eq(index_rate)
    co = calc_co()
    fx = calc_fx()
    k_ir = sum(currency["gmr"] + currency["sr"] for currency in ir.values())
    k_eq = eq["specific"] + eq["general"] + eq["index"]
    k_co = co["directional"] + co["gross"]
    k_fx = fx["charge"]
    total = (SCALING["IR"] * k_ir + SCALING["EQ"] * k_eq + SCALING["CO"] * k_co
             + SCALING["FX"] * k_fx)
    return {"ir": ir, "eq": eq, "co": co, "fx": fx, "k_ir": k_ir, "k_eq": k_eq,
            "k_co": k_co, "k_fx": k_fx, "total": total, "rwa": 12.5 * total}


def main() -> int:
    print(json.dumps(calc_all(), indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
