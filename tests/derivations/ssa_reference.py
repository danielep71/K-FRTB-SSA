"""Sample data and an independent Python re-calculation of the MAR40
simplified standardised approach, used to seed the workbook and to derive
the expected values of the TestCatalogue sheet."""
from datetime import date

AS_OF = date(2026, 9, 30)
REPORTING_CCY = "EUR"
DAYS_PER_YEAR = 365.0

# Units of reporting currency (EUR) per 1 unit of currency. Illustrative.
# USD, GBP, CAD, SEK, NOK and XAU are set so that the FX sample reproduces
# the MAR40.61 Table 9 example exactly (in EUR millions).
FX = [
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
FX_RATE = {c: r for c, r, _ in FX}

# ---- Interest rate positions -------------------------------------------
# (ID, portfolio, instrument, ccy, dir, value, rate type, coupon,
#  final maturity, next repricing, SR category, rating grade, issue ID, comment)
IR = [
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
EQ = [
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
CO = [
    ("CO-001", "Energy", "Brent crude", "Long", 50_000, "bbl", 72.0, "USD", date(2026, 12, 31), "Futures, front month"),
    ("CO-002", "Energy", "Brent crude", "Short", 20_000, "bbl", 72.0, "USD", date(2027, 6, 30), ""),
    ("CO-003", "Energy", "WTI crude", "Long", 10_000, "bbl", 68.0, "USD", date(2026, 12, 31), "Separate commodity from Brent"),
    ("CO-004", "Metals", "Copper", "Long", 200, "t", 9_800.0, "USD", date(2027, 3, 31), "LME"),
    ("CO-005", "Metals", "Copper", "Short", 50, "t", 9_800.0, "USD", date(2026, 12, 31), ""),
    ("CO-006", "Energy", "TTF natural gas", "Short", 100_000, "MWh", 35.0, "EUR", date(2027, 1, 31), ""),
]

# ---- FX positions (reproduce MAR40.61 Table 9 in EUR millions) ----------
# (ID, portfolio, ccy, dir, amount, comment)
FXPOS = [
    ("FX-001", "FX", "CAD", "Long", 80_000_000, "+50m EUR (Table 9: +50)"),
    ("FX-002", "FX", "SEK", "Long", 1_000_000_000, "+100m EUR (Table 9: +100)"),
    ("FX-003", "FX", "GBP", "Long", 125_000_000, "+150m EUR (Table 9: +150)"),
    ("FX-004", "FX", "NOK", "Short", 250_000_000, "-20m EUR (Table 9: -20)"),
    ("FX-005", "FX", "USD", "Short", 200_000_000, "-180m EUR (Table 9: -180)"),
    ("FX-006", "FX", "XAU", "Short", 10_000, "-35m EUR (Table 9: gold -35)"),
]

# ---- Regulatory constants -------------------------------------------------
SCALING = {"IR": 1.30, "EQ": 3.50, "CO": 1.90, "FX": 1.20}
LADDER = [  # band, zone, upper >=3% (years), upper <3%, RW
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
SR = {  # (category, grade, maturity band or 'Any') -> charge
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
}


def years(d):
    return (d - AS_OF).days / DAYS_PER_YEAR


def band_of(t, high):
    col = 2 if high else 3
    return 1 + sum(1 for row in LADDER if row[col] is not None and row[col] < t)


def calc_ir():
    rw = {b: w for b, _, _, _, w in LADDER}
    zone = {b: z for b, z, _, _, _ in LADDER}
    ladder = {}  # (ccy, band) -> [weighted long, weighted short]
    issues = {}  # issue -> [net, rate]
    for (pid, _, _, ccy, d, v, rt, cpn, mat, rep, cat, grd, issue, _) in IR:
        s = (1 if d == "Long" else -1) * v * FX_RATE[ccy]
        t = years(rep if rt == "Floating" else mat)
        high = rt == "Floating" or cpn >= 0.03
        b = band_of(t, high)
        wp = s * rw[b]
        cell = ladder.setdefault((ccy, b), [0.0, 0.0])
        if wp > 0:
            cell[0] += wp
        else:
            cell[1] += -wp
        if cat != "NONE":
            tm = years(mat)
            mb = "<=6m" if tm <= 0.5 else ("6m-24m" if tm <= 2 else ">24m")
            rate = SR.get((cat, grd, mb), SR.get((cat, grd, "Any")))
            key = issue or pid
            if key not in issues:
                issues[key] = [0.0, rate, ccy]
            issues[key][0] += s
    out = {}
    for ccy in sorted({c for c, _ in ladder}):
        vd = 0.0
        zl = {1: 0.0, 2: 0.0, 3: 0.0}
        zs = {1: 0.0, 2: 0.0, 3: 0.0}
        for b in range(1, 16):
            wl, ws = ladder.get((ccy, b), [0.0, 0.0])
            vd += 0.10 * min(wl, ws)
            n = wl - ws
            if n > 0:
                zl[zone[b]] += n
            else:
                zs[zone[b]] += -n
        rates = {1: 0.40, 2: 0.30, 3: 0.30}
        within = sum(rates[z] * min(zl[z], zs[z]) for z in (1, 2, 3))
        z1, z2, z3 = (zl[z] - zs[z] for z in (1, 2, 3))

        def offset(a, b):
            if a * b < 0:
                m = min(abs(a), abs(b))
                return m, a - (m if a > 0 else -m), b - (m if b > 0 else -m)
            return 0.0, a, b

        m12, z1r, z2r = offset(z1, z2)
        m23, z2rr, z3r = offset(z2r, z3)
        m13, _, _ = offset(z1r, z3r)
        between = 0.40 * (m12 + m23) + 1.00 * m13
        net = abs(z1 + z2 + z3)
        gmr = vd + within + between + net
        sr = sum(abs(n) * r for n, r, c in issues.values() if c == ccy)
        out[ccy] = dict(vd=vd, within=within, between=between, net=net, gmr=gmr, sr=sr)
    return out


def calc_eq():
    rates = {"Single name": 0.08, "Index": 0.02}
    nets, markets = {}, {}
    for (_, _, mkt, typ, ident, ccy, d, v, _) in EQ:
        s = (1 if d == "Long" else -1) * v * FX_RATE[ccy]
        nets.setdefault((mkt, typ, ident.upper()), 0.0)
        nets[(mkt, typ, ident.upper())] += s
        markets[mkt] = markets.get(mkt, 0.0) + s
    specific = sum(abs(n) * 0.08 for (m, t, i), n in nets.items() if t == "Single name")
    index = sum(abs(n) * 0.02 for (m, t, i), n in nets.items() if t == "Index")
    general = sum(0.08 * abs(n) for n in markets.values())
    return dict(specific=specific, general=general, index=index)


def calc_co():
    net, gross = {}, {}
    for (_, _, com, d, q, _, p, ccy, _, _) in CO:
        s = (1 if d == "Long" else -1) * q * p * FX_RATE[ccy]
        net[com.upper()] = net.get(com.upper(), 0.0) + s
        gross[com.upper()] = gross.get(com.upper(), 0.0) + abs(s)
    directional = sum(0.15 * abs(n) for n in net.values())
    gross3 = sum(0.03 * g for g in gross.values())
    return dict(directional=directional, gross=gross3)


def calc_fx():
    net = {}
    for (_, _, ccy, d, a, _) in FXPOS:
        if ccy == REPORTING_CCY:
            continue
        net[ccy] = net.get(ccy, 0.0) + (1 if d == "Long" else -1) * a * FX_RATE[ccy]
    longs = sum(v for c, v in net.items() if c != "XAU" and v > 0)
    shorts = sum(-v for c, v in net.items() if c != "XAU" and v < 0)
    gold = abs(net.get("XAU", 0.0))
    onop = max(longs, shorts) + gold
    return dict(longs=longs, shorts=shorts, gold=gold, onop=onop, charge=0.08 * onop)


def calc_all():
    ir = calc_ir()
    eq = calc_eq()
    co = calc_co()
    fx = calc_fx()
    k_ir = sum(v["gmr"] + v["sr"] for v in ir.values())
    k_eq = eq["specific"] + eq["general"] + eq["index"]
    k_co = co["directional"] + co["gross"]
    k_fx = fx["charge"]
    total = SCALING["IR"] * k_ir + SCALING["EQ"] * k_eq + SCALING["CO"] * k_co + SCALING["FX"] * k_fx
    return dict(ir=ir, eq=eq, co=co, fx=fx, k_ir=k_ir, k_eq=k_eq, k_co=k_co, k_fx=k_fx,
                total=total, rwa=12.5 * total)


if __name__ == "__main__":
    import json
    print(json.dumps(calc_all(), indent=1, default=float))
