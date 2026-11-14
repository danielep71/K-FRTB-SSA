"""Compare the output sheets of a workbook run by the VBA engine with the
formula prototype on the same inputs.

Usage: python tools/compare_outputs.py VBA_RUN.xlsx FORMULA.xlsx

Both files must hold calculated values (the formula prototype recalculated,
the VBA run saved after RunSSA). Calculation rows are compared where the
PositionID is not blank, Buckets bucket rows where the bucket name is not
blank (the prototype numbers empty rows; the engine leaves them blank);
other aggregate areas are compared in full. Exits 1 on
any difference.
"""
import datetime
import sys

import openpyxl

TOL = 1e-6


def norm(v):
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, datetime.datetime):
        return (v - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0
    if isinstance(v, (int, float)):
        return float(v)
    return str(v)


def same(a, b):
    a, b = norm(a), norm(b)
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) <= TOL * max(1.0, abs(a), abs(b))
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    return a == b


def compare(ws_v, ws_f, rows, cols, key_col=None, skip=()):
    diffs = []
    for r in rows:
        if key_col and norm(ws_f.cell(r, key_col).value) is None and norm(ws_v.cell(r, key_col).value) is None:
            continue
        for c in cols:
            ref = ws_v.cell(r, c).coordinate
            if ref in skip:
                continue
            a, b = ws_v.cell(r, c).value, ws_f.cell(r, c).value
            if not same(a, b):
                diffs.append(f"{ws_v.title}!{ref}: VBA={a!r} formula={b!r}")
    return diffs


def last_row(ws, col):
    r = ws.max_row
    while r > 1 and ws.cell(r, col).value in (None, ""):
        r -= 1
    return r


def main():
    v = openpyxl.load_workbook(sys.argv[1], data_only=True)
    f = openpyxl.load_workbook(sys.argv[2], data_only=True)
    diffs = []
    checked = 0
    for sh, width in (("IRCalc", 22), ("EQCalc", 12), ("COCalc", 8), ("FXCalc", 7)):
        rows = range(5, max(last_row(v[sh], 3), last_row(f[sh], 3)) + 1)
        d = compare(v[sh], f[sh], rows, range(3, 3 + width), key_col=3)
        diffs += d
        checked += len(rows) * width
    L = v["IR_Ladder"]
    diffs += compare(L, f["IR_Ladder"], range(6, 306), range(3, 16))
    checked += 300 * 13
    Z = v["IR_Zones"]
    diffs += compare(Z, f["IR_Zones"], range(6, 27), range(3, 33))
    checked += 21 * 30
    B = v["Buckets"]
    diffs += compare(B, f["Buckets"], range(7, 38), range(3, 13), key_col=4)
    tot = 37
    cs = tot + 3
    diffs += compare(B, f["Buckets"], range(cs + 2, cs + 33), range(3, 13), key_col=4)
    fs = cs + 2 + 30 + 3
    diffs += compare(B, f["Buckets"], range(fs + 2, fs + 25), range(3, 12))
    checked += 31 * 10 * 2 + 23 * 9
    R = v["Results"]
    diffs += compare(R, f["Results"], range(7, 12), range(5, 16))
    diffs += compare(R, f["Results"], range(14, 18), [5])
    checked += 5 * 11 + 4
    C = v["Checks"]
    diffs += compare(C, f["Checks"], range(7, 20), [6, 7])
    diffs += compare(C, f["Checks"], [4], [4, 6])
    checked += 13 * 2 + 2
    T = v["TestCatalogue"]
    diffs += compare(T, f["TestCatalogue"], range(7, 26), [9, 13])
    checked += 19 * 2
    print(f"cells compared: {checked}; differences: {len(diffs)}")
    for d in diffs[:60]:
        print("  " + d)
    sys.exit(1 if diffs else 0)


if __name__ == "__main__":
    main()
