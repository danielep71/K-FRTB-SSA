"""Rebuild and verify K-FRTB-SSA end to end without Excel (Linux, LibreOffice).

Usage: python tools/verify.py

1. The repository's static VBA checks (format, source, public API, dependencies).
2. Build the formula prototype and the macro-free template.
3. Recalculate the prototype in LibreOffice; all TestCatalogue tests PASS.
4. Run RunSSA (VBA engine, LibreOffice VBA mode) on the template and compare
   every output area with the prototype: no difference allowed.
5. Same on a negative scenario (rejected rows, MIXED issue, changed
   constant, option charge) applied to both workbooks, findings included.
6. A run with unusable Params values must STOP and list them.
7. TEST_Suite.Test_RunAll: every assertion passes.

Exits 1 at the first failed step. Outputs go to build/.
"""
import datetime
import pathlib
import subprocess
import sys

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
BUILD = ROOT / "test-results" / "prototype"
TOOLS = ROOT / "tools" / "prototype"
MODULES = [ROOT / "src" / "core" / f"{m}.bas" for m in (
    "CORE_Config", "CORE_Util", "CORE_Parameters", "CORE_InterestRate", "CORE_Equity",
    "CORE_Commodity", "CORE_ForeignExchange", "CORE_Engine")] + [
    ROOT / "src" / "modules" / "FRTBSSA_Formulas.bas",
    ROOT / "src" / "workbook" / "M_Main.bas",
    ROOT / "tests" / "modules" / "TEST_Suite.bas"]


def step(title):
    print(f"\n== {title}")


def call(*args):
    out = subprocess.run([sys.executable, *map(str, args)], capture_output=True, text=True)
    print(out.stdout.strip())
    if out.returncode != 0:
        print(out.stderr.strip())
    return out


def harness(workbook, out, macro):
    res = call(TOOLS / "lo_vba_harness.py", workbook, out, macro, *MODULES)
    line = [ln for ln in res.stdout.splitlines() if ln.startswith("RESULT:")]
    return line[-1][8:] if line else "FAILED harness"


def fail(message):
    print(f"\nVERIFY FAILED: {message}")
    sys.exit(1)


def negative(src, dst):
    """Apply the same invalid rows and parameter changes to a workbook."""
    wb = openpyxl.load_workbook(src)

    def put(ws, r, values):
        for k, v in enumerate(values):
            ws.cell(r, 3 + k).value = v

    d = datetime.date
    ir = wb["IRPositions"]
    put(ir, 17, ["IR-BAD1", "T", "Bond", "EUR", "Long", 1e6, "Fixed", None, d(2030, 1, 1), None, "OTHER", "BB+ to BB-", ""])
    put(ir, 18, ["IR-BAD2", "T", "Bond", "ZAR", "Long", 1e6, "Fixed", 0.05, d(2030, 1, 1), None, "NONE", "", ""])
    put(ir, 19, ["ir-001", "T", "Bond", "EUR", "Long", 1e6, "Fixed", 0.05, d(2030, 1, 1), None, "NONE", "", ""])
    put(ir, 20, ["IR-MIX", "T", "Bond", "EUR", "Long", 2e6, "Fixed", 0.04, d(2034, 3, 1), None, "QUAL", "A+ to BBB-", "BTP 4.00% 03/34"])
    put(ir, 21, ["IR-FRN", "T", "FRN", "EUR", "Short", 3e6, "floating", None, d(2029, 1, 1), d(2027, 1, 1), "GOV", "AAA to AA-", ""])
    put(ir, 22, ["IR-OTH", "T", "Bond", "EUR", "Long", 1e6, "Fixed", 0.05, d(2030, 1, 1), None, "OTHER", "A+ to BBB-", ""])
    eq = wb["EQPositions"]
    put(eq, 16, ["EQ-BAD", "T", "IT", "Basket", "B", "EUR", "Long", 1e6])
    put(eq, 17, ["EQ-NEG", "T", "FR", "Single name", "BNP", "EUR", "Short", 2e6])
    co = wb["COPositions"]
    put(co, 11, ["CO-BAD", "T", "Gold bar", "Long", "ten", "oz", 3500, "EUR"])
    put(co, 12, ["CO-007", "T", "Wheat", "Short", 1000, "t", 250, "EUR"])
    fx = wb["FXPositions"]
    put(fx, 11, ["FX-007", "FX", "EUR", "Long", 5e6])
    put(fx, 12, ["FX-008", "FX", "JPY", "Long", -5])
    put(fx, 13, ["FX-009", "FX", "CHF", "Long", 4e6])
    assert wb["Params"]["C10"].value == "ScalingIR"
    wb["Params"]["E10"] = 1.5
    wb["OptionCharges"]["F8"] = 1000
    wb.save(dst)


def findings(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Checks"]
    start = next(r for r in range(20, 60) if ws.cell(r, 3).value == "#") + 1
    return {(ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value, ws.cell(r, 9).value)
            for r in range(start, start + 100) if ws.cell(r, 4).value}


def main():
    BUILD.mkdir(parents=True, exist_ok=True)
    proto, template = BUILD / "K-FRTB-SSA_Prototype.xlsx", ROOT / "src" / "workbook" / "Template.xlsx"

    step("1. Static source checks")
    for check in ("format_vba.py", "check_vba_source.py", "check_vba_public_api.py",
                  "check_vba_dependencies.py"):
        args = [ROOT / "tools" / check] + (["--check"] if check == "format_vba.py" else [])
        if call(*args).returncode:
            fail(f"static checks: {check}")

    step("2. Build the formula prototype and the template")
    call(TOOLS / "build_prototype.py", proto)
    call(TOOLS / "build_prototype.py", template, "--template")

    step("3. Recalculate the prototype")
    harness(proto, BUILD / "prototype_values.xlsx", "-")
    summary = openpyxl.load_workbook(BUILD / "prototype_values.xlsx", data_only=True)["TestCatalogue"]["C4"].value
    print(summary)
    if not summary or not summary.startswith("19 of 19"):
        fail("prototype TestCatalogue")

    step("4. RunSSA on the sample data vs the prototype")
    result = harness(template, BUILD / "vba_sample.xlsx", "M_Main.RunSSA_Silent")
    print(result)
    if not result.startswith("RESULT=OK"):
        fail("sample run")
    if call(TOOLS / "compare_outputs.py", BUILD / "vba_sample.xlsx", BUILD / "prototype_values.xlsx").returncode:
        fail("sample comparison")

    step("5. Negative scenario on both workbooks")
    negative(template, BUILD / "neg_template.xlsx")
    negative(proto, BUILD / "neg_prototype.xlsx")
    harness(BUILD / "neg_prototype.xlsx", BUILD / "neg_prototype_values.xlsx", "-")
    result = harness(BUILD / "neg_template.xlsx", BUILD / "vba_negative.xlsx", "M_Main.RunSSA_Silent")
    print(result)
    if not result.startswith("RESULT=INCOMPLETE"):
        fail("negative run outcome")
    if call(TOOLS / "compare_outputs.py", BUILD / "vba_negative.xlsx", BUILD / "neg_prototype_values.xlsx").returncode:
        fail("negative comparison")
    vba_f, proto_f = findings(BUILD / "vba_negative.xlsx"), findings(BUILD / "neg_prototype_values.xlsx")
    print(f"findings: VBA {len(vba_f)}, prototype {len(proto_f)}")
    if vba_f != proto_f:
        fail(f"findings differ: {sorted(vba_f ^ proto_f)}")

    step("6. Unusable Params values stop the run")
    wb = openpyxl.load_workbook(template)
    wb["Params"]["E6"] = "not a date"
    wb["OptionCharges"]["F5"] = "abc"
    wb.save(BUILD / "stop_template.xlsx")
    result = harness(BUILD / "stop_template.xlsx", BUILD / "vba_stop.xlsx", "M_Main.RunSSA_Silent")
    print(result)
    stopped = findings(BUILD / "vba_stop.xlsx")
    print(sorted(stopped))
    if not result.startswith("RESULT=STOPPED; errors=2") or len(stopped) != 2:
        fail("stopped run")

    step("7. TEST_Suite.Test_RunAll")
    result = harness(template, BUILD / "vba_tests.xlsx", "TEST_Suite.Test_RunAll")
    print(result)
    if not result.startswith("PASS=") or "FAIL=0" not in result:
        fail("VBA test suite")

    print("\nVERIFY PASSED")


if __name__ == "__main__":
    main()
