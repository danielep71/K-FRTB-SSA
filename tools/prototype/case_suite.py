#!/usr/bin/env python3
"""Write the VBA suite that runs the registered test cases in the workbook.

Usage: python tools/prototype/case_suite.py [--write]

It reads each JSON fixture of tests/fixtures/ whose calculation is one of the
project's, CALC-02 to CALC-08, with its expected file of tests/expected/, and
writes tests/modules/TEST_Cases.bas: a suite of the VBA test harness with one
case per test case. A case types the fixture's records on the input sheets,
runs the calculation through RunSSA_Silent, compares every expected output
with the cell that holds it, a number within the output's tolerance and a
text exactly, records every number as a result, and puts the inputs back as
they were, which it asserts as inputs-restored. The inputs and the expected
values are literals of the module, so Excel reads no file, and the cells are
those that tools/prototype/run_cases.py reads: both use the same table of
outputs.

Without an option it compares the tracked module with what it would write,
and exits 0 when they are equal and 1 when they differ. With --write it
writes the module. Nothing is compiled or run: the module runs in Excel, and
tools/devbuild/lo_smoke.py runs it in LibreOffice as an early warning.
"""

from __future__ import annotations

import argparse
import datetime
import json
import re
import sys
from decimal import Decimal
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import format_vba  # noqa: E402
import run_cases  # noqa: E402

ROOT = run_cases.ROOT
MODULE = ROOT / "tests" / "modules" / "TEST_Cases.bas"
SUITE = "cases"
# The date of the fixed text of the module: its helpers and entry points.
WRITTEN = "2026-11-17"
RULE = "'" + "=" * 78
THIN = "'" + "-" * 78
WIDTH = 96
FIRST_COLUMN = run_cases.FIRST_COLUMN
# Dataset: the constant of CORE_Config that names the header of its input sheet.
INPUT_HEADERS = {"DS-06": "NM_IR_POS_HEADER", "DS-07": "NM_EQ_POS_HEADER",
                 "DS-08": "NM_CO_POS_HEADER", "DS-09": "NM_FX_POS_HEADER"}
# Calculation sheet: the constant of CORE_Config that names its header.
CALC_HEADERS = {"IRCalc": "NM_IR_CALC_HEADER", "EQCalc": "NM_EQ_CALC_HEADER",
                "COCalc": "NM_CO_CALC_HEADER", "FXCalc": "NM_FX_CALC_HEADER"}
# Quantity of the whole calculation that a workbook name holds: how the module names it.
NAMES = {"FX_ONOP": "NM_FX_ONOP", "SSA_Requirement": '"SSA_Requirement"',
         "SSA_RWA": '"SSA_RWA"'}


class SuiteError(ValueError):
    """A fixture or an expected file the module cannot be written from."""


def centred(name: str) -> str:
    return "'" + " " * ((79 - len(name)) // 2) + name


def header(name: str, sections: list[tuple[str, str]], updated: str) -> list[str]:
    lines = ["'", RULE, centred(name), THIN]
    for index, (title, text) in enumerate([*sections, ("UPDATED", updated)]):
        if index:
            lines.append("'")
        lines.append(f"' {title}")
        lines += [f"'   {line}".rstrip() for line in text.strip("\n").split("\n")]
    return [*lines, RULE, "'"]


def banner(title: str) -> list[str]:
    return ["", THIN, centred(title), THIN, ""]


def procedure(signature: list[str], sections: list[tuple[str, str]],
              dims: list[tuple[str, str, str]], body: list[str], updated: str = WRITTEN
              ) -> list[str]:
    """A procedure in the house style. A line of ``body`` that is a label stays at
    the margin; any other is indented, with its own indentation kept."""
    match = re.search(r"(Sub|Function) (\w+)", signature[0])
    assert match is not None
    lines = [*signature, *header(match.group(2), sections, updated)]
    if dims:
        names = max(len(name) for name, _, _ in dims)
        kinds = max(len(kind) for _, kind, _ in dims)
        lines += [f"    Dim {name.ljust(names)}   As {kind.ljust(kinds)}   '{comment}"
                  for name, kind, comment in dims]
        lines.append("")
    for line in body:
        if not line.strip():
            lines.append("")
        elif re.fullmatch(r"\w+:", line):
            lines.append(line)
        else:
            lines.append("        " + line)
    return [*lines, "", f"End {match.group(1)}", "", ""]


def wrapped(statement: str) -> list[str]:
    """A statement on lines of at most WIDTH characters after the indentation,
    broken after a comma outside a string literal."""
    limit = WIDTH - 8
    if len(statement) <= limit:
        return [statement]
    breaks = []
    quoted = False
    for index, character in enumerate(statement):
        if character == '"':
            quoted = not quoted
        elif character == "," and not quoted:
            breaks.append(index + 1)
    lines = []
    start = 0
    indent = ""
    while len(indent) + len(statement) - start > limit:
        room = limit - len(indent) - 2
        fitting = [point for point in breaks if start < point <= start + room]
        if not fitting:
            raise SuiteError(f"a statement cannot be wrapped: {statement[start:start + 60]}...")
        lines.append(indent + statement[start:fitting[-1]] + " _")
        start = fitting[-1] + 1
        indent = "    "
    lines.append(indent + statement[start:])
    return lines


def text_literal(value: str) -> str:
    if not value.isascii() or any(ord(character) < 32 for character in value):
        raise SuiteError(f"a text is not printable ASCII: {value!r}")
    return '"' + value.replace('"', '""') + '"'


def number_literal(value: float) -> str:
    """A number as the VBE keeps it: at most 15 significant digits, no exponent."""
    number = float(value)
    if number != number or number in (float("inf"), float("-inf")):
        raise SuiteError(f"not a finite number: {value!r}")
    if number == int(number) and abs(number) < 1e15:
        return str(int(number)) + ("" if abs(number) < 2 ** 31 else "#")
    return format(Decimal(repr(float(f"{number:.15g}"))), "f")


def date_literal(value: str) -> str | None:
    try:
        date = datetime.date.fromisoformat(value)
    except ValueError:
        return None
    return f"DateSerial({date.year}, {date.month}, {date.day})"


def cell_literal(key: str, value: object) -> str:
    """A fixture value as the statement that types it: as run_cases.cell_value."""
    if value is None:
        return "Empty"
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, (int, float)):
        return number_literal(value)
    if isinstance(value, str):
        date = date_literal(value) if key.endswith("_date") else None
        return date if date is not None else text_literal(value)
    raise SuiteError(f"{key}: a value of type {type(value).__name__} cannot be typed in a cell")


def pascal(case_id: str) -> str:
    return "".join(part.capitalize() for part in re.split(r"[^A-Za-z0-9]+", case_id) if part)


def slug(text: str) -> str:
    return "-".join(part for part in re.split(r"[^a-z0-9]+", text.lower()) if part)


def input_lines(fixture: dict[str, Any]) -> list[str]:
    """The statements that type a fixture's records, as run_cases.write_inputs does."""
    datasets = fixture["datasets"]
    date = date_literal(fixture["reporting_date"])
    if date is None:
        raise SuiteError(f"{fixture['id']}: reporting_date is not a date")
    lines = ["ClearInputs", f'SetNamed "AsOfDate", {date}',
             f'SetNamed "ReportingCcy", {text_literal(fixture["reporting_currency_code"])}']
    for index, record in enumerate(datasets.get("DS-05", []), 1):
        lines += wrapped(f"PutRate {index}, {text_literal(record['id'])}, "
                         f"{cell_literal('fx_rate', record['inputs'].get('fx_rate'))}")
    for record in datasets.get("DS-10", []):
        if record["id"] not in run_cases.RESULTS_ROW:
            raise SuiteError(f"{fixture['id']}: no option charge cell for {record['id']}")
        lines += wrapped(f"SetNamed NM_OPTIONS_PREFIX & {text_literal(record['id'])}, "
                         f"{cell_literal('charge_amount', record['inputs'].get('charge_amount'))}")
    for dataset, (_, keys) in run_cases.POSITIONS.items():
        for index, record in enumerate(datasets.get(dataset, []), 1):
            values = [text_literal(record["id"]),
                      *("Empty" if key is None else cell_literal(key, record["inputs"].get(key))
                        for key in keys)]
            while values[-1] == "Empty":
                values.pop()
            lines += wrapped(f"PutRow {INPUT_HEADERS[dataset]}, {index}, "
                             f"Array({', '.join(values)})")
    return lines


def reader(fixture: dict[str, Any], output: dict[str, Any]) -> tuple[str, bool]:
    """The expression that reads an expected output, as run_cases.actual does, and
    whether it is a number already (True) or the value of a cell (False)."""
    quantity, scope = output["quantity"], output["scope"]
    results_first = run_cases.RESULTS_ROW["IR"] - 1
    if scope is None:
        if quantity in run_cases.NAMED:
            return f"NamedCell({NAMES[run_cases.NAMED[quantity]]})", False
        name, column = run_cases.TOTALS[quantity]
        return (f"ResultsCell({run_cases.RESULTS_ROW[name] - results_first}, "
                f"{column - FIRST_COLUMN})"), False
    dataset, record = scope["dataset"], scope["record"]
    if dataset == "DS-10":
        return (f"ResultsCell({run_cases.RESULTS_ROW[record] - results_first}, "
                f"{run_cases.BY_CLASS[quantity] - FIRST_COLUMN})"), False
    if dataset == "DS-05":
        if quantity == "fx_net_position":
            return f"FxNet({text_literal(record)})", False
        offsets = ", ".join(str(column - FIRST_COLUMN)
                            for column in run_cases.ZONES_OUTPUTS[quantity])
        return f"ZonesTotal({text_literal(record)}, Array({offsets}))", True
    index = [item["id"] for item in fixture["datasets"][dataset]].index(record) + 1
    sheet, column = run_cases.ROW_OUTPUTS[quantity]
    return f"RowCell({CALC_HEADERS[sheet]}, {index}, {column - FIRST_COLUMN})", False


def assertion_names(expected: dict[str, Any]) -> list[str]:
    """One name per expected output, unique in the case: the quantity, the record it
    is scoped to, and the output's position where two would otherwise be equal."""
    plain = [slug(output["quantity"]) + ("" if output["scope"] is None
                                          else "-of-" + slug(output["scope"]["record"]))
             for output in expected["outputs"]]
    names = [name if plain.count(name) == 1 else f"{name}-output-{index}"
             for index, name in enumerate(plain, 1)]
    if len(set(names)) != len(names) or not all(names):
        raise SuiteError(f"{expected['fixture']}: the outputs do not give unique names")
    return names


def case_lines(fixture: dict[str, Any], expected: dict[str, Any]) -> tuple[list[str], int, bool]:
    """The statements of a case between keeping the inputs and putting them back,
    the assertions among them, and whether one compares a number."""
    case_id = fixture["id"]
    lines = [f"Inputs{pascal(case_id)}", "ResultText = RunSSA_Silent()"]
    numeric = False
    for name, output in zip(assertion_names(expected), expected["outputs"], strict=True):
        expression, is_number = reader(fixture, output)
        if "text" in output:
            lines += wrapped(f'TEST_Harness.AssertEqualText "{name}", '
                             f"{text_literal(output['text'])}, PlainText({expression})")
            continue
        numeric = True
        wanted = float(output["value"])
        if not is_number:
            # A cell that is empty holds a zero, and only a zero: as run_cases.agrees.
            expression = f"{'ZeroIfEmpty' if wanted == 0 else 'NumberOf'}({expression})"
        tolerance = output["tolerance"]
        scope = output["scope"]
        where = "vbNullString, vbNullString" if scope is None else \
            f"{text_literal(scope['dataset'])}, {text_literal(scope['record'])}"
        lines += wrapped(f"Actual = {expression}")
        lines += wrapped(f'TEST_Harness.AssertNear "{name}", {number_literal(wanted)}, Actual, '
                         f"{number_literal(tolerance['absolute'])}, "
                         f"{number_literal(tolerance['relative'])}")
        lines += wrapped(f"TEST_Harness.RecordResult {text_literal(case_id)}, vbNullString, "
                         f"{text_literal(output['quantity'])}, {where}, Actual")
    return lines, len(expected["outputs"]), numeric


def comment(text: str, width: int = 72) -> str:
    """A text as comment lines of a header: ASCII, wrapped at word ends."""
    if not text.isascii():
        raise SuiteError(f"a description is not ASCII: {text!r}")
    lines, line = [], ""
    for word in text.split():
        if line and len(line) + 1 + len(word) > width:
            lines.append(line)
            line = word
        else:
            line = f"{line} {word}".strip()
    return "\n".join([*lines, line])


def loaded() -> list[tuple[dict[str, Any], dict[str, Any]]]:
    pairs = []
    for fixture in run_cases.cases([]):
        path = run_cases.EXPECTED / f"{fixture['id']}.json"
        pairs.append((fixture, json.loads(path.read_text(encoding="utf-8"))))
    if not pairs:
        raise SuiteError("no test case of the project's calculations")
    return pairs


def updated(expected: dict[str, Any]) -> str:
    return max(str(output["reference"]["date"]) for output in expected["outputs"])


FIXED_HEADER = """' PURPOSE
'   The cases suite of the workbook's tests: every registered test case of
'   the project's calculations, as a case of TEST_Harness that types the
'   test case's inputs on the input sheets, runs the calculation through
'   RunSSA_Silent and compares every expected output with the cell that
'   holds it. GENERATED by tools/prototype/case_suite.py from tests/fixtures/
'   and tests/expected/: do not edit it; change the test case and write the
'   module again.
'
' PUBLIC SURFACE
'   RunCaseTests is the documented test entry point. Test_RunCases runs the
'   same suite and returns its result line where a caller cannot read an
'   error, as tools/devbuild/lo_macro.py. CertifyCases certifies the run and
'   saves its log, its results and its certification log as evidence of an
'   Excel run. Option Private Module keeps them out of the workbook's
'   external surface. Never part of the production import set.
'
' DEPENDENCIES
'   TEST_Harness; M_Main, CORE_Config and CORE_Util; the workbook's names
'   and sheets.
'
' STATE OWNERSHIP
'   Owns the copy of the inputs a case keeps while it runs. Each case
'   replaces the inputs of the workbook with those of its test case: the
'   reporting date and currency, the FX table, the option charges and the
'   four position sheets. It puts back the inputs the workbook had, on the
'   success and the error path, and asserts it. It rewrites the output
'   sheets through RunSSA_Silent and clears them when it puts the inputs
'   back.
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START. An unexpected error inside a case is
'   recorded as a failure of that case, and the run continues. Test_RunCases
'   contains those errors; CertifyCases records them, saves the log, and
'   raises only when the certification did not pass.
'
' USAGE
'   In a development workbook, type TEST_Cases.RunCaseTests in the Immediate
'   window. To record an Excel run, type
'   TEST_Cases.CertifyCases "<folder>", "<checkout>\\tests\\fixtures"
'   (INSTALLATION.md, Recording an Excel run).
'
' COMPATIBILITY
'   Excel VBA; also runs under LibreOffice's VBA mode through the tools of
'   tools/devbuild/, which is never Excel evidence.
'
' REFERENCE
'   Each expected value, its class and its tolerance are those of the test
'   case's file in tests/expected/ (docs/methodology/TEST_CASES.md)."""


def entry_points(cases: int, assertions: int) -> list[str]:
    lines = banner("ENTRY POINTS")
    lines += procedure(["Public Sub RunCaseTests()"], [
        ("PURPOSE", "Run the cases suite. A passing run ends with\n"
                    f"RESULT=PASS; suite={SUITE}; operation=PASS; completeness=COMPLETE;\n"
                    f"cases={cases}; assertions={assertions}; failures=0; cleanup=PASS; "
                    "recovery=NONE."),
        ("STATE OWNERSHIP", "Replaces the inputs for each case and puts them back; leaves the\n"
                            "output sheets cleared."),
        ("ERROR POLICY", "As TEST_Harness.EndRun: raises when the run does not pass."),
    ], [], ["RunSuite"])
    lines += procedure(["Public Function Test_RunCases() As String"], [
        ("PURPOSE", "Run the cases suite for a caller that reads a result and not an error."),
        ("RETURNS", "The RESULT line of the run, as TEST_Harness.LastResult gives it."),
        ("STATE OWNERSHIP", "As RunCaseTests."),
        ("ERROR POLICY", "Contains the error of a run that does not pass, or of a refused "
                         "start:\nthe result line says so."),
    ], [], ["On Error Resume Next", "RunSuite", "On Error GoTo 0",
            "Test_RunCases = TEST_Harness.LastResult()"])
    lines += procedure(["Public Sub CertifyCases( _", "    ByVal Folder As String, _",
                        "    ByVal Fixtures As String)"], [
        ("PURPOSE", "Certify the cases suite's run as evidence of an Excel run: run it, save\n"
                    f"its log as {SUITE}.log and its results as {SUITE}-results.json, and "
                    f"write\n{SUITE}-certification.log."),
        ("INPUTS", "Folder: an existing folder for the files, without a trailing separator.\n"
                   "Fixtures: the full path of tests/fixtures/ in the checkout; not read,\n"
                   "because the inputs are in this module, and taken so that every\n"
                   "certification entry point has the same two arguments."),
        ("STATE OWNERSHIP", "As RunCaseTests."),
        ("ERROR POLICY", "The error of the run is recorded, not raised, so that the log is "
                         "always\nsaved. Raises ERR_TEST_FAILED after writing the "
                         "certification log when\nthe certification did not pass. An error of "
                         "BeginCertification,\nSaveLog or SaveResults propagates; run "
                         "TEST_Harness.ResetTests before\ntrying again."),
    ], [("RunErrorNumber", "Long", "Error that ended the run; 0 for none"),
        ("RunErrorDescription", "String", "Its description")], [
        "TEST_Harness.BeginCertification SUITE_NAME",
        "On Error Resume Next",
        "RunSuite",
        "RunErrorNumber = Err.Number",
        "RunErrorDescription = Err.Description",
        "On Error GoTo 0",
        f'TEST_Harness.SaveLog Folder & "\\{SUITE}.log"',
        f'TEST_Harness.Echo "SAVED=" & Folder & "\\{SUITE}.log; mode=NORMAL; error=" & _',
        '    CStr(RunErrorNumber) & IIf(RunErrorNumber = 0, vbNullString, " " & '
        "RunErrorDescription)",
        'TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription',
        f'TEST_Harness.SaveResults Folder & "\\{SUITE}-results.json"',
        f'TEST_Harness.EndCertification Folder & "\\{SUITE}-certification.log"'])
    return lines


def runner(names: list[str]) -> list[str]:
    return banner("RUNNER") + procedure(["Private Sub RunSuite()"], [
        ("PURPOSE", "Run every case in a fixed order between BeginRun and EndRun."),
        ("ERROR POLICY", "An error escaping a case handler is passed to EndRun, which reports "
                         "it\nand re-raises it unchanged. BeginRun's refusal propagates "
                         "directly."),
    ], [("ErrNumber", "Long", "Runner error number; 0 for none"),
        ("ErrSource", "String", "Runner error source"),
        ("ErrDescription", "String", "Runner error description")], [
        "TEST_Harness.BeginRun SUITE_NAME, \"NORMAL\", EXPECTED_CASES, EXPECTED_ASSERTIONS",
        "On Error GoTo RunFailed",
        *names,
        "",
        "Finish:",
        "On Error GoTo 0",
        "TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription",
        "Exit Sub",
        "",
        "RunFailed:",
        "ErrNumber = Err.Number",
        "ErrSource = Err.Source",
        "ErrDescription = Err.Description",
        "Resume Finish"])


HELPERS: list[tuple[list[str], list[tuple[str, str]], list[tuple[str, str, str]], list[str]]] = [
    (["Private Function InputArea( _", "    ByVal Index As Long) _", "    As Range"], [
        ("PURPOSE", "One of the areas that hold the inputs a test case replaces."),
        ("INPUTS", "Index   1 to 4: the rows of the four position sheets, INPUT_ROWS of\n"
                   "        them; 5: the FX table, with its descriptions; 6 and 7: the\n"
                   "        reporting date and currency; 8 to 11: the option charges."),
    ], [("Header", "Range", "Header of a position sheet"), ("Codes", "Range", "FX table codes"),
        ("Rates", "Range", "FX table rates")], [
        "Select Case Index",
        "    Case 1: Set Header = ThisWorkbook.Names(NM_IR_POS_HEADER).RefersToRange",
        "            Set InputArea = Header.Offset(1, 0).Resize(INPUT_ROWS, IR_INPUT_WIDTH)",
        "    Case 2: Set Header = ThisWorkbook.Names(NM_EQ_POS_HEADER).RefersToRange",
        "            Set InputArea = Header.Offset(1, 0).Resize(INPUT_ROWS, EQ_INPUT_WIDTH)",
        "    Case 3: Set Header = ThisWorkbook.Names(NM_CO_POS_HEADER).RefersToRange",
        "            Set InputArea = Header.Offset(1, 0).Resize(INPUT_ROWS, CO_INPUT_WIDTH)",
        "    Case 4: Set Header = ThisWorkbook.Names(NM_FX_POS_HEADER).RefersToRange",
        "            Set InputArea = Header.Offset(1, 0).Resize(INPUT_ROWS, FX_INPUT_WIDTH)",
        "    Case 5: Set Codes = ThisWorkbook.Names(NM_FX_CCY).RefersToRange",
        "            Set Rates = ThisWorkbook.Names(NM_FX_RATE).RefersToRange",
        "            Set InputArea = Codes.Resize(Codes.Rows.Count, _",
        "                                         Rates.Column - Codes.Column + 1)",
        '    Case 6: Set InputArea = ThisWorkbook.Names("AsOfDate").RefersToRange',
        '    Case 7: Set InputArea = ThisWorkbook.Names("ReportingCcy").RefersToRange',
        '    Case 8: Set InputArea = ThisWorkbook.Names(NM_OPTIONS_PREFIX & "IR").RefersToRange',
        '    Case 9: Set InputArea = ThisWorkbook.Names(NM_OPTIONS_PREFIX & "EQ").RefersToRange',
        '    Case 10: Set InputArea = ThisWorkbook.Names(NM_OPTIONS_PREFIX & "CO").RefersToRange',
        '    Case 11: Set InputArea = ThisWorkbook.Names(NM_OPTIONS_PREFIX & "FX").RefersToRange',
        "End Select"]),
    (["Private Sub KeepInputs()"], [
        ("PURPOSE", "Keep what every input area holds, before a case replaces it."),
        ("STATE OWNERSHIP", "Sets mKept."),
    ], [("i", "Long", "Input area")], [
        "For i = 1 To INPUT_AREAS",
        "    mKept(i) = InputArea(i).Value",
        "Next i"]),
    (["Private Function RestoreInputs() As String"], [
        ("PURPOSE", "Put the sheets back as they were before a case: write what "
                    "KeepInputs\nkept, check that every input cell holds it, and clear the "
                    "output\nsheets, which show the results of the test case."),
        ("RETURNS", "An empty text when the sheets are restored; otherwise what failed."),
        ("STATE OWNERSHIP", "Writes every input area and clears the output sheets."),
        ("ERROR POLICY", "Contains its own error and returns it, so that the case asserts "
                         "the\nresult on the success and the error path."),
    ], [("i", "Long", "Input area"), ("Differences", "Long", "Input cells that differ"),
        ("Cleared", "String", "ClearOutputs_Silent result line")], [
        "On Error GoTo Failed",
        "'Emptied first: an empty value written over a cell that holds something",
        "'does not leave it empty in every host.",
        "For i = 1 To INPUT_AREAS",
        "    InputArea(i).ClearContents",
        "    InputArea(i).Value = mKept(i)",
        "Next i",
        "Differences = InputDifferences()",
        "Cleared = ClearOutputs_Silent()",
        "If Differences <> 0 Then",
        '    RestoreInputs = CStr(Differences) & " input cell(s) do not hold what they held"',
        'ElseIf Left$(Cleared, 14) <> "RESULT=CLEARED" Then',
        '    RestoreInputs = "the output sheets are not cleared: " & Cleared',
        "End If",
        "Exit Function",
        "",
        "Failed:",
        'RestoreInputs = "error " & CStr(Err.Number) & " while restoring: " & Err.Description']),
    (["Private Function InputDifferences() As Long"], [
        ("PURPOSE", "Count the input cells that do not hold what KeepInputs kept."),
        ("RETURNS", "The number of cells; 0 when the inputs are those kept."),
    ], [("Current", "Variant", "What an area holds now"), ("i", "Long", "Input area"),
        ("r", "Long", "Row"), ("c", "Long", "Column"), ("Count", "Long", "Cells that differ")], [
        "For i = 1 To INPUT_AREAS",
        "    Current = InputArea(i).Value",
        "    If IsArray(Current) Then",
        "        For r = LBound(Current, 1) To UBound(Current, 1)",
        "            For c = LBound(Current, 2) To UBound(Current, 2)",
        "                If Not SameValue(Current(r, c), mKept(i)(r, c)) Then Count = Count + 1",
        "            Next c",
        "        Next r",
        "    ElseIf Not SameValue(Current, mKept(i)) Then",
        "        Count = Count + 1",
        "    End If",
        "Next i",
        "InputDifferences = Count"]),
    (["Private Function SameValue( _", "    ByVal First As Variant, _",
      "    ByVal Second As Variant) _", "    As Boolean"], [
        ("PURPOSE", "Whether two cell values are the same value of the same type."),
        ("ERROR POLICY", "An error value equals only the same error value."),
    ], [], [
        "If VarType(First) <> VarType(Second) Then",
        "    SameValue = False",
        "ElseIf IsEmpty(First) Then",
        "    SameValue = True",
        "ElseIf IsError(First) Then",
        "    SameValue = (CStr(First) = CStr(Second))",
        "Else",
        "    SameValue = (First = Second)",
        "End If"]),
    (["Private Sub ClearInputs()"], [
        ("PURPOSE", "Empty every input area a test case fills: the position sheets, the FX\n"
                    "table and the option charges. The reporting date and currency are\n"
                    "always set by the case."),
        ("STATE OWNERSHIP", "Clears the contents of those areas; their formats stay."),
    ], [("i", "Long", "Input area")], [
        "For i = 1 To INPUT_AREAS",
        "    If i <> 6 And i <> 7 Then InputArea(i).ClearContents",
        "Next i"]),
    (["Private Sub SetNamed( _", "    ByVal CellName As String, _",
      "    ByVal Value As Variant)"], [
        ("PURPOSE", "Type a value in the cell a workbook name refers to."),
    ], [], ["ThisWorkbook.Names(CellName).RefersToRange.Value = Typed(Value)"]),
    (["Private Sub PutRate( _", "    ByVal Index As Long, _", "    ByVal Code As String, _",
      "    ByVal Rate As Variant)"], [
        ("PURPOSE", "Type one row of the FX table: a currency and its rate."),
        ("INPUTS", "Index   Row of the table, from 1.\nCode    Currency code.\n"
                   "Rate    Its rate; Empty leaves the cell empty."),
    ], [], [
        "ThisWorkbook.Names(NM_FX_CCY).RefersToRange.Cells(Index, 1).Value = Typed(Code)",
        "If Not IsEmpty(Rate) Then",
        "    ThisWorkbook.Names(NM_FX_RATE).RefersToRange.Cells(Index, 1).Value = Typed(Rate)",
        "End If"]),
    (["Private Sub PutRow( _", "    ByVal HeaderName As String, _", "    ByVal Index As Long, _",
      "    ByVal Values As Variant)"], [
        ("PURPOSE", "Type one row of a position sheet."),
        ("INPUTS", "HeaderName   Workbook name of the sheet's header cell.\n"
                   "Index        Row under the header, from 1.\n"
                   "Values       The cells of the row from the PositionID; Empty leaves a\n"
                   "             cell empty."),
    ], [("Target", "Range", "First cell of the row"), ("c", "Long", "Cell of the row")], [
        "Set Target = ThisWorkbook.Names(HeaderName).RefersToRange.Offset(Index, 0)",
        "For c = LBound(Values) To UBound(Values)",
        "    If Not IsEmpty(Values(c)) Then",
        "        Target.Offset(0, c - LBound(Values)).Value = Typed(Values(c))",
        "    End If",
        "Next c"]),
    (["Private Function Typed( _", "    ByVal Value As Variant) _", "    As Variant"], [
        ("PURPOSE", "A value as it is assigned to a cell so that the cell holds it as the\n"
                    "test case gives it: a text stays a text."),
        ("RETURNS", "A text with an apostrophe before it, which makes the cell hold the\n"
                    "text as typed: without it, a text that reads as a number or a date\n"
                    "becomes one. Any other value unchanged."),
    ], [], [
        "If VarType(Value) = vbString Then",
        "    Typed = \"'\" & Value",
        "Else",
        "    Typed = Value",
        "End If"]),
    (["Private Function RowCell( _", "    ByVal HeaderName As String, _",
      "    ByVal Index As Long, _", "    ByVal ColumnOffset As Long) _", "    As Variant"], [
        ("PURPOSE", "The value of one cell of a calculation sheet."),
        ("INPUTS", "HeaderName     Workbook name of the sheet's header cell.\n"
                   "Index          Row under the header, from 1: the position's row.\n"
                   "ColumnOffset   Columns to the right of the PositionID."),
    ], [], ["RowCell = ThisWorkbook.Names(HeaderName).RefersToRange.Offset(Index, "
            "ColumnOffset).Value"]),
    (["Private Function ResultsCell( _", "    ByVal RowOffset As Long, _",
      "    ByVal ColumnOffset As Long) _", "    As Variant"], [
        ("PURPOSE", "The value of one cell of the table of Results."),
        ("INPUTS", "RowOffset      Rows under the header: 1 interest rate, 2 equity,\n"
                   "               3 commodity, 4 foreign exchange.\n"
                   "ColumnOffset   Columns to the right of the risk class."),
    ], [], ["ResultsCell = ThisWorkbook.Names(NM_RESULTS_HEADER).RefersToRange _",
            "    .Offset(RowOffset, ColumnOffset).Value"]),
    (["Private Function NamedCell( _", "    ByVal CellName As String) _", "    As Variant"], [
        ("PURPOSE", "The value of the cell a workbook name refers to."),
    ], [], ["NamedCell = ThisWorkbook.Names(CellName).RefersToRange.Value"]),
    (["Private Function CurrencyRow( _", "    ByVal HeaderName As String, _",
      "    ByVal ColumnOffset As Long, _", "    ByVal Code As String) _", "    As Range"], [
        ("PURPOSE", "The row of a table by currency: the first row under the header "
                    "whose\ncode is the one given, without regard to case."),
        ("INPUTS", "HeaderName     Workbook name of the table's header cell.\n"
                   "ColumnOffset   Column of the code, to the right of the header.\n"
                   "Code           Currency code."),
        ("RETURNS", "The cell of that row in the header's column."),
        ("ERROR POLICY", "Raises error 5 when no row within SEARCH_ROWS holds the code; the\n"
                         "case records it."),
    ], [("Header", "Range", "Header cell"), ("r", "Long", "Row under the header")], [
        "Set Header = ThisWorkbook.Names(HeaderName).RefersToRange",
        "For r = 1 To SEARCH_ROWS",
        "    If CodeText(Header.Offset(r, ColumnOffset).Value) = CodeText(Code) Then",
        "        Set CurrencyRow = Header.Offset(r, 0)",
        "        Exit Function",
        "    End If",
        "Next r",
        'Err.Raise 5, "TEST_Cases.CurrencyRow", "No row of " & HeaderName & " holds " & Code & "."',
    ]),
    (["Private Function ZonesTotal( _", "    ByVal Code As String, _",
      "    ByVal ColumnOffsets As Variant) _", "    As Double"], [
        ("PURPOSE", "The sum of some cells of a currency's row of IR_Zones."),
        ("INPUTS", "Code            Currency code.\n"
                   "ColumnOffsets   The columns to add, to the right of the currency."),
        ("RETURNS", "The sum; an empty cell adds nothing."),
        ("ERROR POLICY", "As CurrencyRow, and NumberOf for a cell that holds no number."),
    ], [("First", "Range", "The currency's cell"), ("Item", "Variant", "One column offset"),
        ("Total", "Double", "Sum")], [
        "Set First = CurrencyRow(NM_ZONES_HEADER, 0, Code)",
        "For Each Item In ColumnOffsets",
        "    Total = Total + ZeroIfEmpty(First.Offset(0, CLng(Item)).Value)",
        "Next Item",
        "ZonesTotal = Total"]),
    (["Private Function FxNet( _", "    ByVal Code As String) _", "    As Variant"], [
        ("PURPOSE", "The net position of a currency in the FX block of Buckets, in the\n"
                    "reporting currency."),
        ("ERROR POLICY", "As CurrencyRow."),
    ], [], [
        "FxNet = CurrencyRow(NM_BK_FX_HEADER, FX_BLOCK_CODE, Code).Offset(0, FX_BLOCK_NET).Value"]),
    (["Private Function ZeroIfEmpty( _", "    ByVal Value As Variant) _", "    As Double"], [
        ("PURPOSE", "The number a cell holds, where the expected value is zero: an empty\n"
                    "cell holds a zero."),
        ("ERROR POLICY", "As NumberOf for a cell that is not empty."),
    ], [], [
        "If IsEmpty(Value) Then",
        "    ZeroIfEmpty = 0#",
        "ElseIf VarType(Value) = vbString And Len(Value) = 0 Then",
        "    ZeroIfEmpty = 0#",
        "Else",
        "    ZeroIfEmpty = NumberOf(Value)",
        "End If"]),
    (["Private Function NumberOf( _", "    ByVal Value As Variant) _", "    As Double"], [
        ("PURPOSE", "The number a cell holds, for an assertion that compares numbers."),
        ("ERROR POLICY", "Raises error 13 when Value is not a number, so that an empty cell, a\n"
                         "text or an error value is never compared as 0; the case records it."),
    ], [], [
        "If Not IsNumberValue(Value) Then",
        '    Err.Raise 13, "TEST_Cases.NumberOf", "The result is not a number."',
        "End If",
        "NumberOf = CDbl(Value)"]),
]


# How a case ends: an error of the case is recorded first, then the sheets are
# restored, which is asserted once, on the success and the error path. The handler
# is off during the cleanup, so that an error there cannot come back to it.
CLEANUP = ["CleanUp:",
           "On Error GoTo 0",
           "If RestoreDue Then",
           "    Detail = RestoreInputs()",
           "End If",
           'TEST_Harness.AssertTrue "inputs-restored", Len(Detail) = 0, Detail',
           "Exit Sub",
           "",
           "CaseFailed:",
           "TEST_Harness.RecordCaseError",
           "Resume CleanUp"]


def module() -> str:
    """The text of the module, with the CRLF line ends of a VBA source file."""
    pairs = loaded()
    names = [f"Case{pascal(fixture['id'])}" for fixture, _ in pairs]
    if len(set(names)) != len(names):
        raise SuiteError("two test cases give the same procedure name")
    bodies = []
    assertions = 0
    for fixture, expected in pairs:
        if expected.get("fixture") != fixture["id"] or "outputs" not in expected:
            raise SuiteError(f"{fixture['id']}: its expected file is not that of a JSON fixture")
        lines, count, numeric = case_lines(fixture, expected)
        assertions += count + 1  # And inputs-restored.
        bodies.append((fixture, expected, lines, numeric))
    newest = max(WRITTEN, *(updated(expected) for _, expected in pairs))
    text = ['Attribute VB_Name = "TEST_Cases"', RULE, "' MODULE: TEST_Cases", THIN,
            *FIXED_HEADER.replace("\\\\", "\\").split("\n"), "'", "' UPDATED", f"'   {newest}",
            "'", "' AUTHOR", "'   Daniele Penza", RULE, "", THIN, "' MODULE SETTINGS", THIN,
            "    'Require explicit declarations; keep the tests inside this project.",
            "    Option Explicit", "    Option Private Module", "", THIN, "' MODULE CONSTANTS",
            THIN,
            "    'A complete run starts every case and evaluates every assertion once.",
            f'        Private Const SUITE_NAME          As String = "{SUITE}"',
            f"        Private Const EXPECTED_CASES      As Long = {len(pairs)}",
            f"        Private Const EXPECTED_ASSERTIONS As Long = {assertions}",
            "    'The inputs a test case replaces (InputArea).",
            "        Private Const INPUT_AREAS         As Long = 11    'Areas kept and put back",
            f"        Private Const INPUT_ROWS          As Long = {run_cases.CAPACITY}   "
            "'Rows of a position sheet kept and cleared",
            "    'Where a result per currency is found.",
            "        Private Const SEARCH_ROWS         As Long = 60    "
            "'Rows searched under a header",
            f"        Private Const FX_BLOCK_CODE       As Long = "
            f"{run_cases.FX_BLOCK_CODE - FIRST_COLUMN}     'Buckets FX block: currency",
            f"        Private Const FX_BLOCK_NET        As Long = "
            f"{run_cases.FX_BLOCK_NET - FIRST_COLUMN}     'Buckets FX block: net position",
            "", THIN, "' MODULE STATE", THIN,
            "        Private mKept(1 To 11)   As Variant   'What each input area held",
            "", ""]
    text += entry_points(len(pairs), assertions)
    text += runner(names)
    text += banner("CASES")
    for fixture, expected, lines, numeric in bodies:
        name = pascal(fixture["id"])
        date = updated(expected)
        text += procedure([f"Private Sub Inputs{name}()"], [
            ("PURPOSE", f"Type the inputs of the test case {fixture['id']}."[:75]),
            ("STATE OWNERSHIP", "Replaces the inputs of the workbook."),
        ], [], input_lines(fixture), date)
        dims = [("ResultText", "String", "RunSSA_Silent result line, not compared")]
        if numeric:
            dims.append(("Actual", "Double", "Number read from the workbook"))
        dims += [("RestoreDue", "Boolean", "True once the inputs are kept"),
                 ("Detail", "String", "Why the sheets could not be restored")]
        text += procedure([f"Private Sub Case{name}()"], [
            ("PURPOSE", comment(f"Test case {fixture['id']}, on {fixture['calculation']}: "
                                f"{fixture['description']}.")),
            ("STATE OWNERSHIP", "Replaces the inputs of the workbook and rewrites the output\n"
                                "sheets; puts the inputs back and clears the output sheets."),
        ], dims, ["On Error GoTo CaseFailed", f'TEST_Harness.BeginCase "{fixture["id"]}"',
                  "KeepInputs", "RestoreDue = True", *lines, "", *CLEANUP], date)
    text += banner("HELPERS")
    for signature, sections, dims, body in HELPERS:
        text += procedure(signature, sections, dims, body)
    while text[-1] == "":
        text.pop()
    return format_vba.format_text("\n".join(text) + "\n")


def encoded(text: str) -> bytes:
    return text.encode("ascii")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--write", action="store_true", help="write the module")
    args = parser.parse_args(argv)
    try:
        data = encoded(module())
    except (SuiteError, OSError, ValueError, KeyError) as error:
        print(f"case-suite: ERROR ({error})", file=sys.stderr)
        return 2
    shown = MODULE.relative_to(ROOT).as_posix()
    if args.write:
        MODULE.write_bytes(data)
        print(f"case-suite: wrote {shown}")
        return 0
    if not MODULE.is_file() or MODULE.read_bytes() != data:
        print(f"case-suite: FAIL ({shown} is not what the test cases give; run "
              "python tools/prototype/case_suite.py --write)")
        return 1
    print(f"case-suite: PASS ({shown} is what the test cases give)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
