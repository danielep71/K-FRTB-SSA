#!/usr/bin/env python3
"""Write the VBA suite that runs the registered test cases in the workbook.

Usage: python tools/prototype/case_suite.py [--write]

It reads each JSON fixture of tests/fixtures/ whose calculation is one of the
project's, CALC-02 to CALC-09, with its expected file of tests/expected/, and
writes tests/modules/TEST_Cases.bas: a suite of the VBA test harness with one
case per test case. A case types the fixture's parameters and records on the
input sheets,
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
import textwrap
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
WRITTEN = "2026-11-19"
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


def flow(text: str, width: int = 72) -> list[str]:
    """The lines of a field's text: a line that is too long is wrapped at word ends,
    under the indentation it starts with; an empty line stays empty."""
    lines = []
    for raw in text.strip("\n").split("\n"):
        if len(raw) <= width:
            lines.append(raw.rstrip())
            continue
        indent = raw[:len(raw) - len(raw.lstrip())]
        hanging = indent + ("  " if raw.lstrip().startswith("- ") else "")
        lines += textwrap.wrap(raw.strip(), width, initial_indent=indent,
                               subsequent_indent=hanging, break_long_words=False,
                               break_on_hyphens=False)
    return lines


def arguments(*items: tuple[str, str]) -> str:
    """The text of an INPUTS field: each argument's name on a line of its own, its
    description below it, and an empty line between two arguments."""
    return "\n\n".join(name + "\n" + "\n".join("  " + line for line in text.split("\n"))
                       for name, text in items)


def fields(purpose: str, why: str, behavior: str, *, inputs: str = "None",
           returns: str = "Nothing", error: str = "None", deps: str = "None",
           notes: str = "None") -> list[tuple[str, str]]:
    """The eight fields of a header that come before UPDATED, in their order."""
    return [("PURPOSE", purpose), ("WHY THIS EXISTS", why), ("INPUTS", inputs),
            ("RETURNS", returns), ("BEHAVIOR", behavior), ("ERROR POLICY", error),
            ("DEPENDENCIES", deps), ("NOTES", notes)]


def field_lines(sections: list[tuple[str, str]], updated: str) -> list[str]:
    """The nine fields, a lone apostrophe between two of them."""
    lines: list[str] = []
    for index, (title, text) in enumerate([*sections, ("UPDATED", updated)]):
        if index:
            lines.append("'")
        lines.append(f"' {title}")
        lines += [f"'   {line}".rstrip() for line in flow(text)]
    return lines


def module_header(sections: list[tuple[str, str]], updated: str) -> list[str]:
    """The component header. A test module keeps the `=` of its outer banners."""
    return ["'", RULE, "' MODULE: TEST_Cases", THIN, *field_lines(sections, updated), RULE]


def header(name: str, sections: list[tuple[str, str]], updated: str) -> list[str]:
    """The procedure header: the title banner, the nine fields, the closing banner."""
    return ["'", THIN, centred(name), THIN, *field_lines(sections, updated), THIN]


def group(title: str) -> list[str]:
    """A group banner, with the blank line that separates it from the signature."""
    return ["'", THIN, "'", centred(title), "'", THIN, "'", ""]


def section(title: str) -> list[str]:
    return [THIN, f"' {title}", THIN]


def aligned(rows: list[tuple[str, str, str]], *, tight: bool = False) -> list[str]:
    """Declaration rows (head, rest, comment) at four spaces, the `As` in one column
    and the comment in another, both on a multiple of four. With ``tight`` the `As`
    follows the longest head after one blank."""
    kind = 4 + max(len(head) for head, _, _ in rows) + 1
    if not tight:
        kind += -kind % 4
    note = kind + max(len(rest) for _, rest, _ in rows) + 1
    note += -note % 4
    return ["    " + head.ljust(kind - 4) + rest.ljust(note - kind) + "'" + comment
            for head, rest, comment in rows]


def body_line(line: str) -> str:
    """A line of a body section. A label stays at the margin and an intent comment
    stands at four spaces, or at the indentation of the code it describes when that
    is nested; any other line is code at eight spaces, with its own indentation."""
    text = line.lstrip()
    if re.fullmatch(r"\w+:", line):
        return line
    if text.startswith("'"):
        return " " * (4 if line == text else 8 + len(line) - len(text)) + text
    return "        " + line


def procedure(signature: list[str], title: str, head: list[tuple[str, str]],
              dims: list[tuple[str, str, str]], sections: list[tuple[str, list[str]]],
              updated: str = WRITTEN) -> list[str]:
    """A procedure in the house style: the signature, the header, a DECLARE section for
    ``dims``, one section for each of ``sections`` and the End statement. A line of a
    section that is a label stays at the margin; a line that starts with an apostrophe
    is an intent comment; any other is code."""
    match = re.search(r"(Sub|Function) (\w+)", signature[0])
    assert match is not None
    assert match.group(1) == "Sub" or dict(head)["RETURNS"] != "Nothing", signature[0]
    lines = [*signature, *([""] if len(signature) == 1 else []), *header(title, head, updated)]
    if dims:
        lines += ["", *section("DECLARE")]
        lines += aligned([(f"Dim {name}", f"As {kind}", comment) for name, kind, comment in dims])
    for name, body in sections:
        lines += ["", *section(name), *(body_line(line) for line in body)]
    return [*lines, "", f"End {match.group(1)}", ""]


def wrapped(statement: str) -> list[str]:
    """A statement on lines of at most WIDTH characters after the indentation,
    broken after a comma outside a string literal, or else after a blank inside
    one: the literal is then closed, joined with `& _` and opened again."""
    limit = WIDTH - 8
    if len(statement) <= limit:
        return [statement]
    commas = []
    blanks = []
    quoted = False
    for index, character in enumerate(statement):
        if character == '"':
            quoted = not quoted
        elif character == "," and not quoted:
            commas.append(index + 1)
        elif character == " " and quoted:
            blanks.append(index + 1)
    lines = []
    start = 0
    opening = ""
    indent = ""
    while len(indent) + len(opening) + len(statement) - start > limit:
        room = limit - len(indent) - len(opening) - 2
        comma = [point for point in commas if start < point <= start + room]
        blank = [point for point in blanks if start < point <= start + room - 3]
        if comma:
            lines.append(indent + opening + statement[start:comma[-1]] + " _")
            start, opening = comma[-1] + 1, ""
        elif blank:
            lines.append(indent + opening + statement[start:blank[-1]] + '" & _')
            start, opening = blank[-1], '"'
        else:
            raise SuiteError(f"a statement cannot be wrapped: {statement[start:start + 60]}...")
        indent = "    "
    lines.append(indent + opening + statement[start:])
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


# Position dataset: how the module's text names the sheet that holds it.
POSITION_KINDS = {"DS-06": "interest rate", "DS-07": "equity", "DS-08": "commodity",
                  "DS-09": "foreign exchange"}


def input_sections(fixture: dict[str, Any]) -> list[tuple[str, list[str]]]:
    """The sections that type a fixture's records, as run_cases.write_inputs does."""
    datasets = fixture["datasets"]
    date = date_literal(fixture["reporting_date"])
    if date is None:
        raise SuiteError(f"{fixture['id']}: reporting_date is not a date")
    sections = [
        ("CLEAR INPUTS", ["'Empty the input areas that the test case fills", "ClearInputs"]),
        ("SET REPORTING SETTINGS", [
            "'Set the reporting date", f'SetNamed "AsOfDate", {date}',
            "'Set the reporting currency",
            f'SetNamed "ReportingCcy", {text_literal(fixture["reporting_currency_code"])}',
            *parameter_lines(fixture)])]
    rates = ["'Type one row of the FX table for each currency of the test case"]
    for index, record in enumerate(datasets.get("DS-05", []), 1):
        rates += wrapped(f"PutRate {index}, {text_literal(record['id'])}, "
                         f"{cell_literal('fx_rate', record['inputs'].get('fx_rate'))}")
    if len(rates) > 1:
        sections.append(("TYPE FX TABLE", rates))
    charges = ["'Type the option charge of each risk class that the test case gives"]
    for record in datasets.get("DS-10", []):
        if record["id"] not in run_cases.RESULTS_ROW:
            raise SuiteError(f"{fixture['id']}: no option charge cell for {record['id']}")
        charges += wrapped(f"SetNamed NM_OPTIONS_PREFIX & {text_literal(record['id'])}, "
                           f"{cell_literal('charge_amount', record['inputs'].get('charge_amount'))}")
    if len(charges) > 1:
        sections.append(("TYPE OPTION CHARGES", charges))
    for dataset, (_, keys) in run_cases.POSITIONS.items():
        kind = POSITION_KINDS[dataset]
        rows = [f"'Type the {kind} positions of the test case, one row each"]
        for index, record in enumerate(datasets.get(dataset, []), 1):
            values = [text_literal(record["id"]),
                      *("Empty" if key is None else cell_literal(key, record["inputs"].get(key))
                        for key in keys)]
            while values[-1] == "Empty":
                values.pop()
            rows += wrapped(f"PutRow {INPUT_HEADERS[dataset]}, {index}, "
                            f"Array({', '.join(values)})")
        if len(rows) > 1:
            sections.append((f"TYPE {kind.upper()} POSITIONS", rows))
    return sections + month_sections(fixture)


def month_sections(fixture: dict[str, Any]) -> list[tuple[str, list[str]]]:
    """The section that types the records of DS-11, as run_cases.write_months does."""
    records = fixture["datasets"].get("DS-11", [])
    if len(records) > run_cases.MONTH_ROWS:
        raise SuiteError(f"{fixture['id']}: more than {run_cases.MONTH_ROWS} months")
    months = ["'Type the month ends of the test case, one row each from the first row"]
    for index, record in enumerate(records, 1):
        values = [cell_literal(key, record["inputs"].get(key)) for key in run_cases.MONTH_KEYS]
        while values and values[-1] == "Empty":
            values.pop()
        if values:
            months += wrapped(f"PutRow NM_EL_HEADER, {index}, Array({', '.join(values)})")
    return [("TYPE MONTHS", months)] if len(months) > 1 else []


def parameter_lines(fixture: dict[str, Any]) -> list[str]:
    """The statements that set the parameters, as run_cases.write_inputs: the regime always,
    its reference value where the fixture does not set it, and the total own funds where it
    does; ClearInputs has emptied them."""
    parameters = fixture.get("parameters", {})
    for key in parameters.keys() - run_cases.PARAMETERS.keys():
        raise SuiteError(f"{fixture['id']}: no cell for the parameter {key}")
    lines = []
    for key, (name, default) in run_cases.PARAMETERS.items():
        value = parameters.get(key, default)
        if value is not None:
            what = key.removesuffix("_code").removesuffix("_amount").replace("_", " ")
            lines += [f"'Set the {what}", f'SetNamed "{name}", {cell_literal(key, value)}']
    return lines


def reader(fixture: dict[str, Any], output: dict[str, Any]) -> tuple[str, bool]:
    """The expression that reads an expected output, as run_cases.actual does, and
    whether it is a number already (True) or the value of a cell (False)."""
    quantity, scope = output["quantity"], output["scope"]
    results_first = run_cases.RESULTS_ROW["IR"] - 1
    if (cell := run_cases.eligibility_cell(quantity)) is not None:
        row, column = cell
        if scope is None:
            return (f"BlockCell({row - run_cases.BLOCK_ROW}, "
                    f"{column - run_cases.BLOCK_COLUMN})"), False
        index = [item["id"] for item in fixture["datasets"]["DS-11"]].index(scope["record"]) + 1
        return f"RowCell(NM_EL_HEADER, {index}, {column - FIRST_COLUMN})", False
    if scope is None:
        if quantity == run_cases.RUN_STATUS:
            return "RunLine()", False
        if quantity in run_cases.NAMED:
            return f"NamedCell({NAMES[run_cases.NAMED[quantity]]})", False
        if quantity in run_cases.RESULTS_CELLS:
            row, column = run_cases.RESULTS_CELLS[quantity]
            return f"ResultsCell({row - results_first}, {column - FIRST_COLUMN})", False
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


def assertion_groups(fixture: dict[str, Any], expected: dict[str, Any]
                     ) -> tuple[list[tuple[str, list[str]]], bool]:
    """The statements of a case that compare the expected outputs, one group for each
    output with the intent that states it, and whether one compares a number."""
    case_id = fixture["id"]
    groups = []
    numeric = False
    for name, output in zip(assertion_names(expected), expected["outputs"], strict=True):
        expression, is_number = reader(fixture, output)
        scope = output["scope"]
        what = output["quantity"] + ("" if scope is None else f" of {scope['record']}")
        if "text" in output:
            groups.append((f"Compare {what} with its expected text",
                           wrapped(f'TEST_Harness.AssertEqualText "{name}", '
                                   f"{text_literal(output['text'])}, PlainText({expression})")))
            continue
        numeric = True
        wanted = float(output["value"])
        if not is_number:
            # A cell that is empty holds a zero, and only a zero: as run_cases.agrees.
            expression = f"{'ZeroIfEmpty' if wanted == 0 else 'NumberOf'}({expression})"
        tolerance = output["tolerance"]
        where = "vbNullString, vbNullString" if scope is None else \
            f"{text_literal(scope['dataset'])}, {text_literal(scope['record'])}"
        groups.append((f"Compare {what} with its expected value and record it", [
            *wrapped(f"Actual = {expression}"),
            *wrapped(f'TEST_Harness.AssertNear "{name}", {number_literal(wanted)}, Actual, '
                     f"{number_literal(tolerance['absolute'])}, "
                     f"{number_literal(tolerance['relative'])}"),
            *wrapped(f"TEST_Harness.RecordResult {text_literal(case_id)}, vbNullString, "
                     f"{text_literal(output['quantity'])}, {where}, Actual")]))
    return groups, numeric


def case_lines(fixture: dict[str, Any], expected: dict[str, Any]) -> tuple[list[str], int, bool]:
    """The statements of a case between keeping the inputs and putting them back,
    the assertions among them, and whether one compares a number."""
    groups, numeric = assertion_groups(fixture, expected)
    lines = [f"Inputs{pascal(fixture['id'])}", "ResultText = RunSSA_Silent()"]
    for _, statements in groups:
        lines += statements
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


FIXED_HEADER = [
    ("PURPOSE",
     "The cases suite of the workbook's tests: every registered test case of the "
     "project's calculations, as a case of TEST_Harness that types the test case's inputs "
     "on the input sheets, runs the calculation through RunSSA_Silent and compares every "
     "expected output with the cell that holds it. GENERATED by "
     "tools/prototype/case_suite.py from tests/fixtures/ and tests/expected/: do not edit "
     "it; change the test case and write the module again."),
    ("WHY THIS EXISTS",
     "The inputs and the expected values of the test cases are data files; the suite "
     "holds them as literals of this module, so that Excel reads no file and the "
     "calculation is checked in the workbook itself."),
    ("INPUTS", "None at module level"),
    ("RETURNS", "Nothing at module level"),
    ("BEHAVIOR",
     "Gives three entry points:\n"
     "  - RunCaseTests is the documented test entry point\n"
     "  - Test_RunCases runs the same suite and returns its result line where a caller "
     "cannot read an error, as tools/devbuild/lo_macro.py\n"
     "  - CertifyCases certifies the run and saves its log, its results and its "
     "certification log as evidence of an Excel run\n"
     "\n"
     "Each case replaces the inputs of the workbook with those of its test case: the "
     "reporting date and currency, the regime and the total own funds, the FX table, the "
     "option charges, the four position sheets and the month ends of Eligibility. It "
     "rewrites the output sheets through RunSSA_Silent, puts back "
     "the inputs the workbook had, on the success and the error path, and asserts it. "
     "It clears the output sheets when it puts the inputs back."),
    ("ERROR POLICY",
     "As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED, or the "
     "unexpected error that escaped the runner. A refused start raises "
     "TEST_Harness.ERR_TEST_DIRTY_START. An unexpected error inside a case is recorded "
     "as a failure of that case, and the run continues. Test_RunCases contains those "
     "errors; CertifyCases records them, saves the log, and raises only when the "
     "certification did not pass."),
    ("DEPENDENCIES",
     "TEST_Harness; M_Main, CORE_Config and CORE_Util; the workbook's names and "
     "sheets.\n"
     "\n"
     "Reference: Each expected value, its class and its tolerance are those of the "
     "test case's file in tests/expected/ (docs/methodology/TEST_CASES.md)."),
    ("NOTES",
     "Option Explicit requires explicit declarations. Option Private Module keeps the "
     "entry points out of the workbook's external surface and the tests inside this "
     "project. Never part of the production import set.\n"
     "\n"
     "State ownership: owns the copy of the inputs a case keeps while it runs, their "
     "values and their formulas (mRows, mKept and mFormulas).\n"
     "\n"
     "Usage: in a development workbook, type TEST_Cases.RunCaseTests in the Immediate "
     "window. To record an Excel run, type\n"
     '    TEST_Cases.CertifyCases "<folder>", "<checkout>\\tests\\fixtures"\n'
     "(INSTALLATION.md, Recording an Excel run).\n"
     "\n"
     "Compatibility: Excel VBA; also runs under LibreOffice's VBA mode through the "
     "tools of tools/devbuild/, which is never Excel evidence.\n"
     "\n"
     "Author: Daniele Penza"),
]


def entry_points(cases: int, assertions: int) -> list[str]:
    lines = group("ENTRY POINTS")
    lines += procedure(["Public Sub RunCaseTests()"], "RUN CASES SUITE", fields(
        "Run the cases suite. A passing run ends with\n"
        f"RESULT=PASS; suite={SUITE}; operation=PASS; completeness=COMPLETE;\n"
        f"cases={cases}; assertions={assertions}; failures=0; cleanup=PASS; "
        "recovery=NONE.",
        "The documented entry point of the suite: it raises when the run does not pass, "
        "as TEST_Harness.EndRun does.",
        "Calls RunSuite, which runs every case between BeginRun and EndRun.",
        error="As TEST_Harness.EndRun: raises when the run does not pass.",
        deps="RunSuite; TEST_Harness",
        notes="Replaces the inputs for each case and puts them back; leaves the output "
              "sheets cleared."),
        [], [("RUN SUITE", ["'Run every case, which raises when the run does not pass",
                            "RunSuite"])])
    lines += procedure(["Public Function Test_RunCases() As String"], "RUN CASES FOR RESULT",
        fields(
            "Run the cases suite for a caller that reads a result and not an error.",
            "A caller that cannot read an error, as tools/devbuild/lo_macro.py, still "
            "needs the outcome of the run.",
            "Runs the suite with the error of the run suppressed, then returns the result "
            "line that TEST_Harness keeps.",
            returns="The RESULT line of the run, as TEST_Harness.LastResult gives it.",
            error="Contains the error of a run that does not pass, or of a refused start: "
                  "the result line says so.",
            deps="RunSuite; TEST_Harness (LastResult)",
            notes="State as RunCaseTests."),
        [], [("RUN SUITE", ["'Contain the error of a run that does not pass or of a refused start",
                            "On Error Resume Next", "RunSuite", "On Error GoTo 0"]),
             ("RETURN RESULT", ["'Return the result line of the run",
                                "Test_RunCases = TEST_Harness.LastResult()"])])
    lines += procedure(["Public Sub CertifyCases( _", "    ByVal Folder As String, _",
                        "    ByVal Fixtures As String)"], "CERTIFY CASES SUITE", fields(
        "Certify the cases suite's run as evidence of an Excel run: run it, save its log as "
        f"{SUITE}.log and its results as {SUITE}-results.json, and write "
        f"{SUITE}-certification.log.",
        "Evidence of an Excel run is its log, its results and its certification log; "
        "this entry point makes the three files from one run.",
        "Begins the certification, runs the suite with its error recorded and not raised, "
        "saves the log, certifies the run, saves the results and ends the certification.",
        inputs=arguments(
            ("Folder", "String: an existing folder for the files, without a trailing "
                       "separator."),
            ("Fixtures", "String: the full path of tests/fixtures/ in the checkout. Not read, "
                         "because the inputs are in this module; taken so that every "
                         "certification entry point has the same two arguments.")),
        error="The error of the run is recorded, not raised, so that the log is always "
              "saved. Raises ERR_TEST_FAILED after writing the certification log when the "
              "certification did not pass. An error of BeginCertification, SaveLog or "
              "SaveResults propagates; run TEST_Harness.ResetTests before trying again.",
        deps="RunSuite; TEST_Harness (BeginCertification, SaveLog, Echo, CertifyRun, "
             "SaveResults, EndCertification)",
        notes="State as RunCaseTests."),
        [("RunErrorNumber", "Long", "Error that ended the run; 0 for none"),
         ("RunErrorDescription", "String", "Its description")], [
        ("BEGIN CERTIFICATION", ["'Start the certification of the suite",
                                 "TEST_Harness.BeginCertification SUITE_NAME"]),
        ("RUN SUITE", ["'Run the suite and keep the error that ended it, so that the log is saved",
                       "On Error Resume Next", "RunSuite",
                       "RunErrorNumber = Err.Number",
                       "RunErrorDescription = Err.Description", "On Error GoTo 0"]),
        ("SAVE LOG", ["'Save the log and say where it is",
                      f'TEST_Harness.SaveLog Folder & "\\{SUITE}.log"',
                      f'TEST_Harness.Echo "SAVED=" & Folder & "\\{SUITE}.log; mode=NORMAL; error=" & _',
                      '    CStr(RunErrorNumber) & IIf(RunErrorNumber = 0, vbNullString, " " & '
                      "RunErrorDescription)"]),
        ("CERTIFY RUN", ["'Certify the normal run with the error that ended it",
                         'TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, '
                         "RunErrorDescription"]),
        ("SAVE RESULTS", ["'Save the results the cases recorded",
                          f'TEST_Harness.SaveResults Folder & "\\{SUITE}-results.json"']),
        ("END CERTIFICATION", ["'Write the certification log, which raises when the certification failed",
                               f'TEST_Harness.EndCertification Folder & "\\{SUITE}-certification.log"'])])
    return lines


def runner(names: list[str]) -> list[str]:
    return group("RUNNER") + procedure(["Private Sub RunSuite()"], "RUN CASES", fields(
        "Run every case in a fixed order between BeginRun and EndRun.",
        "The three entry points run the suite the same way: the order of the cases and "
        "the report of the run are in one place.",
        "Begins the run with the number of cases and of assertions it expects, lets the "
        "suite write the sheets, calls every case in a fixed order and ends the run with "
        "the error that escaped a case, if any.",
        error="An error escaping a case handler is passed to EndRun, which reports it and "
              "re-raises it unchanged. BeginRun's refusal propagates directly.",
        deps="TEST_Harness (BeginRun, EndRun); AllowWritingToSheets; every Case procedure "
             "of this module; the constants SUITE_NAME, EXPECTED_CASES and "
             "EXPECTED_ASSERTIONS"),
        [("ErrNumber", "Long", "Runner error number; 0 for none"),
         ("ErrSource", "String", "Runner error source"),
         ("ErrDescription", "String", "Runner error description")], [
        ("BEGIN RUN", ["'Begin the run; a refused start raises here, before the handler is on",
                       'TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, '
                       "EXPECTED_ASSERTIONS"]),
        ("RUN CASES", ["'Keep an error that escapes a case, to report it at the end of the run",
                       "On Error GoTo RunFailed",
                       "'Let the suite write the sheets of the workbook",
                       "AllowWritingToSheets",
                       "'Run every case, in a fixed order",
                       *names]),
        ("END RUN", ["Finish:",
                     "'Stop catching errors before the run ends",
                     "On Error GoTo 0",
                     "'End the run: it raises when the run did not pass",
                     "TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription"]),
        ("EXIT PROCEDURE", ["'Leave before the error handler", "Exit Sub"]),
        ("ERROR HANDLER", ["RunFailed:",
                           "'Keep the error that escaped a case",
                           "ErrNumber = Err.Number",
                           "ErrSource = Err.Source",
                           "ErrDescription = Err.Description",
                           "'Go on to the end of the run, which reports the error",
                           "Resume Finish"])])


def helpers() -> list[str]:
    """The procedures of the fixed text of the module that the cases call."""
    lines = group("HELPERS")
    lines += procedure(["Private Sub AllowWritingToSheets()"], "ALLOW WRITING TO SHEETS", fields(
        "Let the suite write the sheets of its workbook, which the template protects: "
        "apply UserInterfaceOnly to each of them, as AllowWriting does for the sheets "
        "that RunSSA writes.",
        "The cases type their inputs on sheets that the template protects.",
        "Calls AllowWriting for every worksheet of the workbook.",
        error="None of its own; an error of AllowWriting propagates.",
        deps="AllowWriting of M_Main; ThisWorkbook.Worksheets",
        notes="Changes the protection mode of every sheet, not who may change it."),
        [("Sheet", "Worksheet", "One sheet")], [
        ("ALLOW WRITING", ["'Apply UserInterfaceOnly to every sheet of the workbook",
                           "For Each Sheet In ThisWorkbook.Worksheets",
                           "    AllowWriting Sheet",
                           "Next Sheet"])])
    lines += procedure(["Private Function InputArea( _", "    ByVal Index As Long) _",
                        "    As Range"], "GET INPUT AREA", fields(
        "One of the areas that hold the inputs a test case replaces.",
        "A case keeps, clears, types and restores its inputs area by area, and the "
        "areas are numbered so that one loop can visit them all.",
        "Returns the range of the area that Index names: the table of a position sheet "
        "with the rows PositionRows gives, the FX table, the reporting date, the "
        "reporting currency, the option charge of a risk class, the regime, the total own "
        "funds or the month ends of Eligibility.",
        inputs=arguments(("Index", "Long: 1 to 4: the rows of the four position sheets, "
                                   "as many as PositionRows gives; 5: the FX table, with its "
                                   "descriptions; 6 and 7: the reporting date and currency; "
                                   "8 to 11: the option charges; 12 and 13: the regime and "
                                   "the total own funds; 14: the input columns of the rows "
                                   "of Eligibility.")),
        returns="The range of the area. Nothing for an Index outside 1 to 14, since the "
                "selection has no other branch.",
        deps="PositionHeader; PositionRows; the IR_, EQ_ and FX_ and CO_INPUT_WIDTH, "
             "NM_EL_HEADER, ELIGIBILITY_ROWS and EL_INPUT_WIDTH constants of CORE_Config; "
             "the workbook's names"),
        [("Header", "Range", "Header of a position sheet"),
         ("Codes", "Range", "FX table codes"), ("Rates", "Range", "FX table rates")], [
        ("SELECT AREA", [
            "'Pick the area by its number: position sheets, FX table, date, currency, charges",
            "Select Case Index",
            "    Case 1: Set Header = PositionHeader(1)",
            "            Set InputArea = Header.Offset(1, 0).Resize(PositionRows(1), IR_INPUT_WIDTH)",
            "    Case 2: Set Header = PositionHeader(2)",
            "            Set InputArea = Header.Offset(1, 0).Resize(PositionRows(2), EQ_INPUT_WIDTH)",
            "    Case 3: Set Header = PositionHeader(3)",
            "            Set InputArea = Header.Offset(1, 0).Resize(PositionRows(3), CO_INPUT_WIDTH)",
            "    Case 4: Set Header = PositionHeader(4)",
            "            Set InputArea = Header.Offset(1, 0).Resize(PositionRows(4), FX_INPUT_WIDTH)",
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
            '    Case 12: Set InputArea = ThisWorkbook.Names("Regime").RefersToRange',
            '    Case 13: Set InputArea = ThisWorkbook.Names("TotalOwnFunds").RefersToRange',
            "    Case 14: Set InputArea = ThisWorkbook.Names(NM_EL_HEADER).RefersToRange _",
            "                .Offset(1, 0).Resize(ELIGIBILITY_ROWS, EL_INPUT_WIDTH)",
            "End Select"])])
    lines += procedure(["Private Function PositionHeader( _", "    ByVal Index As Long) _",
                        "    As Range"], "GET POSITION HEADER", fields(
        "The header cell of one of the four position sheets.",
        "Four sheets hold positions and each has its own header name; callers hold the "
        "sheet as a number.",
        "Maps the number to the workbook name of the sheet's header cell and returns the "
        "cell that name refers to.",
        inputs=arguments(("Index", "Long: 1 interest rate, 2 equity, 3 commodity, "
                                   "4 foreign exchange.")),
        returns="The header cell. Nothing for an Index outside 1 to 4, since the selection "
                "has no other branch.",
        deps="The NM_IR_POS_HEADER, NM_EQ_POS_HEADER, NM_CO_POS_HEADER and "
             "NM_FX_POS_HEADER names of CORE_Config"),
        [], [("SELECT HEADER", ["'Take the header name of the sheet the number selects",
                                "Select Case Index",
                                "    Case 1: Set PositionHeader = ThisWorkbook.Names(NM_IR_POS_HEADER).RefersToRange",
                                "    Case 2: Set PositionHeader = ThisWorkbook.Names(NM_EQ_POS_HEADER).RefersToRange",
                                "    Case 3: Set PositionHeader = ThisWorkbook.Names(NM_CO_POS_HEADER).RefersToRange",
                                "    Case 4: Set PositionHeader = ThisWorkbook.Names(NM_FX_POS_HEADER).RefersToRange",
                                "End Select"])])
    lines += procedure(["Private Function PositionRows( _", "    ByVal Index As Long) _",
                        "    As Long"], "GET POSITION ROWS", fields(
        "The rows of a position sheet a test case replaces and puts back: those of the "
        "table, from the header to the last row that holds a PositionID, as the engine "
        "reads it, and at least INPUT_ROWS.",
        "The table of a position sheet has no fixed length, and a case must put back "
        "every row it replaced.",
        "Returns the row count that KeepInputs stored for the sheet, or INPUT_ROWS "
        "before the first KeepInputs.",
        inputs=arguments(("Index", "Long: as PositionHeader.")),
        returns="The rows kept by the last KeepInputs; INPUT_ROWS before the first.",
        deps="mRows; INPUT_ROWS"),
        [], [("RETURN ROWS", ["'Return the rows kept, or the least number of rows before the first keep",
                              "If mRows(Index) > 0 Then",
                              "    PositionRows = mRows(Index)",
                              "Else",
                              "    PositionRows = INPUT_ROWS",
                              "End If"])])
    lines += procedure(["Private Function TableRows( _", "    ByVal Index As Long) _",
                        "    As Long"], "COUNT TABLE ROWS", fields(
        "The rows of a position sheet that hold the table now: from the row under the "
        "header to the last row that holds a PositionID, as M_Main reads a table, and at "
        "least INPUT_ROWS.",
        "KeepInputs keeps a position sheet through the rows its table holds, whatever "
        "their number.",
        "Finds the last row of the header's column that holds a value, takes its distance "
        "from the header and raises it to INPUT_ROWS if it is less.",
        inputs=arguments(("Index", "Long: as PositionHeader.")),
        returns="The number of rows, at least INPUT_ROWS.",
        deps="PositionHeader; INPUT_ROWS"),
        [("Header", "Range", "Header cell"), ("Sheet", "Worksheet", "Sheet of the table"),
         ("Last", "Long", "Last row under the header that holds a PositionID")], [
        ("LOCATE TABLE", ["'Take the header cell and the sheet that holds the table",
                          "Set Header = PositionHeader(Index)",
                          "Set Sheet = Header.Worksheet"]),
        ("COUNT ROWS", ["'Count the rows from the header to the last row of its column",
                        "Last = Sheet.Cells(Sheet.Rows.Count, Header.Column).End(xlUp).Row - Header.Row",
                        "'Never fewer rows than the least number kept",
                        "If Last < INPUT_ROWS Then Last = INPUT_ROWS",
                        "TableRows = Last"])])
    lines += procedure(["Private Sub KeepInputs()"], "KEEP INPUTS", fields(
        "Keep what every input area holds, before a case replaces it: its values and its "
        "formulas. A position sheet is kept through the rows of its table, whatever "
        "their number.",
        "A case replaces the inputs of the workbook and must put them back as they were.",
        "Stores the rows of each position sheet, then the values and the formulas of "
        "every input area.",
        deps="TableRows; InputArea; AsTable; INPUT_AREAS",
        notes="State ownership: sets mRows, mKept and mFormulas."),
        [("i", "Long", "Input area")], [
        ("KEEP ROWS", ["'Keep the rows that the table of each position sheet holds",
                       "For i = 1 To 4",
                       "    mRows(i) = 0",
                       "    mRows(i) = TableRows(i)",
                       "Next i"]),
        ("KEEP AREAS", ["'Keep the values and the formulas of every input area",
                        "For i = 1 To INPUT_AREAS",
                        "    mKept(i) = InputArea(i).Value",
                        "    mFormulas(i) = AsTable(InputArea(i).Formula)",
                        "Next i"])])
    lines += procedure(["Private Function RestoreInputs() As String"], "RESTORE INPUTS", fields(
        "Put the sheets back as they were before a case: write what KeepInputs kept, "
        "values and then formulas, check that every input cell holds it, and clear the "
        "output sheets, which show the results of the test case.",
        "A case changes the inputs and the outputs of the workbook; it puts the sheets "
        "back as they were, and the suite asserts it.",
        "Empties and rewrites every input area with its kept values, writes back the "
        "formulas, writes again as a text each kept text that is not a formula, counts "
        "the input cells that differ from what was kept and clears the output sheets.",
        returns="An empty text when the sheets are restored; otherwise what failed.",
        error="Contains its own error and returns it, so that the case asserts the result "
              "on the success and the error path.",
        deps="InputArea; AsTable; IsFormula; PutValue; InputDifferences; "
             "ClearOutputs_Silent of M_Main; mKept and mFormulas",
        notes="Writes every input area and clears the output sheets. Each area is emptied "
              "before it is written: an empty value written over a cell that holds "
              "something does not leave it empty in every host."),
        [("i", "Long", "Input area"), ("r", "Long", "Row"), ("c", "Long", "Column"),
         ("Differences", "Long", "Input cells that differ"),
         ("Cleared", "String", "ClearOutputs_Silent result line")], [
        ("INITIALIZE", ["'Contain an error of the restoration and return it",
                        "On Error GoTo Failed"]),
        ("WRITE INPUTS", [
            "'Empty each area before writing it, as an empty value does not clear a cell in every host",
            "For i = 1 To INPUT_AREAS",
            "    InputArea(i).ClearContents",
            "    InputArea(i).Value = mKept(i)",
            "    'A cell that held a formula gets it back: the value above is its result",
            "    For r = 1 To UBound(mFormulas(i), 1)",
            "        For c = 1 To UBound(mFormulas(i), 2)",
            "            If IsFormula(mFormulas(i)(r, c)) Then",
            "                InputArea(i).Cells(r, c).Formula = mFormulas(i)(r, c)",
            "            End If",
            "        Next c",
            "    Next r",
            "    'A text the write above made a number, a date or a logical value is a text again",
            "    For r = 1 To UBound(mFormulas(i), 1)",
            "        For c = 1 To UBound(mFormulas(i), 2)",
            "            If VarType(AsTable(mKept(i))(r, c)) = vbString Then",
            "                If Not IsFormula(mFormulas(i)(r, c)) Then",
            "                    PutValue InputArea(i).Cells(r, c), AsTable(mKept(i))(r, c)",
            "                End If",
            "            End If",
            "        Next c",
            "    Next r",
            "Next i"]),
        ("CHECK AND CLEAR", [
            "'Count the input cells that do not hold what was kept",
            "Differences = InputDifferences()",
            "'Clear the output sheets, which show the results of the test case",
            "Cleared = ClearOutputs_Silent()",
            "'Report what failed first: the inputs, then the outputs",
            "If Differences <> 0 Then",
            '    RestoreInputs = CStr(Differences) & " input cell(s) do not hold what they held"',
            'ElseIf Left$(Cleared, 14) <> "RESULT=CLEARED" Then',
            '    RestoreInputs = "the output sheets are not cleared: " & Cleared',
            "End If"]),
        ("EXIT PROCEDURE", ["'Leave before the error handler", "Exit Function"]),
        ("ERROR HANDLER", [
            "Failed:",
            "'Return the error as the result of the restoration",
            'RestoreInputs = "error " & CStr(Err.Number) & " while restoring: " & Err.Description'])])
    lines += procedure(["Private Function InputDifferences() As Long"], "COUNT DIFFERENCES", fields(
        "Count the input cells that do not hold what KeepInputs kept: a value that "
        "differs, a formula that differs, a formula where the cell held none and none "
        "where it held one.",
        "RestoreInputs asserts that the sheets are as they were; this is the count it "
        "asserts.",
        "Reads every input area, values and formulas, and compares each cell with what "
        "KeepInputs kept.",
        returns="The number of cells; 0 when the inputs are those kept.",
        deps="InputArea; AsTable; SameValue; IsFormula; mKept and mFormulas"),
        [("Current", "Variant", "What an area holds now: values"),
         ("CurrentFormulas", "Variant", "What an area holds now: formulas"),
         ("i", "Long", "Input area"), ("r", "Long", "Row"), ("c", "Long", "Column"),
         ("Count", "Long", "Cells that differ")], [
        ("COMPARE AREAS", [
            "'Compare every cell of every area with what was kept",
            "For i = 1 To INPUT_AREAS",
            "    Current = AsTable(InputArea(i).Value)",
            "    CurrentFormulas = AsTable(InputArea(i).Formula)",
            "    For r = 1 To UBound(Current, 1)",
            "        For c = 1 To UBound(Current, 2)",
            "            'Count a differing value, a gained or lost formula, or a differing formula",
            "            If Not SameValue(Current(r, c), AsTable(mKept(i))(r, c)) Then",
            "                Count = Count + 1",
            "            ElseIf IsFormula(CurrentFormulas(r, c)) <> IsFormula(mFormulas(i)(r, c)) Then",
            "                Count = Count + 1",
            "            ElseIf IsFormula(CurrentFormulas(r, c)) Then",
            "                If CurrentFormulas(r, c) <> mFormulas(i)(r, c) Then Count = Count + 1",
            "            End If",
            "        Next c",
            "    Next r",
            "Next i"]),
        ("RETURN COUNT", ["'Return the number of cells that differ", "InputDifferences = Count"])])
    lines += procedure(["Private Function AsTable( _", "    ByVal Value As Variant) _",
                        "    As Variant"], "RETURN TABLE", fields(
        "What a range returned as a table of one or more rows and columns: an area of one "
        "cell returns its value alone.",
        "The Value and Formula of a range of one cell are not arrays, and the callers "
        "index every area by row and column.",
        "Returns Value when it is an array, else a table of one cell that holds it.",
        inputs=arguments(("Value", "Variant: the Value or the Formula of a range.")),
        returns="Value when it is an array, else a table of one cell holding it.",
        notes="The table of one cell is a local array that is returned by value."),
        [("Table(1 To 1, 1 To 1)", "Variant", "The table of one cell")], [
        ("RETURN TABLE", ["'Return an array as it is and any other value as a table of one cell",
                          "If IsArray(Value) Then",
                          "    AsTable = Value",
                          "Else",
                          "    Table(1, 1) = Value",
                          "    AsTable = Table",
                          "End If"])])
    lines += procedure(["Private Function IsFormula( _", "    ByVal Value As Variant) _",
                        "    As Boolean"], "TEST FORMULA", fields(
        "Whether what a cell holds, as its Formula returns it, is a formula.",
        "The Formula of a cell returns its text or its value, and the cell holds a "
        "formula only when that text starts with an equals sign.",
        "Tests that Value is a text whose first character is an equals sign.",
        inputs=arguments(("Value", "Variant: the Formula of a cell.")),
        returns="True for a text that starts with an equals sign, else False."),
        [], [("TEST VALUE", ["'A formula is a text that starts with an equals sign",
                             "If VarType(Value) = vbString Then",
                             '    IsFormula = (Left$(Value, 1) = "=")',
                             "End If"])])
    lines += procedure(["Private Function SameValue( _", "    ByVal First As Variant, _",
                        "    ByVal Second As Variant) _", "    As Boolean"], "COMPARE VALUES", fields(
        "Whether two cell values are the same value of the same type.",
        "The check of the restored inputs compares what a cell holds now with what was "
        "kept, and a value must be equal in type as well as in content.",
        "Compares the types first, then treats two empty values as equal, two error "
        "values by their text and any other pair by comparison.",
        inputs=arguments(("First", "Variant: a cell value."), ("Second", "Variant: a cell value.")),
        returns="True when the values are the same, else False.",
        error="An error value equals only the same error value."),
        [], [("COMPARE", ["'Values of another type differ; compare the others by their kind",
                          "If VarType(First) <> VarType(Second) Then",
                          "    SameValue = False",
                          "ElseIf IsEmpty(First) Then",
                          "    SameValue = True",
                          "ElseIf IsError(First) Then",
                          "    SameValue = (CStr(First) = CStr(Second))",
                          "Else",
                          "    SameValue = (First = Second)",
                          "End If"])])
    lines += procedure(["Private Sub ClearInputs()"], "CLEAR INPUTS", fields(
        "Empty every input area a test case fills: the position sheets, the FX table, the "
        "option charges, the total own funds and the month ends. The reporting date and "
        "currency and the regime are always set by the case.",
        "A case replaces the inputs of the workbook with those of its test case, so the "
        "areas it fills must be empty before it types.",
        "Clears the contents of every input area except 6, 7 and 12, the reporting date "
        "and currency and the regime.",
        deps="InputArea; INPUT_AREAS",
        notes="Clears the contents of those areas; their formats stay."),
        [("i", "Long", "Input area")], [
        ("CLEAR AREAS", ["'Empty every area but the reporting date and currency and the regime",
                         "For i = 1 To INPUT_AREAS",
                         "    If i <> 6 And i <> 7 And i <> 12 Then InputArea(i).ClearContents",
                         "Next i"])])
    lines += procedure(["Private Sub SetNamed( _", "    ByVal CellName As String, _",
                        "    ByVal Value As Variant)"], "SET NAMED VALUE", fields(
        "Type a value in the cell a workbook name refers to.",
        "The reporting date, the reporting currency, the option charges, the regime and "
        "the total own funds are cells that the workbook names.",
        "Types Value in the cell that CellName refers to, through PutValue.",
        inputs=arguments(("CellName", "String: a workbook name that refers to one cell."),
                         ("Value", "Variant: what the cell holds; as PutValue.")),
        error="An error of the name or of PutValue propagates.",
        deps="PutValue"),
        [], [("TYPE VALUE", ["'Type the value in the cell the name refers to",
                             "PutValue ThisWorkbook.Names(CellName).RefersToRange, Value"])])
    lines += procedure(["Private Sub PutRate( _", "    ByVal Index As Long, _",
                        "    ByVal Code As String, _", "    ByVal Rate As Variant)"],
                       "PUT FX RATE", fields(
        "Type one row of the FX table: a currency and its rate.",
        "The FX table of a test case is typed one row at a time.",
        "Types the currency code in the row Index of the FX table and, unless Rate is "
        "Empty, its rate beside it.",
        inputs=arguments(("Index", "Long: row of the table, from 1."),
                         ("Code", "String: currency code."),
                         ("Rate", "Variant: its rate; Empty leaves the cell empty.")),
        deps="PutValue; the NM_FX_CCY and NM_FX_RATE names of CORE_Config"),
        [], [("TYPE CODE", ["'Type the currency code in the row of the table",
                            "PutValue ThisWorkbook.Names(NM_FX_CCY).RefersToRange.Cells(Index, 1), Code"]),
             ("TYPE RATE", ["'Type the rate, unless it is Empty",
                            "If Not IsEmpty(Rate) Then",
                            "    PutValue ThisWorkbook.Names(NM_FX_RATE).RefersToRange.Cells(Index, 1), Rate",
                            "End If"])])
    lines += procedure(["Private Sub PutRow( _", "    ByVal HeaderName As String, _",
                        "    ByVal Index As Long, _", "    ByVal Values As Variant)"],
                       "PUT INPUT ROW", fields(
        "Type one row of a position sheet or of Eligibility.",
        "The positions and the month ends of a test case are typed one row at a time, "
        "from the first column of the table on.",
        "Finds the row Index under the header and types each value of Values in the "
        "cells of the row, from the header's column to the right.",
        inputs=arguments(("HeaderName", "String: workbook name of the sheet's header cell."),
                         ("Index", "Long: row under the header, from 1."),
                         ("Values", "Variant: the cells of the row from the header's column: "
                                    "the PositionID, or the month end; Empty leaves a cell "
                                    "empty.")),
        deps="PutValue"),
        [("Target", "Range", "First cell of the row"), ("c", "Long", "Cell of the row")], [
        ("LOCATE ROW", ["'Take the first cell of the row under the header",
                        "Set Target = ThisWorkbook.Names(HeaderName).RefersToRange.Offset(Index, 0)"]),
        ("TYPE CELLS", ["'Type each value that is not Empty in its cell of the row",
                        "For c = LBound(Values) To UBound(Values)",
                        "    If Not IsEmpty(Values(c)) Then",
                        "        PutValue Target.Offset(0, c - LBound(Values)), Values(c)",
                        "    End If",
                        "Next c"])])
    lines += procedure(["Private Sub PutValue( _", "    ByVal Target As Range, _",
                        "    ByVal Value As Variant)"], "PUT VALUE", fields(
        "Type a value in a cell so that the cell holds it as the test case gives it: a "
        "text stays a text.",
        "A text that reads as a number, a date or a logical value would become one when "
        "it is written. It is written with the cell formatted as text, and the cell gets "
        "its format back; typing an apostrophe before the text would leave the "
        "apostrophe in a cell that is already formatted as text.",
        "Writes a value that is not a text as it is. For a text, saves the number format "
        "of the cell, writes the text as it is when the cell is already formatted as "
        "text, and otherwise formats the cell as text for the write and puts the format "
        "back.",
        inputs=arguments(("Target", "Range: the cell."),
                         ("Value", "Variant: what it holds; any other value than a text is "
                                   "written as it is.")),
        error="An error of the write is raised again once the format is back.",
        notes="Changes the number format of Target for the write and puts it back, on "
              "the success and the error path."),
        [("Saved", "String", "Number format of the cell"),
         ("ErrNumber", "Long", "Error of the write; 0 for none"),
         ("ErrSource", "String", "Its source"), ("ErrDescription", "String", "Its description")], [
        ("WRITE OTHER VALUES", ["'Write a value that is not a text as it is",
                                "If VarType(Value) <> vbString Then",
                                "    Target.Value = Value",
                                "    Exit Sub",
                                "End If"]),
        ("CHECK FORMAT", ["'Keep the number format of the cell",
                          "Saved = CStr(Target.NumberFormat)",
                          "'A cell that is formatted as text already holds the text as typed",
                          'If Saved = "@" Then',
                          "    Target.Value = Value",
                          "    Exit Sub",
                          "End If"]),
        ("WRITE TEXT", ["'Keep an error of the write, so that the format is put back",
                        "On Error GoTo Failed",
                        "'Format the cell as text for the write",
                        'Target.NumberFormat = "@"',
                        "Target.Value = Value"]),
        ("RESTORE FORMAT", ["Restored:",
                            "'Stop catching errors before the format is put back",
                            "On Error GoTo 0",
                            "'Put the number format of the cell back",
                            "Target.NumberFormat = Saved",
                            "'Raise the error of the write again, now that the format is back",
                            "If ErrNumber <> 0 Then Err.Raise ErrNumber, ErrSource, ErrDescription"]),
        ("EXIT PROCEDURE", ["'Leave before the error handler", "Exit Sub"]),
        ("ERROR HANDLER", ["Failed:",
                           "'Keep the error of the write",
                           "ErrNumber = Err.Number",
                           "ErrSource = Err.Source",
                           "ErrDescription = Err.Description",
                           "'Go on to put the format back",
                           "Resume Restored"])])
    lines += procedure(["Private Function RowCell( _", "    ByVal HeaderName As String, _",
                        "    ByVal Index As Long, _", "    ByVal ColumnOffset As Long) _",
                        "    As Variant"], "READ ROW CELL", fields(
        "The value of one cell of a calculation sheet.",
        "The expected outputs of a position are cells of its row on a calculation sheet.",
        "Reads the cell that lies Index rows under the header and ColumnOffset columns "
        "to the right of it.",
        inputs=arguments(("HeaderName", "String: workbook name of the sheet's header cell."),
                         ("Index", "Long: row under the header, from 1: the position's row."),
                         ("ColumnOffset", "Long: columns to the right of the PositionID.")),
        returns="The value of the cell, as Range.Value gives it."),
        [], [("READ CELL", ["'Read the cell at the row and column offsets from the header",
                            "RowCell = ThisWorkbook.Names(HeaderName).RefersToRange.Offset(Index, "
                            "ColumnOffset).Value"])])
    lines += procedure(["Private Function ResultsCell( _", "    ByVal RowOffset As Long, _",
                        "    ByVal ColumnOffset As Long) _", "    As Variant"],
                       "READ RESULTS CELL", fields(
        "The value of one cell of the table of Results.",
        "The totals and the requirements of the calculation are cells of the table of "
        "Results.",
        "Reads the cell that lies RowOffset rows under the header of Results and "
        "ColumnOffset columns to the right of it.",
        inputs=arguments(("RowOffset", "Long: rows under the header: 1 interest rate, "
                                       "2 equity, 3 commodity, 4 foreign exchange."),
                         ("ColumnOffset", "Long: columns to the right of the risk class.")),
        returns="The value of the cell, as Range.Value gives it.",
        deps="The NM_RESULTS_HEADER name of CORE_Config"),
        [], [("READ CELL", ["'Read the cell at the row and column offsets from the header of Results",
                            "ResultsCell = ThisWorkbook.Names(NM_RESULTS_HEADER).RefersToRange _",
                            "    .Offset(RowOffset, ColumnOffset).Value"])])
    lines += procedure(["Private Function NamedCell( _", "    ByVal CellName As String) _",
                        "    As Variant"], "READ NAMED CELL", fields(
        "The value of the cell a workbook name refers to.",
        "Some outputs of the calculation are cells that the workbook names.",
        "Reads the cell that CellName refers to.",
        inputs=arguments(("CellName", "String: a workbook name that refers to one cell.")),
        returns="The value of the cell, as Range.Value gives it."),
        [], [("READ CELL", ["'Read the cell the name refers to",
                            "NamedCell = ThisWorkbook.Names(CellName).RefersToRange.Value"])])
    lines += procedure(["Private Function RunLine() _", "    As String"], "READ RUN STATUS", fields(
        "The status line of Results, without the time of the run at its end.",
        "The status line ends with the time of the run, which differs from one run to the "
        "next and no case compares.",
        "Reads the cell of the status line as text and cuts it where the last word Run "
        "between two blanks starts, as run_cases.run_line does in the Python runner; a "
        "line without that word is returned whole.",
        returns="The status line, trimmed, without the time of the run.",
        deps="PlainText of CORE_Util; the NM_RESULTS_STATUS name of CORE_Config"),
        [("Shown", "String", "Status line on the sheet"), ("Cut", "Long", "Start of the run time")], [
        ("READ LINE", ["'Read the status line as text",
                       "Shown = PlainText(ThisWorkbook.Names(NM_RESULTS_STATUS).RefersToRange.Value)"]),
        ("CUT RUN TIME", ["'Drop the time of the run that ends the line",
                          'Cut = InStrRev(Shown, " Run ")',
                          "If Cut > 0 Then Shown = Left$(Shown, Cut - 1)",
                          "RunLine = Shown"])])
    lines += procedure(["Private Function BlockCell( _", "    ByVal RowOffset As Long, _",
                        "    ByVal ColumnOffset As Long) _", "    As Variant"],
                       "READ RESULT BLOCK CELL", fields(
        "The value of one cell of the result block of Eligibility.",
        "The status and the restart condition of each approach are cells of the result "
        "block of Eligibility.",
        "Reads the cell that lies RowOffset rows under the first row of the block and "
        "ColumnOffset columns to the right of its first column.",
        inputs=arguments(("RowOffset", "Long: 0 for Article 325a, 1 for Article 94."),
                         ("ColumnOffset", "Long: columns to the right of the status.")),
        returns="The value of the cell, as Range.Value gives it.",
        deps="The NM_EL_RESULTS name of CORE_Config"),
        [], [("READ CELL", ["'Read the cell at the row and column offsets in the result block",
                            "BlockCell = ThisWorkbook.Names(NM_EL_RESULTS).RefersToRange _",
                            "    .Cells(1 + RowOffset, 1 + ColumnOffset).Value"])])
    lines += procedure(["Private Function CurrencyRow( _", "    ByVal HeaderName As String, _",
                        "    ByVal ColumnOffset As Long, _", "    ByVal Code As String) _",
                        "    As Range"], "FIND CURRENCY ROW", fields(
        "The row of a table by currency: the first row under the header whose code is the "
        "one given, without regard to case.",
        "The row of a currency in a table is found by its code, not by its position.",
        "Searches the rows under the header, up to SEARCH_ROWS, for a code equal to Code "
        "once both are normalised by CodeText, and returns the header's cell of the first "
        "row that matches.",
        inputs=arguments(("HeaderName", "String: workbook name of the table's header cell."),
                         ("ColumnOffset", "Long: column of the code, to the right of the "
                                          "header."),
                         ("Code", "String: currency code.")),
        returns="The cell of that row in the header's column.",
        error="Raises error 5 when no row within SEARCH_ROWS holds the code; the case "
              "records it.",
        deps="CodeText of CORE_Util; SEARCH_ROWS"),
        [("Header", "Range", "Header cell"), ("r", "Long", "Row under the header")], [
        ("LOCATE HEADER", ["'Take the header cell of the table",
                           "Set Header = ThisWorkbook.Names(HeaderName).RefersToRange"]),
        ("SEARCH ROWS", ["'Return the row whose code is the one given, whatever its case",
                         "For r = 1 To SEARCH_ROWS",
                         "    If CodeText(Header.Offset(r, ColumnOffset).Value) = CodeText(Code) Then",
                         "        Set CurrencyRow = Header.Offset(r, 0)",
                         "        Exit Function",
                         "    End If",
                         "Next r"]),
        ("RAISE NOT FOUND", ["'No row holds the code",
                             'Err.Raise 5, "TEST_Cases.CurrencyRow", "No row of " & HeaderName & " holds " & Code & "."'])])
    lines += procedure(["Private Function ZonesTotal( _", "    ByVal Code As String, _",
                        "    ByVal ColumnOffsets As Variant) _", "    As Double"], "SUM ZONES CELLS", fields(
        "The sum of some cells of a currency's row of IR_Zones.",
        "Some expected outputs are sums of several columns of the zones table.",
        "Finds the currency's row of the zones table and adds the cells at the column "
        "offsets given, an empty cell as zero.",
        inputs=arguments(("Code", "String: currency code."),
                         ("ColumnOffsets", "Variant: the columns to add, to the right of the "
                                           "currency.")),
        returns="The sum; an empty cell adds nothing.",
        error="As CurrencyRow, and NumberOf for a cell that holds no number.",
        deps="CurrencyRow; ZeroIfEmpty; the NM_ZONES_HEADER name of CORE_Config"),
        [("First", "Range", "The currency's cell"), ("Item", "Variant", "One column offset"),
         ("Total", "Double", "Sum")], [
        ("LOCATE ROW", ["'Find the row of the currency in the zones table",
                        "Set First = CurrencyRow(NM_ZONES_HEADER, 0, Code)"]),
        ("ADD CELLS", ["'Add the cell at each column offset",
                       "For Each Item In ColumnOffsets",
                       "    Total = Total + ZeroIfEmpty(First.Offset(0, CLng(Item)).Value)",
                       "Next Item"]),
        ("RETURN SUM", ["'Return the sum", "ZonesTotal = Total"])])
    lines += procedure(["Private Function FxNet( _", "    ByVal Code As String) _",
                        "    As Variant"], "READ FX NET POSITION", fields(
        "The net position of a currency in the FX block of Buckets, in the reporting "
        "currency.",
        "The net position of a currency is an output of the foreign exchange block of "
        "Buckets.",
        "Finds the currency's row of the FX block and reads the net position column.",
        inputs=arguments(("Code", "String: currency code.")),
        returns="The value of the cell, as Range.Value gives it.",
        error="As CurrencyRow.",
        deps="CurrencyRow; the NM_BK_FX_HEADER name of CORE_Config; FX_BLOCK_CODE and "
             "FX_BLOCK_NET"),
        [], [("READ CELL", ["'Read the net position in the row of the currency",
                            "FxNet = CurrencyRow(NM_BK_FX_HEADER, FX_BLOCK_CODE, Code).Offset(0, FX_BLOCK_NET).Value"])])
    lines += procedure(["Private Function ZeroIfEmpty( _", "    ByVal Value As Variant) _",
                        "    As Double"], "READ ZERO IF EMPTY", fields(
        "The number a cell holds, where the expected value is zero: an empty cell holds a "
        "zero.",
        "An empty cell is read as zero for an output whose expected value is zero, and "
        "only for it, as run_cases.agrees does in the Python runner.",
        "Returns 0 for an empty cell and for an empty text, and otherwise the number "
        "that NumberOf reads.",
        inputs=arguments(("Value", "Variant: what a cell holds.")),
        returns="The number; 0 for an empty cell or an empty text.",
        error="As NumberOf for a cell that is not empty.",
        deps="NumberOf"),
        [], [("READ NUMBER", ["'An empty cell or an empty text is zero; anything else must be a number",
                              "If IsEmpty(Value) Then",
                              "    ZeroIfEmpty = 0#",
                              "ElseIf VarType(Value) = vbString And Len(Value) = 0 Then",
                              "    ZeroIfEmpty = 0#",
                              "Else",
                              "    ZeroIfEmpty = NumberOf(Value)",
                              "End If"])])
    lines += procedure(["Private Function NumberOf( _", "    ByVal Value As Variant) _",
                        "    As Double"], "READ NUMBER", fields(
        "The number a cell holds, for an assertion that compares numbers.",
        "An empty cell, a text or an error value read as 0 would pass a comparison with "
        "zero.",
        "Checks that Value is a number and converts it to a Double.",
        inputs=arguments(("Value", "Variant: what a cell holds.")),
        returns="The number, as a Double.",
        error="Raises error 13 when Value is not a number, so that an empty cell, a text "
              "or an error value is never compared as 0; the case records it.",
        deps="IsNumberValue of CORE_Util"),
        [], [("READ NUMBER", ["'Refuse a value that is not a number, and convert the others",
                              "If Not IsNumberValue(Value) Then",
                              '    Err.Raise 13, "TEST_Cases.NumberOf", "The result is not a number."',
                              "End If",
                              "NumberOf = CDbl(Value)"])])
    return lines


def input_notes(fixture: dict[str, Any], sections: list[tuple[str, list[str]]]) -> str:
    """The BEHAVIOR of an Inputs procedure, from the records of the fixture."""
    datasets = fixture["datasets"]
    items = ["  - the reporting date and the reporting currency",
             "  - the regime" + (", and the total own funds"
                                 if "total_own_funds_amount" in fixture.get("parameters", {}) else "")]
    if datasets.get("DS-05"):
        items.append(f"  - the FX table: {len(datasets['DS-05'])} row(s)")
    if datasets.get("DS-10"):
        items.append(f"  - the option charges: {len(datasets['DS-10'])} risk class(es)")
    for dataset, kind in POSITION_KINDS.items():
        if datasets.get(dataset):
            items.append(f"  - the {kind} positions: {len(datasets[dataset])} row(s)")
    if datasets.get("DS-11"):
        items.append(f"  - the month ends of Eligibility: {len(datasets['DS-11'])} row(s)")
    return "Empties the input areas, then types:\n" + "\n".join(items)


# How a case ends: an error of the case is recorded first, then the sheets are
# restored, which is asserted once, on the success and the error path. The handler
# is off during the cleanup, so that an error there cannot come back to it.
CLEAN_UP = [("CLEAN UP", ["CleanUp:",
                          "'Stop catching errors, so that an error of the cleanup cannot come back here",
                          "On Error GoTo 0",
                          "'Put the sheets back as they were, on the success and the error path",
                          "If RestoreDue Then",
                          "    Detail = RestoreInputs()",
                          "End If",
                          "'Assert once that the sheets are restored, whichever path came here",
                          'TEST_Harness.AssertTrue "inputs-restored", Len(Detail) = 0, Detail']),
            ("EXIT PROCEDURE", ["'Leave before the error handler", "Exit Sub"]),
            ("ERROR HANDLER", ["'Record the error as a failure of the case, then clean up",
                               "CaseFailed:",
                               "TEST_Harness.RecordCaseError",
                               "Resume CleanUp"])]


def case_procedures(fixture: dict[str, Any], expected: dict[str, Any], numeric: bool) -> list[str]:
    """The procedure that types a test case's inputs and the procedure that is the case."""
    name = pascal(fixture["id"])
    date = updated(expected)
    sections = input_sections(fixture)
    lines = procedure([f"Private Sub Inputs{name}()"], "TYPE CASE INPUTS", fields(
        f"Type the inputs of the test case {fixture['id']}.",
        "The inputs of the test case are literals of the module, so that Excel reads no "
        "file; each case types them on the input sheets of the workbook.",
        input_notes(fixture, sections),
        error="An error of a procedure it calls propagates to the case.",
        deps="ClearInputs; SetNamed; PutRate; PutRow",
        notes="Replaces the inputs of the workbook."), [], sections, date)
    dims = [("ResultText", "String", "RunSSA_Silent result line, not compared")]
    if numeric:
        dims.append(("Actual", "Double", "Number read from the workbook"))
    dims += [("RestoreDue", "Boolean", "True once the inputs are kept"),
             ("Detail", "String", "Why the sheets could not be restored")]
    groups, _ = assertion_groups(fixture, expected)
    checks = [line for intent, statements in groups for line in (f"'{intent}", *statements)]
    lines += procedure([f"Private Sub Case{name}()"], "RUN CASE", fields(
        comment(f"Test case {fixture['id']}, on {fixture['calculation']}: "
                f"{fixture['description']}."),
        "A registered test case is one case of the suite: the workbook calculates its "
        "inputs and every expected output is compared with the cell that holds it.",
        f"Keeps the inputs of the workbook, types those of the test case and runs the "
        f"calculation. Compares each of the {len(groups)} expected outputs with the cell "
        "that holds it, a number within its tolerance and a text exactly, and records "
        "every number as a result. Then puts the inputs back and asserts it.",
        error="An unexpected error is recorded as a failure of the case and the run "
              "continues; the inputs are put back on the success and the error path.",
        deps="TEST_Harness (BeginCase, AssertNear, AssertEqualText, RecordResult, "
             f"AssertTrue, RecordCaseError); Inputs{name}; KeepInputs; RestoreInputs; "
             "RunSSA_Silent of M_Main",
        notes="Replaces the inputs of the workbook and rewrites the output sheets; puts "
              "the inputs back and clears the output sheets."), dims, [
        ("INITIALIZE", ["'Record an error of the case as its failure and go on with the cleanup",
                        "On Error GoTo CaseFailed",
                        f'TEST_Harness.BeginCase "{fixture["id"]}"']),
        ("KEEP INPUTS", ["'Keep what the input areas hold, to put it back at the end",
                         "KeepInputs",
                         "'From here the inputs are put back, whichever path ends the case",
                         "RestoreDue = True"]),
        ("RUN CALCULATION", ["'Type the inputs of the test case", f"Inputs{name}",
                             "'Run the calculation; its result line is not compared",
                             "ResultText = RunSSA_Silent()"]),
        ("CHECK OUTPUTS", checks), *CLEAN_UP], date)
    return lines


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
        _, count, numeric = case_lines(fixture, expected)
        assertions += count + 1  # And inputs-restored.
        bodies.append((fixture, expected, numeric))
    newest = max(WRITTEN, *(updated(expected) for _, expected in pairs))
    text = ['Attribute VB_Name = "TEST_Cases"', "", "Option Explicit", "Option Private Module",
            "", *module_header(FIXED_HEADER, newest), "", *section("PRIVATE CONSTANTS"),
            "    'A complete run starts every case and evaluates every assertion once",
            *aligned([("Private Const SUITE_NAME", f'As String = "{SUITE}"',
                       "Name of the suite in its log"),
                      ("Private Const EXPECTED_CASES", f"As Long = {len(pairs)}",
                       "Cases a complete run starts"),
                      ("Private Const EXPECTED_ASSERTIONS", f"As Long = {assertions}",
                       "Assertions a complete run evaluates")], tight=True),
            "", "    'The inputs a test case replaces (InputArea)",
            *aligned([("Private Const INPUT_AREAS", "As Long = 14", "Areas kept and put back"),
                      ("Private Const INPUT_ROWS", f"As Long = {run_cases.CAPACITY}",
                       "Fewest rows of a position sheet kept")]),
            "", "    'Where a result per currency is found",
            *aligned([("Private Const SEARCH_ROWS", "As Long = 60", "Rows searched under a header"),
                      ("Private Const FX_BLOCK_CODE",
                       f"As Long = {run_cases.FX_BLOCK_CODE - FIRST_COLUMN}",
                       "Buckets FX block: currency"),
                      ("Private Const FX_BLOCK_NET",
                       f"As Long = {run_cases.FX_BLOCK_NET - FIRST_COLUMN}",
                       "Buckets FX block: net position")]),
            "", *section("PRIVATE STATE"),
            *aligned([("Private mRows(1 To 4)", "As Long", "Rows of each position sheet kept"),
                      ("Private mKept(1 To 14)", "As Variant",
                       "What each input area held: values"),
                      ("Private mFormulas(1 To 14)", "As Variant",
                       "What each input area held: formulas")]),
            "", ""]
    text += entry_points(len(pairs), assertions)
    text += runner(names)
    text += group("CASES")
    for fixture, expected, numeric in bodies:
        text += case_procedures(fixture, expected, numeric)
    text += helpers()
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
