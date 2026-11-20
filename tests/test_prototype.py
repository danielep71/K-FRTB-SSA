"""Tests of the tools under tools/prototype/.

They need openpyxl, pinned in .github/requirements.txt; install it with
``python -m pip install -r .github/requirements.txt``. Without it the tests
are skipped locally and fail in GitHub Actions. None needs LibreOffice: the
tools that drive it are tested with their subprocess replaced.
"""

from __future__ import annotations

import contextlib
import datetime
import importlib
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
import warnings
import xml.etree.ElementTree as ElementTree
import zipfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools" / "prototype"))
sys.path.insert(0, str(ROOT / "tools" / "devbuild"))

import ssa_reference  # noqa: E402

try:
    import openpyxl
except ImportError:  # pragma: no cover - installed from .github/requirements.txt
    openpyxl = None

TEMPLATE = ROOT / "src" / "workbook" / "Template.xlsx"
# The two timestamps openpyxl writes; the generator fixes them.
CORE_TIMES = re.compile(rb"<dcterms:(created|modified)[^>]*>[^<]*</dcterms:\1>")


def tool(name: str) -> Any:
    """A module of tools/prototype/ that imports openpyxl."""
    return importlib.import_module(name)


def canonical_parts(path: Path) -> dict[str, bytes]:
    """The parts of a package, each XML part in canonical form."""
    parts = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            data = package.read(name)
            if name == "docProps/core.xml":
                data = CORE_TIMES.sub(b"", data)
            if name.endswith((".xml", ".rels")):
                data = ElementTree.canonicalize(data.decode("utf-8")).encode("utf-8")
            parts[name] = data
    return parts


class Reference(unittest.TestCase):
    """The Python re-calculation of the sample data."""

    def test_the_sample_requirement(self) -> None:
        result = ssa_reference.calc_all()
        self.assertAlmostEqual(result["total"], 48_267_441.0, places=6)
        self.assertAlmostEqual(result["rwa"], 603_343_012.5, places=6)
        self.assertAlmostEqual(result["rwa"], 12.5 * result["total"], places=6)
        scaling = ssa_reference.SCALING
        scaled = (scaling["IR"] * result["k_ir"] + scaling["EQ"] * result["k_eq"]
                  + scaling["CO"] * result["k_co"] + scaling["FX"] * result["k_fx"])
        self.assertAlmostEqual(result["total"], scaled, places=6)

    def test_the_fx_sample_is_the_published_example(self) -> None:
        # MAR40.61, Table 9, in millions: longs 300, shorts 200, gold 35.
        fx = ssa_reference.calc_fx()
        self.assertAlmostEqual(fx["longs"], 300e6, places=6)
        self.assertAlmostEqual(fx["shorts"], 200e6, places=6)
        self.assertAlmostEqual(fx["gold"], 35e6, places=6)
        self.assertAlmostEqual(fx["onop"], 335e6, places=6)
        self.assertAlmostEqual(fx["charge"], 26.8e6, places=6)

    def test_time_band_bounds_are_inclusive(self) -> None:
        self.assertEqual(ssa_reference.band_of(1.0, True), 4)
        self.assertEqual(ssa_reference.band_of(1.0001, True), 5)
        # The low-coupon ladder has shorter bands from one year on.
        self.assertEqual(ssa_reference.band_of(1.9, False), 5)
        self.assertEqual(ssa_reference.band_of(1.95, True), 5)
        self.assertEqual(ssa_reference.band_of(1.95, False), 6)
        self.assertEqual(ssa_reference.band_of(25.0, True), 13)
        self.assertEqual(ssa_reference.band_of(25.0, False), 15)

    def test_a_position_without_an_issue_id_nets_with_no_other(self) -> None:
        # An entered issue ID never equals the key of a position without one.
        self.assertNotEqual(ssa_reference.issue_key("R1", ""), ssa_reference.issue_key("X", "R1"))
        self.assertNotEqual(ssa_reference.issue_key("R1", ""),
                            ssa_reference.issue_key("X", "POS:R1"))
        self.assertEqual(ssa_reference.issue_key("X", "iss9"), ssa_reference.issue_key("Y", "ISS9"))
        row = ("R1", "T", "Bond", "EUR", "Long", 1e6, "Fixed", 0.05, datetime.date(2030, 1, 1),
               None, "OTHER", "BB+ to BB-", "", "")
        other = ("X1", *row[1:4], "Short", *row[5:12], "R1", "")
        with unittest.mock.patch.object(ssa_reference, "IR", [row, other]):
            # 8% of each position, not of a net of zero.
            self.assertAlmostEqual(ssa_reference.calc_ir()["EUR"]["sr"], 160_000.0, places=6)

    def test_the_published_example_under_the_crr(self) -> None:
        # The worked example C.2 of SRC-03: 1,000,000 of the zone 1 position is matched against zone 3,
        # so the CRR's 150% (Article 339(9)(f)) adds 50% of it to the published charge of 4,580,000.
        weighted = {2: [150_000.0, 0.0], 3: [0.0, 200_000.0], 4: [1_050_000.0, 0.0],
                    7: [1_125_000.0, 0.0], 10: [500_000.0, 5_625_000.0]}
        zone = {band: z for band, z, _, _, _ in ssa_reference.LADDER}
        basel = ssa_reference.ladder_charges(weighted, zone)
        crr = ssa_reference.ladder_charges(weighted, zone, 1.50)
        self.assertEqual((basel["between"], basel["gmr"]), (1_450_000.0, 4_580_000.0))
        self.assertEqual((crr["between"], crr["gmr"]), (1_950_000.0, 5_080_000.0))
        self.assertEqual((crr["vd"], crr["within"], crr["net"]),
                         (basel["vd"], basel["within"], basel["net"]))

    def test_the_index_charge_of_the_crr_is_nil(self) -> None:
        # Article 344: the specific risk of an index future that is exchange traded and diversified is
        # ignored, and its general risk stays in the net position of its market.
        basel, crr = ssa_reference.calc_eq(), ssa_reference.calc_eq(0.0)
        self.assertGreater(basel["index"], 0.0)
        self.assertEqual(crr["index"], 0.0)
        self.assertEqual((crr["specific"], crr["general"]), (basel["specific"], basel["general"]))
        self.assertAlmostEqual(
            ssa_reference.calc_all()["total"] - ssa_reference.calc_all(1.0, 0.0)["total"],
            ssa_reference.SCALING["EQ"] * basel["index"], places=6)

    def test_offsets_between_zones(self) -> None:
        self.assertEqual(ssa_reference.offset(100.0, -30.0), (30.0, 70.0, 0.0))
        self.assertEqual(ssa_reference.offset(-100.0, 30.0), (30.0, -70.0, 0.0))
        self.assertEqual(ssa_reference.offset(100.0, 30.0), (0.0, 100.0, 30.0))

    def test_main_prints_the_result_as_json(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(ssa_reference.main(), 0)
        self.assertIn('"total": 48267441.0', output.getvalue())


@unittest.skipIf(openpyxl is None and not os.environ.get("GITHUB_ACTIONS"),
                 "openpyxl is not installed: python -m pip install -r .github/requirements.txt")
class NeedsOpenpyxl(unittest.TestCase):
    """Base of the tests that build or read a workbook."""

    def test_openpyxl_is_available(self) -> None:
        self.assertIsNotNone(openpyxl, "install .github/requirements.txt before running the tests")


class Generator(NeedsOpenpyxl):
    """build_prototype.py: the template and the formula workbook."""

    folder: tempfile.TemporaryDirectory[str]
    template: Path
    formula: Path

    @classmethod
    def setUpClass(cls) -> None:
        # Each form is built once for the class: a build takes seconds.
        cls.folder = tempfile.TemporaryDirectory(prefix="prototype-")
        cls.addClassCleanup(cls.folder.cleanup)
        cls.template = Path(cls.folder.name) / "Template.xlsx"
        cls.formula = Path(cls.folder.name) / "Formula.xlsx"
        if openpyxl is not None:
            tool("build_prototype").build(cls.template, True)
            tool("build_prototype").build(cls.formula, False)

    def test_the_tracked_template_is_what_the_generator_writes(self) -> None:
        built = canonical_parts(self.template)
        tracked = canonical_parts(TEMPLATE)
        self.assertEqual(sorted(built), sorted(tracked))
        for name in tracked:
            with self.subTest(part=name):
                self.assertEqual(
                    built[name], tracked[name],
                    "write the template again: python tools/prototype/build_prototype.py "
                    "src/workbook/Template.xlsx --template")

    def test_the_tracked_matrix_is_what_the_generator_writes(self) -> None:
        generator = tool("build_prototype")
        written = generator.protection_matrix(generator.assemble(True))
        tracked = json.loads((ROOT / "docs" / "methodology" / "workbook-protection.json")
                             .read_text(encoding="utf-8"))
        self.assertEqual(written, tracked,
                         "write the matrix again: python tools/prototype/build_prototype.py "
                         "docs/methodology/workbook-protection.json --matrix")
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "matrix.json"
            generator.build_matrix(target)
            self.assertEqual(target.read_text(encoding="utf-8"),
                             (ROOT / "docs" / "methodology" / "workbook-protection.json")
                             .read_text(encoding="utf-8"))

    def test_every_sheet_is_protected_without_a_password_in_both_forms(self) -> None:
        for path in (self.template, self.formula):
            with zipfile.ZipFile(path) as package:
                sheets = [name for name in package.namelist()
                          if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)]
                self.assertEqual(len(sheets), 18)
                for name in sheets:
                    with self.subTest(path=path.name, sheet=name):
                        element = re.search(rb"<sheetProtection[^>]*>",
                                            package.read(name))
                        self.assertIsNotNone(element)
                        assert element is not None
                        attributes = dict(re.findall(rb'(\w+)="([^"]*)"', element.group(0)))
                        self.assertEqual(attributes[b"sheet"], b"1")
                        # Users may size columns and rows: 0 is allowed, the default blocks.
                        self.assertEqual(attributes[b"formatColumns"], b"0")
                        self.assertEqual(attributes[b"formatRows"], b"0")
                        for key in (b"password", b"hashValue", b"saltValue", b"algorithmName",
                                    b"spinCount"):
                            self.assertNotIn(key, attributes)

    def test_only_the_cells_with_the_input_fill_are_unlocked(self) -> None:
        assert openpyxl is not None
        generator = tool("build_prototype")
        for path in (self.template, self.formula):
            book = openpyxl.load_workbook(path)
            for sheet in book.worksheets:
                unlocked = {cell.coordinate for row in sheet.iter_rows() for cell in row
                            if cell.has_style and not cell.protection.locked}
                expected = {cell.coordinate for row in sheet.iter_rows() for cell in row
                            if generator.is_input(sheet.title, cell)}
                with self.subTest(path=path.name, sheet=sheet.title):
                    self.assertEqual(unlocked, expected)
        # The colours: yellow is an input everywhere, grey only on Params, where it is a
        # MAR40 constant that a what-if may change; elsewhere grey is an expected value.
        book = openpyxl.load_workbook(self.template)
        self.assertTrue(book["Params"]["E6"].protection.locked is False)
        self.assertTrue(book["Params"]["G38"].protection.locked is False)
        self.assertTrue(book["TestCatalogue"]["J7"].protection.locked)
        self.assertTrue(book["Results"]["C6"].protection.locked)
        self.assertTrue(book["IRPositions"]["C4"].protection.locked)
        self.assertFalse(book["IRPositions"]["C5"].protection.locked)

    def test_the_vba_applies_the_options_of_the_matrix_to_the_sheets_it_writes(self) -> None:
        matrix = json.loads((ROOT / "docs" / "methodology" / "workbook-protection.json")
                            .read_text(encoding="utf-8"))
        source = (ROOT / "src" / "workbook" / "M_Main.bas").read_bytes().decode(
            "cp1252").replace("\r\n", "\n")
        call = re.search(r"Late\.Protect ((?:[^\n]*_\n)*[^\n]*)", source)
        assert call is not None
        passed = dict(re.findall(r"(\w+):=(\w+)", call.group(1).replace("_\n", " ")))
        defaults = {"formatCells": True, "formatColumns": True, "formatRows": True,
                    "insertColumns": True, "insertRows": True, "insertHyperlinks": True,
                    "deleteColumns": True, "deleteRows": True, "sort": True,
                    "autoFilter": True, "pivotTables": True, "objects": False,
                    "scenarios": False}
        allow = {"formatCells": "AllowFormattingCells", "formatColumns": "AllowFormattingColumns",
                 "formatRows": "AllowFormattingRows", "insertColumns": "AllowInsertingColumns",
                 "insertRows": "AllowInsertingRows", "insertHyperlinks": "AllowInsertingHyperlinks",
                 "deleteColumns": "AllowDeletingColumns", "deleteRows": "AllowDeletingRows",
                 "sort": "AllowSorting", "autoFilter": "AllowFiltering",
                 "pivotTables": "AllowUsingPivotTables"}
        for sheet in ("shResults", "shIRZones", "shIRLadder", "shBuckets", "shIRCalc", "shEQCalc",
                      "shCOCalc", "shFXCalc", "shChecks"):
            options = matrix["sheets"][sheet]["protection"]
            with self.subTest(sheet=sheet):
                self.assertTrue(options["sheet"])
                self.assertEqual(passed["Contents"], "True")
                self.assertEqual(passed["UserInterfaceOnly"], "True")
                self.assertEqual(passed["DrawingObjects"] == "True",
                                 options.get("objects", defaults["objects"]))
                self.assertEqual(passed["Scenarios"] == "True",
                                 options.get("scenarios", defaults["scenarios"]))
                # An action the matrix does not block is passed as allowed, and only those.
                allowed = {name for name, blocked in defaults.items()
                           if name in allow and not options.get(name, blocked)}
                self.assertEqual({key for key, value in passed.items()
                                  if key.startswith("Allow") and value == "True"},
                                 {allow[name] for name in allowed})

    def test_the_sheets_that_code_writes_are_those_the_contract_names(self) -> None:
        contract = (ROOT / "docs" / "methodology" / "WORKBOOK_CONTRACT.md").read_text(
            encoding="utf-8")
        table = contract.split("## Sheets\n", 1)[1].split("\n\n- **Code name:**", 1)[0]
        rows = [[cell.strip() for cell in line.strip("|").split("|")]
                for line in table.splitlines() if re.match(r"\| +\d+ +\|", line)]
        self.assertEqual(len(rows), 18)
        written = {row[1].strip("`") for row in rows if row[4] == "Yes"}
        source = (ROOT / "src" / "workbook" / "M_Main.bas").read_bytes().decode(
            "cp1252").replace("\r\n", "\n")
        listed = re.search(r"SheetNames = Array\(([^)]*)\)", source.replace("_\n", " "))
        assert listed is not None
        constants = re.findall(r"SH_\w+", listed.group(1))
        config = (ROOT / "src" / "core" / "CORE_Config.bas").read_bytes().decode("cp1252")
        names = {re.search(rf'Const {name}\s+As String = "(\w+)"', config).group(1)  # type: ignore[union-attr]
                 for name in constants}
        self.assertEqual(names, written)

    def test_two_builds_are_byte_identical_and_carry_no_clock(self) -> None:
        generator = tool("build_prototype")
        for template, built in ((True, self.template), (False, self.formula)):
            with self.subTest(template=template), tempfile.TemporaryDirectory() as folder:
                again = Path(folder) / "again.xlsx"
                generator.build(again, template)
                first, second = built.read_bytes(), again.read_bytes()
                self.assertEqual(first, second)
                with zipfile.ZipFile(io.BytesIO(first)) as package:
                    for info in package.infolist():
                        self.assertEqual(info.date_time, generator.ZIP_TIME)
                    core = package.read("docProps/core.xml").decode()
                self.assertEqual(core.count(generator.PACKAGE_TIME), 2)

    def test_the_template_has_no_formula_in_an_output_area_and_the_formula_workbook_has(
            self) -> None:
        assert openpyxl is not None
        template = openpyxl.load_workbook(self.template)
        formula = openpyxl.load_workbook(self.formula)
        for sheet in ("IRCalc", "EQCalc", "COCalc", "FXCalc"):
            with self.subTest(sheet=sheet):
                self.assertIsNone(template[sheet]["D6"].value)
                self.assertTrue(str(formula[sheet]["D6"].value).startswith("="))
        self.assertEqual(template["Results"]["C4"].value, "Not calculated: run RunSSA.")
        self.assertEqual(template.sheetnames, formula.sheetnames)
        self.assertEqual(sorted(template.defined_names), sorted(formula.defined_names))

    def test_the_stamp_names_sit_on_the_readme_sheet(self) -> None:
        assert openpyxl is not None
        generator = tool("build_prototype")
        workbook = openpyxl.load_workbook(TEMPLATE)
        sheet = workbook["README"]
        for label, name in generator.STAMP:
            with self.subTest(name=name):
                target, _, cell = workbook.defined_names[name].attr_text.partition("!")
                self.assertEqual(target, "README")
                row = sheet[cell.replace("$", "")].row
                self.assertEqual(sheet.cell(row, 3).value, label)
                self.assertIsNone(sheet.cell(row, 4).value)

    def test_the_test_catalogue_expects_the_reference_values(self) -> None:
        assert openpyxl is not None
        sheet = openpyxl.load_workbook(TEMPLATE)["TestCatalogue"]
        expected = {sheet.cell(row, 3).value: sheet.cell(row, 10).value for row in range(7, 26)}
        result = ssa_reference.calc_all()
        self.assertAlmostEqual(expected["T01"], result["fx"]["onop"], places=6)
        self.assertAlmostEqual(expected["T02"], result["fx"]["charge"], places=6)
        self.assertAlmostEqual(expected["T16"], result["total"], places=6)
        self.assertAlmostEqual(expected["T17"], result["rwa"], places=6)

    def test_the_formula_workbook_follows_the_engines_rules(self) -> None:
        assert openpyxl is not None
        sheet = openpyxl.load_workbook(self.formula)["IRCalc"]
        status, issue_key = str(sheet["D5"].value), str(sheet["S5"].value)
        # A date that is a number but not a date serial rejects the row.
        self.assertIn('IF(OR(IRPositions!K5<1,IRPositions!K5>2958465),'
                      '"INVALID: Final maturity date is not a date"', status)
        self.assertIn('"INVALID: Next repricing date is not a date"', status)
        # Every row placeholder of the generator was replaced by the row number.
        self.assertNotIn("{r}", status)
        # A position without an Issue ID is an issue of its own.
        self.assertIn('IF(IFERROR(TRIM(IRPositions!O5),"#ERROR")="","POS:"&$C5,', issue_key)

    def test_duplicate_position_ids_are_counted_as_text_not_as_patterns(self) -> None:
        assert openpyxl is not None
        book = openpyxl.load_workbook(self.formula)
        for sheet in ("IRCalc", "EQCalc", "COCalc", "FXCalc"):
            with self.subTest(sheet=sheet):
                status = str(book[sheet]["D5"].value)
                self.assertIn('IF(SUMPRODUCT(--(UPPER($C$5:$C$504)=UPPER($C5)))>1,'
                              '"INVALID: duplicate PositionID"', status)
                # A criteria function would read "*" or "?" in a PositionID as a pattern.
                self.assertNotIn("COUNTIF($C$5:$C$504,$C5)", status)

    def test_rows_are_grouped_by_typed_keys_as_text_not_as_patterns(self) -> None:
        assert openpyxl is not None
        book = openpyxl.load_workbook(self.formula)
        # A criteria function may only test a constant, and MATCH may only look
        # up a row number: either would read a typed key as a pattern.
        criteria = re.compile(r'\b(COUNTIFS?|SUMIFS?|AVERAGEIFS?|MAXIFS|MINIFS)\((?![^,()]+,"[^"]*"[,)])')
        lookup = re.compile(r"\bMATCH\((?!\d+,|\$C\d+,(?:EQ|CO)Calc!)|\b[VHX]?LOOKUP\(")
        formulas = 0
        for sheet in book:
            for row in sheet.iter_rows():
                for cell in row:
                    value = cell.value
                    if isinstance(value, str) and value.startswith("="):
                        formulas += 1
                        self.assertNotRegex(value, criteria, f"{sheet.title}!{cell.coordinate}")
                        self.assertNotRegex(value, lookup, f"{sheet.title}!{cell.coordinate}")
        self.assertGreater(formulas, 10000)
        calc = book["IRCalc"]
        # The rows of an issue have the same netting key: the engine's own, so
        # that an entered "POS:R1" is not the position R1. The key reads the
        # Issue ID of a row only when the row has an Issue key.
        self.assertEqual(calc["Z5"].value,
                         '=IF($S5="","",IF(IFERROR(TRIM(IRPositions!O5),"#ERROR")="",$S5,'
                         '"ISSUE:"&$S5))')
        self.assertEqual(calc["T5"].value,
                         '=IF($S5="","",SUMPRODUCT((UPPER($Z$5:$Z$504)=UPPER($Z5))*1,$G$5:$G$504))')
        self.assertIn("SUMPRODUCT((UPPER($Z$5:$Z5)=UPPER($Z5))*1)=1", str(calc["U5"].value))
        # No formula compares a whole column of a position sheet: an error
        # value on one row would make an error of every row's result.
        for sheet in book:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        self.assertNotRegex(cell.value, r"Positions!\$?[A-Z]+\$?\d+:",
                                            f"{sheet.title}!{cell.coordinate}")
        # A typed text is read as the engine reads it.
        self.assertEqual(calc["C5"].value, '=IFERROR(TRIM(IRPositions!C5),"#ERROR")')
        # A direction and a rate type are read without the blanks at their ends.
        self.assertIn('IF(AND(IFERROR(TRIM(IRPositions!G5),"#ERROR")<>"Long",'
                      'IFERROR(TRIM(IRPositions!G5),"#ERROR")<>"Short"),', str(calc["D5"].value))
        self.assertIn('IFERROR(TRIM(IRPositions!I5),"#ERROR")="Floating"', str(calc["D5"].value))
        self.assertIsNone(openpyxl.load_workbook(self.template)["IRCalc"]["Z4"].value)
        # A blank currency is not one of FX_Table, whose unused rows are blank, and a
        # code of FX_Table is read trimmed and without regard to case.
        self.assertIn('IF(OR($E5="",SUMPRODUCT((UPPER(TRIM(FX_Ccy))=UPPER(TRIM($E5)))*1)=0),'
                      '"INVALID: currency not in FX_Table"', str(calc["D5"].value))
        self.assertEqual(book["IR_Ladder"]["C6"].value,
                         '=IF(OR(UPPER(TRIM(INDEX(FX_Ccy,1)))="",UPPER(TRIM(INDEX(FX_Ccy,1)))=GoldCode),'
                         '"",UPPER(TRIM(INDEX(FX_Ccy,1))))')
        self.assertEqual(book["IR_Ladder"]["I6"].value,
                         '=IF($C6="",0,SUMPRODUCT((UPPER(IRCalc!$E$5:$E$504)=UPPER($C6))'
                         '*(IRCalc!$K$5:$K$504=$D6)*(IRCalc!$G$5:$G$504>0)*1,IRCalc!$G$5:$G$504))')

    def test_the_engine_limit_is_the_number_of_rows_the_template_is_built_for(self) -> None:
        generator = tool("build_prototype")
        config = (ROOT / "src" / "core" / "CORE_Config.bas").read_text(encoding="cp1252")
        match = re.search(r"Public Const MAX_POSITION_ROWS As Long = (\d+)", config)
        assert match is not None
        self.assertEqual(int(match.group(1)), generator.CAP)
        self.assertEqual(int(match.group(1)), 500)

    def test_every_output_sheet_repeats_the_status_line(self) -> None:
        assert openpyxl is not None
        generator = tool("build_prototype")
        self.assertEqual(generator.STATUS_SHEETS, ("IR_Zones", "IR_Ladder", "Buckets", "IRCalc",
                                                   "EQCalc", "COCalc", "FXCalc", "Checks"))
        for form in (self.template, self.formula):
            book = openpyxl.load_workbook(form)
            for sheet in generator.STATUS_SHEETS:
                with self.subTest(form=form.name, sheet=sheet):
                    self.assertEqual(book[sheet]["E1"].value, "=Results_Status")
            self.assertEqual(book.defined_names["Results_Status"].attr_text, "Results!$C$4")
        template = openpyxl.load_workbook(self.template)
        self.assertEqual(template["Results"]["C4"].value, "Not calculated: run RunSSA.")
        # A status line that is not to be relied on stands out, on every sheet.
        for sheet, cell in (("Results", "C4"), ("IRCalc", "E1"), ("Checks", "E1")):
            rules = [rule.formula[0] for ranges in template[sheet].conditional_formatting
                     if str(ranges.sqref) == cell for rule in ranges.rules]
            with self.subTest(sheet=sheet):
                self.assertEqual(len(rules), 1)
                for word in ('"INCOMPLETE"', '"FAILED"', '"OUTDATED"'):
                    self.assertIn(word, rules[0])

    def test_the_fingerprint_of_the_inputs_has_a_named_cell_that_the_template_leaves_empty(
            self) -> None:
        assert openpyxl is not None
        book = openpyxl.load_workbook(self.template)
        self.assertEqual(book.defined_names["Results_Fingerprint"].attr_text, "Results!$E$18")
        self.assertIsNone(book["Results"]["E18"].value)
        self.assertEqual(book["Results"]["C18"].value, "Input fingerprint")
        # It is outside the areas that a run clears and the comparison reads.
        compare = tool("compare_outputs")
        self.assertFalse(any(area.sheet == "Results" and 18 in area.rows
                             for area in compare.FIXED_AREAS))

    def test_no_other_repository_is_named_in_the_workbook(self) -> None:
        assert openpyxl is not None
        for sheet in openpyxl.load_workbook(TEMPLATE):
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and "K-" in cell.value:
                        self.assertNotRegex(cell.value, r"K-(?!FRTB-SSA)[A-Z]",
                                            f"{sheet.title}!{cell.coordinate}")

    def test_main_writes_the_form_it_is_asked_for(self) -> None:
        generator = tool("build_prototype")
        with tempfile.TemporaryDirectory() as folder, \
                contextlib.redirect_stdout(io.StringIO()) as output:
            target = Path(folder) / "T.xlsx"
            self.assertEqual(generator.main([str(target), "--template"]), 0)
            self.assertEqual(target.read_bytes(), self.template.read_bytes())
        self.assertIn("saved", output.getvalue())


class Comparison(NeedsOpenpyxl):
    """compare_outputs.py on two small workbooks of values."""

    def workbook(self, folder: str, name: str, **cells: object) -> Path:
        """A workbook with the output sheets, one row of each and ``cells`` set."""
        assert openpyxl is not None
        book = openpyxl.Workbook()
        book.remove(book.active)
        for title in ("IRCalc", "EQCalc", "COCalc", "FXCalc", "IR_Ladder", "IR_Zones", "Buckets",
                      "Results", "Checks", "TestCatalogue"):
            book.create_sheet(title)
        book["IRCalc"]["C5"] = "IR-001"
        book["IRCalc"]["D5"] = "OK"
        book["IRCalc"]["G5"] = 25_000_000
        book["IRCalc"]["H5"] = datetime.datetime(2034, 3, 1)
        book["Results"]["K11"] = 48_267_441
        for reference, value in cells.items():
            sheet, _, cell = reference.partition("_")
            book[sheet][cell] = value
        path = Path(folder) / name
        book.save(path)
        return path

    def test_equal_workbooks_have_no_difference(self) -> None:
        compare = tool("compare_outputs")
        with tempfile.TemporaryDirectory() as folder:
            first = self.workbook(folder, "a.xlsx")
            # A date as its serial, a number within the tolerance, an empty text.
            serial = (datetime.datetime(2034, 3, 1) - compare.EPOCH).days
            second = self.workbook(folder, "b.xlsx", IRCalc_H5=serial,
                                   IRCalc_G5=25_000_000.000001, IRCalc_I5="")
            compared, differences = compare.compare_workbooks(first, second)
            self.assertEqual(differences, [])
            self.assertGreater(compared, 4000)
            with contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(compare.main([str(first), str(second)]), 0)
        self.assertIn("differences: 0", output.getvalue())

    def test_a_difference_names_its_cell_and_both_values(self) -> None:
        compare = tool("compare_outputs")
        with tempfile.TemporaryDirectory() as folder:
            first = self.workbook(folder, "a.xlsx")
            second = self.workbook(folder, "b.xlsx", Results_K11=48_267_500, IRCalc_D5="INVALID: x")
            _, differences = compare.compare_workbooks(first, second)
            self.assertEqual(differences, [
                "IRCalc!D5: VBA='OK' formula='INVALID: x'",
                "Results!K11: VBA=48267441 formula=48267500"])
            with contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(compare.main([str(first), str(second)]), 1)
        self.assertIn("differences: 2", output.getvalue())

    def test_a_row_only_one_workbook_uses_is_compared(self) -> None:
        compare = tool("compare_outputs")
        with tempfile.TemporaryDirectory() as folder:
            first = self.workbook(folder, "a.xlsx")
            # The formula workbook numbers an empty Buckets row; the engine leaves it blank.
            numbered = self.workbook(folder, "b.xlsx", Buckets_C8=2)
            self.assertEqual(compare.compare_workbooks(first, numbered)[1], [])
            extra = self.workbook(folder, "c.xlsx", IRCalc_C6="IR-002", IRCalc_D6="OK")
            _, differences = compare.compare_workbooks(first, extra)
        self.assertEqual(differences, ["IRCalc!C6: VBA=None formula='IR-002'",
                                       "IRCalc!D6: VBA=None formula='OK'"])

    def test_values_are_compared_by_kind(self) -> None:
        compare = tool("compare_outputs")
        self.assertTrue(compare.same(None, ""))
        self.assertTrue(compare.same(1, 1.0))
        self.assertTrue(compare.same(1e9, 1e9 + 100))
        self.assertFalse(compare.same(1e9, 1e9 + 10_000))
        # A logical value may come back from LibreOffice as 1 or 0.
        self.assertTrue(compare.same(True, 1))
        self.assertFalse(compare.same(True, False))
        self.assertFalse(compare.same("1", 1))
        self.assertTrue(compare.same(datetime.datetime(1900, 3, 1), 61))

    def test_a_missing_workbook_exits_2(self) -> None:
        compare = tool("compare_outputs")
        with contextlib.redirect_stderr(io.StringIO()) as errors:
            self.assertEqual(compare.main(["no-such-a.xlsx", "no-such-b.xlsx"]), 2)
        self.assertIn("no workbook at no-such-a.xlsx", errors.getvalue())


class Verification(NeedsOpenpyxl):
    """verify.py without LibreOffice: its scenarios and its control flow."""

    def test_the_negative_scenario_adds_its_rows_below_the_sample(self) -> None:
        assert openpyxl is not None
        verify = tool("verify")
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "negative.xlsx"
            verify.negative(TEMPLATE, target)
            with warnings.catch_warnings():
                # The number in a date cell, on purpose.
                warnings.simplefilter("ignore", UserWarning)
                book = openpyxl.load_workbook(target)
            with zipfile.ZipFile(target) as package:
                sheet_xml = package.read("xl/worksheets/sheet3.xml").decode()
        positions = book["IRPositions"]
        first = verify.FIRST_ROW + len(ssa_reference.IR)
        self.assertEqual(positions.cell(first - 1, 3).value, ssa_reference.IR[-1][0])
        added = [positions.cell(row, 3).value for row in range(first, first + 24)]
        self.assertEqual(added, ["IR-INST1", "IR-INST2", "IR-INST3", "IR-INST4", "IR-BAD1", "IR-BAD2", "ir-001", "IR-MIX", "IR-FRN", "IR-OTH",
                                 "IR-DATE", "R1", "IR-R1", "IR-*", "IR-W1", "IR-W2", "IR-W3",
                                 "IR-W4", "IR-P1", "IR-C1", "IR-C2", "IR-C3", "IR-E1", "IR-E2"])
        self.assertIsNone(positions.cell(first + 24, 3).value)
        # An instrument outside the list, none at all, and one typed in another case with blanks.
        self.assertEqual([positions.cell(first + offset, 5).value for offset in range(4)],
                         ["Securitisation", None, " swap LEG ", "Swap  leg"])
        # Error values, not texts that read like them: the Issue ID of a
        # rejected row, which has no coupon, and a direction.
        self.assertIsNone(positions.cell(first + 22, 10).value)
        for offset, column in ((22, 15), (23, 7)):
            with self.subTest(offset=offset, column=column):
                self.assertEqual(positions.cell(first + offset, column).data_type, "e")
        # Keys that a criteria function would read as patterns or comparisons.
        self.assertEqual([positions.cell(first + offset, 15).value for offset in range(14, 19)],
                         ["ISS-?", "ISS-A", "ISS-*", "<>ISS-A", "POS:R1"])
        codes = verify.first_cell(book, "FX_Ccy")
        self.assertEqual(codes.offset(len(ssa_reference.FX) - 1).value, ssa_reference.FX[-1][0])
        self.assertEqual(codes.offset(len(ssa_reference.FX)).value, verify.WILD_CURRENCY)
        self.assertEqual(verify.first_cell(book, "FX_Rate").offset(len(ssa_reference.FX)).value,
                         0.5)
        # The maturity that is not a date, and the Issue ID equal to a PositionID.
        # openpyxl reads a number above the date range in a date cell as an
        # error, so the number is looked for in the sheet itself.
        self.assertRegex(sheet_xml, rf'<c r="K{first + 10}"[^>]*><v>20311231</v>')
        self.assertEqual(positions.cell(first + 12, 15).value, "R1")
        self.assertIsNone(positions.cell(first + 11, 15).value)
        self.assertEqual(verify.named_cell(book, "ScalingIR").value, 1.5)
        self.assertEqual(verify.named_cell(book, "OptionsFX").value, 1000)
        for sheet, rows in (("EQPositions", ssa_reference.EQ), ("COPositions", ssa_reference.CO),
                            ("FXPositions", ssa_reference.FXPOS)):
            with self.subTest(sheet=sheet):
                self.assertIsNotNone(book[sheet].cell(verify.FIRST_ROW + len(rows), 3).value)

    def test_error_values_are_typed_where_an_output_shows_the_text(self) -> None:
        assert openpyxl is not None
        verify = tool("verify")
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "errors.xlsx"
            verify.error_keys(TEMPLATE, target)
            book = openpyxl.load_workbook(target)
        typed = {"IRPositions": (len(ssa_reference.IR), ((0, 15), (1, 15), (2, 6))),
                 "EQPositions": (len(ssa_reference.EQ), ((0, 3), (1, 7), (2, 5))),
                 "COPositions": (len(ssa_reference.CO), ((0, 5), (1, 10))),
                 "FXPositions": (len(ssa_reference.FXPOS), ((0, 5),))}
        for sheet, (rows, cells) in typed.items():
            for offset, column in cells:
                with self.subTest(sheet=sheet, offset=offset, column=column):
                    self.assertEqual(
                        book[sheet].cell(verify.FIRST_ROW + rows + offset, column).data_type, "e")
        self.assertEqual(verify.ERROR_TEXT, "#ERROR")

    def test_error_values_in_an_output_area_are_listed(self) -> None:
        assert openpyxl is not None
        verify = tool("verify")
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "values.xlsx"
            book = openpyxl.load_workbook(TEMPLATE)
            self.assertEqual(verify.error_cells(TEMPLATE), [])
            book["Results"]["E7"].value = "#N/A"
            # Outside every output area, and a text that only reads like an error.
            book["Results"]["A1"].value = "#N/A"
            book["Results"]["E8"].value = "#ERROR"
            book.save(path)
            self.assertEqual(verify.error_cells(path), ["Results!E7"])

    def test_the_unusable_scenario_changes_two_named_cells(self) -> None:
        assert openpyxl is not None
        verify = tool("verify")
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "unusable.xlsx"
            verify.unusable(TEMPLATE, target)
            book = openpyxl.load_workbook(target)
        self.assertEqual(verify.named_cell(book, "AsOfDate").value, "not a date")
        self.assertEqual(verify.named_cell(book, "OptionsIR").value, "abc")

    def test_findings_are_read_below_their_header(self) -> None:
        assert openpyxl is not None
        verify = tool("verify")
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "Checks"
        sheet["C24"] = "#"
        for row, values in ((25, (1, "IRPositions", "IR-012", 16, "INVALID: x")),
                            (26, (2, "Params", "AsOfDate", None, "must be a date"))):
            for column, value in zip((3, 4, 5, 6, 9), values):
                sheet.cell(row, column).value = value
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "checks.xlsx"
            book.save(path)
            self.assertEqual(verify.findings(path), {
                ("IRPositions", "IR-012", 16, "INVALID: x"),
                ("Params", "AsOfDate", None, "must be a date")})

    def test_a_macro_result_is_the_last_result_line(self) -> None:
        verify = tool("verify")
        run = verify.Run(Path("folder"), "python-with-uno")
        done = subprocess.CompletedProcess([], 0, "noise\nRESULT: RESULT=OK; errors=0\n", "")
        with unittest.mock.patch("subprocess.run", return_value=done) as called:
            self.assertEqual(run.macro(Path("w.xlsm"), verify.RUN, Path("o.xlsx")),
                             "RESULT=OK; errors=0")
        self.assertEqual(called.call_args.args[0],
                         ["python-with-uno", str(verify.LO_MACRO), "w.xlsm", verify.RUN, "o.xlsx"])
        failed = subprocess.CompletedProcess([], 1, "lo-macro: FAIL (cannot load)\n", "")
        with unittest.mock.patch("subprocess.run", return_value=failed), \
                self.assertRaisesRegex(verify.Failure, "cannot load"):
            run.macro(Path("w.xlsm"), verify.RUN, Path("o.xlsx"))

    def test_the_suites_step_runs_both_suites_and_needs_each_to_pass(self) -> None:
        verify = tool("verify")
        run = verify.Run(Path("folder"), "python-with-uno")
        results = {"TEST_Suite.Test_RunCore": "RESULT=PASS; suite=core; assertions=90",
                   "TEST_Workbook.Test_RunWorkbook": "RESULT=PASS; suite=workbook; assertions=34",
                   "TEST_Cases.Test_RunCases": "RESULT=PASS; suite=cases; assertions=329"}
        self.assertEqual(verify.TESTS, tuple(results))
        with unittest.mock.patch.object(
                verify.Run, "macro", side_effect=lambda _book, macro, _out: results[macro]) as called:
            self.assertEqual(verify.test_suite(run), "\n".join(results.values()))
        self.assertEqual([call.args[1] for call in called.call_args_list], list(results))
        # Each run saves its own values, so one does not replace the other's.
        self.assertEqual(len({call.args[2] for call in called.call_args_list}), 3)
        for result in ("RESULT=FAIL; suite=workbook; failures=1", ""):
            results["TEST_Workbook.Test_RunWorkbook"] = result
            with unittest.mock.patch.object(
                    verify.Run, "macro", side_effect=lambda _book, macro, _out: results[macro]), \
                    self.assertRaisesRegex(verify.Failure, "TEST_Workbook.Test_RunWorkbook: "
                                           + (result or "no result")):
                verify.test_suite(run)

    def test_verify_stops_at_the_first_step_that_fails(self) -> None:
        verify = tool("verify")
        calls: list[str] = []

        def passing(run: Any) -> str:
            calls.append("pass")
            return "done"

        def failing(run: Any) -> str:
            calls.append("fail")
            raise verify.Failure("3 difference(s)")

        run = verify.Run(Path("folder"), sys.executable)
        with unittest.mock.patch.object(
                verify, "STEPS", (("First", passing), ("Second", failing), ("Third", passing))), \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertFalse(verify.verify(run))
        self.assertEqual(calls, ["pass", "fail"])
        self.assertIn("3 difference(s)", output.getvalue())
        self.assertIn("VERIFY: FAIL (step 2: Second)", output.getvalue())
        with unittest.mock.patch.object(verify, "STEPS", (("Only", passing),)), \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertTrue(verify.verify(run))
        self.assertIn("VERIFY: PASS", output.getvalue())

    def test_without_libreoffice_main_exits_2(self) -> None:
        verify = tool("verify")
        with unittest.mock.patch("shutil.which", return_value=None), \
                contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(verify.main([]), 2)
        self.assertIn("verify: SKIP (needs LibreOffice", output.getvalue())

    def test_the_formula_tests_must_all_pass(self) -> None:
        verify = tool("verify")
        run = verify.Run(Path("folder"), sys.executable)
        with unittest.mock.patch.object(verify.Run, "macro", return_value=""):
            with unittest.mock.patch.object(verify, "test_summary",
                                            return_value="19 of 19 tests PASS"):
                self.assertEqual(verify.formula_tests(run), "19 of 19 tests PASS")
            for summary in ("8 of 19 tests PASS", ""):
                with self.subTest(summary=summary), \
                        unittest.mock.patch.object(verify, "test_summary", return_value=summary), \
                        self.assertRaises(verify.Failure):
                    verify.formula_tests(run)


if __name__ == "__main__":
    unittest.main()
