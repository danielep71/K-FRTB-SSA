"""Static consistency of the VBA harness and its example suite (tests/modules/).

The harness runs only in Excel. These tests check, without Excel, what its
source can show: that the example suite's expected counts equal the cases and
assertions it contains, that every case runs once with a well-formed unique
identifier, that the documented passing RESULT line matches the counts, and
that every record the harness writes is documented. The relation suite is
checked the same way, and each of its cases must pair the relation with its
corruption check; its tests are skipped in a repository that deleted it. Each
project suite of the evidence policy is checked against the policy and the
shape the evidence check reads: its counts, its identifiers, where its
assertions are, and its certification entry point.
"""

from __future__ import annotations

import io
import json
import re
import tempfile
import unittest
from pathlib import Path
from typing import ClassVar

ROOT = Path(__file__).resolve().parents[1]
EXAMPLES = ROOT / "tests/modules/TEST_Examples.bas"
RELATIONS = ROOT / "tests/modules/TEST_Relations.bas"
HARNESS = ROOT / "tests/modules/TEST_Harness.bas"
GUIDE = ROOT / "docs/methodology/TEST_HARNESS.md"
REGISTRY = ROOT / "tests/accuracy/contracts.json"
BUILD_CONFIG = ROOT / "tools/devbuild/workbook.json"
POLICY = ROOT / ".github/excel-evidence-policy.json"
CASE_ID = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
# A RecordResult statement, not a comment that names it.
RECORD = re.compile(r"^\s*TEST_Harness\.RecordResult ", re.M)


def source(path: Path) -> str:
    return path.read_text(encoding="ascii").replace("\r\n", "\n")


def constant(text: str, name: str) -> int:
    match = re.search(rf"^\s*Private Const {name}\s+As Long = (\d+)\b", text, re.M)
    assert match, name
    return int(match.group(1))


def procedures(text: str) -> dict[str, str]:
    """Body of every Sub and Function, by name."""
    pattern = re.compile(r"^(?:Public|Private) (?:Sub|Function) (\w+)\(.*?^End (?:Sub|Function)$",
                         re.M | re.S)
    return {match.group(1): match.group(0) for match in pattern.finditer(text)}


class ExampleSuiteTest(unittest.TestCase):
    text = source(EXAMPLES)
    bodies = procedures(text)
    cases: ClassVar[dict[str, str]] = {name: body for name, body in bodies.items() if name.startswith("Case")}

    def test_runner_calls_every_case_once_in_order(self) -> None:
        runner = self.bodies["RunSuite"]
        called = re.findall(r"^\s+(Case\w+)$", runner, re.M)
        self.assertEqual(called, list(self.cases))

    def test_each_case_begins_once_with_a_unique_identifier(self) -> None:
        identifiers = []
        for name, body in self.cases.items():
            begun = re.findall(r'TEST_Harness\.BeginCase "([^"]+)"', body)
            self.assertEqual(len(begun), 1, name)
            self.assertRegex(begun[0], CASE_ID)
            self.assertIn("TEST_Harness.RecordCaseError", body, name)
            identifiers += begun
        self.assertEqual(len(identifiers), len(set(identifiers)))

    def test_expected_counts_match_the_source(self) -> None:
        assertions = sum(len(re.findall(r"TEST_Harness\.Assert\w+ ", body))
                         for body in self.cases.values())
        self.assertEqual(constant(self.text, "EXPECTED_CASES"), len(self.cases))
        self.assertEqual(constant(self.text, "EXPECTED_ASSERTIONS"), assertions)

    def test_documented_assertion_counts_per_case(self) -> None:
        for name, body in self.cases.items():
            stated = re.search(r"(\d+) assertions?\.", body)
            self.assertIsNotNone(stated, name)
            assert stated is not None
            self.assertEqual(int(stated.group(1)),
                             len(re.findall(r"TEST_Harness\.Assert\w+ ", body)), name)

    def test_documented_passing_line(self) -> None:
        cases = constant(self.text, "EXPECTED_CASES")
        assertions = constant(self.text, "EXPECTED_ASSERTIONS")
        line = ("RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE; "
                f"cases={cases}; assertions={assertions}; failures=0; cleanup=PASS; "
                "recovery=NONE")
        self.assertIn(line, (ROOT / "INSTALLATION.md").read_text(encoding="utf-8"))
        skipped = self.cases["CaseCleanupOnlyFailure"]
        remaining = assertions - len(re.findall(r"TEST_Harness\.Assert\w+ ", skipped))
        self.assertIn(f"`cases={cases - 1}`, `assertions={remaining}`",
                      (ROOT / "INSTALLATION.md").read_text(encoding="utf-8"))


    def results_cases(self) -> dict[str, str]:
        """The cases that record results, found by their RecordResult calls."""
        return {name: body for name, body in self.cases.items() if RECORD.search(body)}

    def recorded(self) -> list[tuple[str, str | None, str | None, str | None, str | None]]:
        """Case, variant, quantity, dataset and record of every RecordResult call."""
        strings = dict(re.findall(r'^\s*Private Const (\w+)\s+As String = "([^"]*)"', self.text,
                                  re.M))

        def value(token: str) -> str | None:
            if token == "vbNullString":
                return None
            return token.strip('"') if token.startswith('"') else strings[token]
        separator = r",\s*(?:_\s+)?"
        calls = "".join(self.results_cases().values())
        recorded = [(case, value(variant), value(quantity), value(dataset), value(record))
                    for case, variant, quantity, dataset, record in re.findall(
                        r'TEST_Harness\.RecordResult "([^"]+)"' + separator
                        + r'(vbNullString|"[^"]+")' + separator + r'(\w+|"[^"]+")' + separator
                        + r'(vbNullString|"[^"]*")' + separator + r'(vbNullString|"[^"]*")'
                        + separator + r"\w+", calls)]
        # Every call is read: none is skipped for a layout the pattern does not know.
        self.assertEqual(len(recorded), len(RECORD.findall(calls)), recorded)
        return recorded

    def test_the_results_case_records_every_output_of_the_contracts(self) -> None:
        recorded = self.recorded()
        registry = json.loads(REGISTRY.read_text(encoding="utf-8")) if REGISTRY.is_file() \
            else {"contracts": []}
        expected = []
        for contract in registry["contracts"]:
            for case in [*contract["calibration_cases"], *contract["holdout_cases"]]:
                data = json.loads((ROOT / "tests/expected" / f"{case}.json")
                                  .read_text(encoding="utf-8"))
                groups = [(None, data["outputs"])] if "outputs" in data \
                    else sorted(data["variants"].items())
                expected += [(case, variant, output["quantity"],
                              *((output["scope"]["dataset"], output["scope"]["record"])
                                if output["scope"] else (None, None)))
                             for variant, outputs in groups for output in outputs
                             if output["quantity"] in contract["quantities"]]
        self.assertEqual(sorted(recorded, key=str), sorted(expected, key=str))
        self.assertEqual(len(recorded), len(set(recorded)))

    def test_the_results_case_reads_files_that_exist(self) -> None:
        paths = [path for body in self.results_cases().values()
                 for path in re.findall(r'FixturePath\(\s*(?:_\s*)?"([^"]+)"\)', body)]
        for path in paths:
            with self.subTest(path=path):
                self.assertTrue((ROOT / "tests/fixtures" / path).is_file())
        # Each recorded case is computed from its own fixture file or folder.
        for case in {item[0] for item in self.recorded()}:
            with self.subTest(case=case):
                self.assertTrue(any(path == f"{case}.json" or path.startswith(f"{case}/")
                                    for path in paths), paths)

    def test_the_certification_saves_the_results_of_the_normal_run(self) -> None:
        body = self.bodies["CertifyExamples"]
        self.assertRegex(body, r'CertifyMode "NORMAL"[^\n]*\n\s+TEST_Harness\.SaveResults ')


class RelationSuiteTest(unittest.TestCase):
    """The relation suite, read when its tests run: a generated repository may
    delete the module, and its tests are then skipped."""

    path: ClassVar[Path] = RELATIONS
    text: ClassVar[str]
    bodies: ClassVar[dict[str, str]]
    cases: ClassVar[dict[str, str]]

    @classmethod
    def setUpClass(cls) -> None:
        if not cls.path.is_file():
            raise unittest.SkipTest(f"{cls.path.name} is not in this repository")
        cls.text = source(cls.path)
        cls.bodies = procedures(cls.text)
        cls.cases = {name: body for name, body in cls.bodies.items() if name.startswith("Case")}

    @staticmethod
    def assertions(body: str) -> list[str]:
        return re.findall(r'TEST_Harness\.Assert\w+ "([^"]+)"', body)

    def test_runner_calls_every_case_once_in_order(self) -> None:
        called = re.findall(r"^\s+(Case\w+)$", self.bodies["RunRelationTests"], re.M)
        self.assertEqual(called, list(self.cases))
        self.assertTrue(self.cases)

    def test_each_case_begins_once_with_a_unique_identifier(self) -> None:
        identifiers = []
        for name, body in self.cases.items():
            begun = re.findall(r'TEST_Harness\.BeginCase "([^"]+)"', body)
            self.assertEqual(len(begun), 1, name)
            self.assertRegex(begun[0], CASE_ID)
            self.assertIn("TEST_Harness.RecordCaseError", body, name)
            identifiers += begun
        self.assertEqual(len(identifiers), len(set(identifiers)))

    def test_every_relation_has_its_corruption_check(self) -> None:
        for name, body in self.cases.items():
            with self.subTest(case=name):
                self.assertEqual(sorted(self.assertions(body)), ["detects-corruption", "holds"])
                self.assertRegex(body, r'AssertTrue "detects-corruption", _\s+Not ')

    def test_expected_counts_match_the_source(self) -> None:
        assertions = sum(len(self.assertions(body)) for body in self.cases.values())
        cases = constant(self.text, "EXPECTED_CASES")
        self.assertEqual(cases, len(self.cases))
        self.assertEqual(constant(self.text, "EXPECTED_ASSERTIONS"), assertions)
        for name, body in self.cases.items():
            stated = re.search(r"(\d+) assertions?\.", body)
            self.assertIsNotNone(stated, name)
            assert stated is not None
            self.assertEqual(int(stated.group(1)), len(self.assertions(body)), name)
        passing = re.sub(r"\n'\s+", " ", self.bodies["RunRelationTests"])
        self.assertIn(f"cases={cases}; assertions={assertions}; failures=0", passing)


def code(text: str) -> str:
    """The statements of a module: no comment line, and no string literal."""
    lines = [line for line in text.split("\n") if not line.lstrip().startswith("'")]
    return re.sub(r'"[^"]*"', '""', "\n".join(lines))


class ProjectSuitesTest(unittest.TestCase):
    """Each suite the evidence policy lists under suites, against the policy."""

    suites: ClassVar[list[dict[str, object]]] = json.loads(
        POLICY.read_text(encoding="utf-8"))["suites"]

    def each(self) -> list[tuple[dict[str, object], str, dict[str, str], list[str]]]:
        """The entry, the source, the procedures and the cases the runner calls."""
        found = []
        for entry in self.suites:
            text = source(ROOT / str(entry["source"]))
            bodies = procedures(text)
            runner = bodies[str(entry.get("runner", "RunSuite"))]
            found.append((entry, text, bodies, re.findall(r"^\s+(Case\w+)$", runner, re.M)))
        return found

    def test_the_policy_lists_the_core_and_the_workbook_suite(self) -> None:
        self.assertEqual([(entry["name"], entry["source"]) for entry in self.suites],
                         [("core", "tests/modules/TEST_Suite.bas"),
                          ("workbook", "tests/modules/TEST_Workbook.bas")])

    def test_the_runner_calls_every_case_once(self) -> None:
        for entry, _, bodies, called in self.each():
            with self.subTest(suite=entry["name"]):
                self.assertEqual(sorted(called),
                                 sorted(name for name in bodies if name.startswith("Case")))
                self.assertEqual(len(called), len(set(called)))

    def test_each_case_begins_once_with_a_unique_identifier_and_records_its_error(self) -> None:
        for entry, _, bodies, called in self.each():
            identifiers = []
            for name in called:
                with self.subTest(suite=entry["name"], case=name):
                    begun = re.findall(r'TEST_Harness\.BeginCase "([^"]+)"', bodies[name])
                    self.assertEqual(len(begun), 1)
                    self.assertRegex(begun[0], CASE_ID)
                    # An error of the case is recorded, never left to end the run.
                    self.assertRegex(bodies[name], r"\n\s+On Error GoTo CaseFailed\n\s+"
                                     r"TEST_Harness\.BeginCase ")
                    self.assertRegex(bodies[name], r"\nCaseFailed:\n\s+"
                                     r"TEST_Harness\.RecordCaseError\n\nEnd Sub$")
                    identifiers += begun
            self.assertEqual(len(identifiers), len(set(identifiers)), entry["name"])

    def test_every_assertion_is_a_named_one_in_the_body_of_a_case(self) -> None:
        # The evidence check reads the assertions of a case in its own body, by their
        # literal names: one in a helper, or with a computed name, would not be counted.
        for entry, text, bodies, called in self.each():
            with self.subTest(suite=entry["name"]):
                named = []
                for name in called:
                    found = re.findall(r'TEST_Harness\.Assert\w+ "([^"]+)"', bodies[name])
                    self.assertEqual(len(found), len(set(found)), name)
                    for item in found:
                        self.assertRegex(item, CASE_ID, name)
                    named += found
                self.assertEqual(len(named), len(re.findall(r"TEST_Harness\.Assert\w+", code(text))))

    def test_the_counts_are_those_of_the_source_and_of_the_policy(self) -> None:
        for entry, text, bodies, called in self.each():
            with self.subTest(suite=entry["name"]):
                assertions = sum(len(re.findall(r"TEST_Harness\.Assert\w+ ", bodies[name]))
                                 for name in called)
                self.assertEqual(constant(text, "EXPECTED_CASES"), len(called))
                self.assertEqual(constant(text, "EXPECTED_ASSERTIONS"), assertions)
                self.assertEqual(entry["expected_cases"], len(called))
                self.assertEqual(entry["expected_assertions"], assertions)
                self.assertIn(f'Private Const SUITE_NAME          As String = "{entry["name"]}"',
                              text)

    def test_the_certification_entry_point_saves_the_logs_the_policy_names(self) -> None:
        for entry, _, bodies, _ in self.each():
            with self.subTest(suite=entry["name"]):
                body = bodies[str(entry["certification_entry_point"])]
                self.assertRegex(body, r"^Public Sub \w+\( _\n    ByVal Folder As String, _\n"
                                 r"    ByVal Fixtures As String\)\n")
                steps = ["TEST_Harness.BeginCertification SUITE_NAME",
                         "On Error Resume Next", "RunSuite", "On Error GoTo 0",
                         f'TEST_Harness.SaveLog Folder & "\\{entry["log"]}"',
                         'TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, '
                         "RunErrorDescription",
                         f'TEST_Harness.EndCertification Folder & "\\{entry["certification_log"]}"']
                statements = [line.strip() for line in body.split("\n")]
                positions = [statements.index(step) for step in steps]
                self.assertEqual(positions, sorted(positions))
                # A suite that names no results file saves none.
                self.assertEqual("results" in entry, "TEST_Harness.SaveResults" in body)

    def test_each_suite_has_an_entry_point_that_raises_and_one_that_returns_the_result(
            self) -> None:
        for entry, text, bodies, _ in self.each():
            with self.subTest(suite=entry["name"]):
                plain = [name for name, body in bodies.items()
                         if re.match(r"Public Sub \w+\(\)\n", body)
                         and re.search(r"^\s+RunSuite$", body, re.M)]
                self.assertEqual(len(plain), 1)
                returning = [name for name, body in bodies.items()
                             if re.match(r"Public Function \w+\(\) As String\n", body)
                             and f"{name} = TEST_Harness.LastResult()" in body]
                self.assertEqual(len(returning), 1)
                self.assertIn("Option Private Module", text)

    def test_the_core_suite_touches_no_workbook(self) -> None:
        statements = code(source(ROOT / "tests/modules/TEST_Suite.bas"))
        for word in ("ThisWorkbook", "ActiveWorkbook", "Worksheets", "Sheets", "Range", "Cells",
                     "Names", "Application", "RunSSA", "RunSSA_Silent", "CheckStatus_Silent",
                     "ClearOutputs_Silent"):
            with self.subTest(word=word):
                self.assertNotRegex(statements, rf"\b{word}\b")

    def test_the_workbook_suite_ends_with_the_output_sheets_cleared(self) -> None:
        text = source(ROOT / "tests/modules/TEST_Workbook.bas")
        bodies = procedures(text)
        last = re.findall(r"^\s+(Case\w+)$", bodies["RunSuite"], re.M)[-1]
        statements = [line.strip() for line in code(bodies[last]).split("\n")]
        cleared = max(index for index, line in enumerate(statements)
                      if line == "ClearOutputs_Silent")
        # Nothing after it runs the calculation again.
        self.assertFalse([line for line in statements[cleared:] if "RunSSA" in line])
        for name, body in bodies.items():
            if name.startswith("Case"):
                # What a case changes it puts back: as many restorations as changes.
                self.assertEqual(len(re.findall(r"^\s+Saved(?:Range)? = ", body, re.M)),
                                 len(re.findall(r" = Saved(?:Range)?$", body, re.M)), name)

    def test_the_guides_name_each_suite_and_its_entry_points(self) -> None:
        guide = GUIDE.read_text(encoding="utf-8")
        for entry, _, bodies, _ in self.each():
            module = Path(str(entry["source"])).stem
            with self.subTest(suite=entry["name"]):
                self.assertIn(f"| `{entry['name']}` | `{module}.bas` |", guide)
                self.assertIn(f"`{module}.{entry['certification_entry_point']}`", guide)
                for name, body in bodies.items():
                    if body.startswith("Public "):
                        self.assertIn(f"`{module}.{name}`", guide)


class MissingRelationSuiteTest(unittest.TestCase):
    def test_a_deleted_relation_suite_skips_its_tests(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            missing = type("Missing", (RelationSuiteTest,),
                           {"path": Path(directory) / "TEST_Relations.bas"})
            suite = unittest.TestLoader().loadTestsFromTestCase(missing)
            result = unittest.TextTestRunner(stream=io.StringIO(), verbosity=0).run(suite)
        self.assertEqual((result.testsRun, len(result.errors), len(result.failures)),
                         (0, 0, 0), result.errors)
        self.assertEqual(len(result.skipped), 1)  # the class, skipped once in setUpClass
        self.assertGreater(suite.countTestCases(), 0)


class HarnessTest(unittest.TestCase):
    text = source(HARNESS)

    def test_the_stamp_name_is_the_one_the_build_writes(self) -> None:
        match = re.search(r'Const STAMP_COMMIT_NAME\s+As String = "([^"]+)"', self.text)
        self.assertIsNotNone(match)
        assert match is not None
        layout = json.loads(BUILD_CONFIG.read_text(encoding="utf-8"))
        self.assertEqual(match.group(1), layout["stamp"]["commit"])

    def test_a_result_scope_is_the_object_of_the_expected_format(self) -> None:
        # The accuracy check matches a scope only as {"dataset": ..., "record": ...}.
        bodies = procedures(self.text)
        self.assertIn('""scope"": " & ScopeJson(Dataset, Record)', bodies["RecordResult"])
        scope = re.sub(r" _\n\s+", " ", bodies["ScopeJson"])  # join continued lines
        self.assertIn('"{""dataset"": " & JsonText(Dataset) & ", ""record"": " & '
                      'JsonText(Record) & "}"', scope)
        self.assertIn('ScopeJson = "null"', scope)

    def test_results_are_cleared_with_the_log(self) -> None:
        bodies = procedures(self.text)
        for name in ("BeginRun", "ResetTests"):
            with self.subTest(procedure=name):
                self.assertIn("Set mResults = New Collection", bodies[name])
        self.assertIn('RequireActiveRun "RecordResult"', bodies["RecordResult"])

    def test_every_documented_record_is_written(self) -> None:
        guide = GUIDE.read_text(encoding="utf-8")
        cells = [line.split("|")[1] for line in guide.splitlines() if line.startswith("| `")]
        documented = {key for cell in cells for key in re.findall(r"`([A-Z_]+)=", cell)}
        written = set(re.findall(r'"([A-Z_]+)=', self.text))
        self.assertEqual(documented, written - {"RESET"})
        self.assertIn("RESET=DONE", guide)

    def test_a_refused_start_writes_its_result_last(self) -> None:
        # The guide: RESULT is always the last record of a run, after RECOVERY.
        body = procedures(self.text)["BeginRun"]
        refused = body[body.index("If mRecoveryRequired Then"):body.index("Err.Raise ERR_TEST_DIRTY_START")]
        self.assertLess(refused.index('Say "RECOVERY=REQUIRED'), refused.index("Say mLastResult"))
        self.assertEqual(refused.count("Say "), 2)

    def test_certification_raises_only_after_writing(self) -> None:
        body = procedures(self.text)["EndCertification"]
        self.assertLess(body.index("WriteRecords Path, mCertLog"),
                        body.index("Err.Raise ERR_TEST_FAILED"))
        self.assertLess(body.index("mCertActive = False"), body.index("WriteRecords"))

    def test_error_values_are_not_compared_with_operators(self) -> None:
        # Comparing an error Variant with "=" or "<>" raises Type mismatch.
        self.assertNotRegex(self.text, r"\bActual\s*(?:=|<>)\s*CVErr\(")
        self.assertIn("CStr(Actual) = CStr(CVErr(ExpectedCode))", self.text)

    def test_documented_failure_ids_are_written(self) -> None:
        for failure in ("harness.case-id", "harness.no-case", "harness.completeness",
                        "runner.unexpected-error", ".unexpected-error"):
            with self.subTest(failure=failure):
                self.assertIn(f'"{failure}"', self.text)


if __name__ == "__main__":
    unittest.main()
