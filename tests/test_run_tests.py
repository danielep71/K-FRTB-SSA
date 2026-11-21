"""Tests for the parallel test runner (tools/run_tests.py)."""

from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import textwrap
import unittest
import unittest.mock
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import run_tests  # noqa: E402

PASSING = textwrap.dedent("""\
    import unittest


    class Passing(unittest.TestCase):
        def test_one(self):
            self.assertTrue(True)

        def test_two(self):
            self.assertEqual(1 + 1, 2)

        @unittest.skip("skipped on purpose")
        def test_skipped(self):
            pass
    """)
FAILING = textwrap.dedent("""\
    import unittest


    class Failing(unittest.TestCase):
        def test_fails(self):
            self.assertEqual(1, 2)
    """)
EMPTY = "import unittest\n"


class RunnerTest(unittest.TestCase):
    def folder(self, modules: dict[str, str]) -> Path:
        """A repository root whose tests/ holds ``modules``, set as the runner's."""
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        root = Path(temporary.name)
        (root / "tests").mkdir()
        for name, text in modules.items():
            (root / "tests" / name).write_text(text, encoding="utf-8")
        for name, value in (("ROOT", root), ("TESTS", root / "tests")):
            patcher = unittest.mock.patch.object(run_tests, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        return root

    def main(self, *argv: str) -> tuple[int, str]:
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            code = run_tests.main(list(argv))
        return code, output.getvalue()

    def test_every_module_passing_passes(self) -> None:
        self.folder({"test_a.py": PASSING, "test_b.py": PASSING})
        code, output = self.main("--jobs", "2")
        self.assertEqual(code, 0)
        self.assertIn("test_a.py: 3 tests, 1 skipped,", output)
        self.assertIn("test_b.py: 3 tests, 1 skipped,", output)
        self.assertIn("Ran 6 tests in 2 modules in ", output)
        self.assertIn("(2 skipped, 2 at a time)", output)
        self.assertTrue(output.rstrip().endswith("OK"))

    def test_a_failing_module_fails_the_run(self) -> None:
        # The negative control: one failing test among passing modules.
        self.folder({"test_a.py": PASSING, "test_b.py": FAILING, "test_c.py": PASSING})
        code, output = self.main()
        self.assertEqual(code, 1)
        self.assertIn("test_b.py: 1 tests, 0 skipped,", output)
        self.assertIn("FAILED (exit code 1)", output)
        self.assertIn("FAILED: test_b.py", output)
        self.assertIn("AssertionError: 1 != 2", output)

    def test_a_module_that_runs_no_test_fails_the_run(self) -> None:
        self.folder({"test_a.py": PASSING, "test_empty.py": EMPTY})
        code, output = self.main()
        self.assertEqual(code, run_tests.NO_TESTS)
        self.assertIn("FAILED: test_empty.py", output)

    def test_a_module_that_does_not_import_fails_the_run(self) -> None:
        self.folder({"test_a.py": "import missing_module_for_the_runner_test\n"})
        code, output = self.main()
        self.assertEqual(code, 1)
        self.assertIn("missing_module_for_the_runner_test", output)

    def test_named_modules_run_alone(self) -> None:
        self.folder({"test_a.py": PASSING, "test_b.py": FAILING})
        code, output = self.main("test_a.py")
        self.assertEqual(code, 0)
        self.assertNotIn("test_b.py", output)
        with self.assertRaises(SystemExit) as raised:
            self.main("test_missing.py")
        self.assertIn("test_missing.py: not a module of tests/", str(raised.exception))

    def test_the_largest_module_starts_first(self) -> None:
        self.folder({"test_a.py": PASSING, "test_b.py": FAILING, "test_c.py": PASSING + "\n" * 9})
        self.assertEqual(run_tests.order(["test_a.py", "test_b.py", "test_c.py"]),
                         ["test_c.py", "test_a.py", "test_b.py"])

    def test_jobs_must_be_positive(self) -> None:
        self.folder({"test_a.py": PASSING})
        with self.assertRaises(SystemExit), contextlib.redirect_stderr(io.StringIO()):
            self.main("--jobs", "0")

    def test_the_runner_finds_the_modules_discover_finds(self) -> None:
        found = run_tests.modules([])
        self.assertIn("test_run_tests.py", found)
        self.assertEqual(found, sorted(path.name for path in (ROOT / "tests").glob("test_*.py")))


if __name__ == "__main__":
    unittest.main()
