#!/usr/bin/env python3
"""Run the repository tests, one process per test module, several at a time.

Each ``tests/test_*.py`` runs as ``python -m unittest discover -s tests -p
<module>``, so it imports and runs as it does under ``python -m unittest
discover -s tests``. The modules share no state, so up to ``--jobs`` of them
run at once, the largest first. Each module's output is printed whole when
the module ends, followed by a summary of every module and its time.

The exit code is 0 when every module passes. Otherwise it is the first
non-zero exit code of a module, in module name order: 1 when a test fails or
errs, as under ``unittest``. A module that runs no test fails too, which
``unittest`` reports with exit code 5 from Python 3.12.

Usage:

    python tools/run_tests.py [-v] [--jobs N] [test_module.py ...]
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import time
from collections.abc import Sequence
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TESTS = ROOT / "tests"
PATTERN = "test_*.py"
RAN = re.compile(r"^Ran (\d+) tests? in ", re.MULTILINE)
SKIPPED = re.compile(r"skipped=(\d+)")
# A module that ran no test: unittest's exit code for it from Python 3.12.
NO_TESTS = 5


@dataclass
class Result:
    """How one module ran: its exit code, output, number of tests, skipped tests
    and time in seconds."""

    module: str
    code: int
    output: str
    tests: int
    skipped: int
    seconds: float


def modules(names: Sequence[str]) -> list[str]:
    """The test modules to run, by file name: those named, or every module that
    ``unittest discover`` would find in tests/."""
    if not names:
        return sorted(path.name for path in TESTS.glob(PATTERN))
    found = []
    for name in names:
        path = TESTS / Path(name).name
        if not path.is_file():
            raise SystemExit(f"run_tests.py: {name}: not a module of tests/")
        found.append(path.name)
    return found


def run(module: str, verbose: bool) -> Result:
    """Run one module in its own process."""
    command = [sys.executable, "-m", "unittest", "discover", "-s", "tests", "-p", module]
    if verbose:
        command.append("-v")
    start = time.monotonic()
    completed = subprocess.run(command, cwd=ROOT, stdout=subprocess.PIPE,
                               stderr=subprocess.STDOUT, text=True, check=False)
    seconds = time.monotonic() - start
    output = completed.stdout
    ran = RAN.findall(output)
    tests = int(ran[-1]) if ran else 0
    skipped = SKIPPED.findall(output)
    code = completed.returncode
    if code == 0 and tests == 0:
        code = NO_TESTS
    return Result(module, code, output, tests, int(skipped[-1]) if skipped else 0, seconds)


def order(names: Sequence[str]) -> list[str]:
    """The modules, largest file first: a large module takes long, and starting it
    first keeps it from finishing alone at the end."""
    return sorted(names, key=lambda name: (-(TESTS / name).stat().st_size, name))


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("modules", nargs="*", help="test module file names (default: all)")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="pass -v to unittest")
    parser.add_argument("--jobs", type=int, default=os.cpu_count() or 1,
                        help="modules run at once (default: the number of CPUs)")
    args = parser.parse_args(argv)
    if args.jobs < 1:
        parser.error("--jobs must be at least 1")
    names = modules(args.modules)
    start = time.monotonic()
    results: dict[str, Result] = {}
    with ThreadPoolExecutor(max_workers=args.jobs) as pool:
        futures = [pool.submit(run, name, args.verbose) for name in order(names)]
        for future in as_completed(futures):
            result = future.result()
            results[result.module] = result
            print(f"===== {result.module} =====", flush=True)
            print(result.output.rstrip(), flush=True)
    seconds = time.monotonic() - start
    print("===== summary =====")
    for name in names:
        result = results[name]
        state = "OK" if result.code == 0 else f"FAILED (exit code {result.code})"
        print(f"{name}: {result.tests} tests, {result.skipped} skipped, "
              f"{result.seconds:.1f} s, {state}")
    failed = [name for name in names if results[name].code != 0]
    total = sum(result.tests for result in results.values())
    skipped = sum(result.skipped for result in results.values())
    print(f"Ran {total} tests in {len(names)} modules in {seconds:.1f} s "
          f"({skipped} skipped, {args.jobs} at a time)")
    if failed:
        print(f"FAILED: {', '.join(failed)}")
        return results[failed[0]].code
    print("OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
