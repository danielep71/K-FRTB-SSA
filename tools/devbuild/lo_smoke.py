#!/usr/bin/env python3
"""Run the harness certification in headless LibreOffice as an early warning.

Usage: python3 tools/devbuild/lo_smoke.py [WORKBOOK]
Without WORKBOOK it builds the development workbook first.

It opens the workbook in LibreOffice with VBA compatibility and runs the
certification entry point of the example suite, then that of each project
suite, into a temporary folder, with the checkout's tests/fixtures as the
folder of the test cases. The entry points, the run logs, the certification
logs and the results files are the names that
.github/excel-evidence-policy.json gives. Then it prints a JSON summary: the
modules LibreOffice loaded, the error each certification raised, each run
log's RESULT and failures, the certification logs, the results files, and the
failures that LibreOffice is not known to cause, and what the RESULT line of a
run says that no failure line does: a run that was refused, a host state that
was not restored, a recovery that is required, and, in a project suite, a run
that is incomplete. It exits 0 when there is none, 1 when
there is one, a certification wrote no log, the build of the development
workbook fails (lo-smoke: FAIL (build: ...)), or LibreOffice could not open
the workbook or run it, and 2 when an option is unknown, WORKBOOK does not
exist, the evidence policy is missing or malformed, or LibreOffice or its
Python UNO bridge (python3-uno) is missing. Run it with the Python that
provides the bridge.

LibreOffice is not a supported host, and this is never Excel evidence. It
only catches some runtime defects before a run on the supported host. It runs
valid VBA differently in known ways:
- vbObjectError in a constant expression stops the whole project from
  compiling, so the harness writes its error numbers as literals;
- Debug.Print raises when it runs headless, so the harness prints through
  TEST_Harness.Echo;
- Application.Build does not exist, so ENVIRONMENT has an empty build;
- after a call that raised under On Error Resume Next, Err.Number reads 0,
  so the assertions on raised errors fail (KNOWN_FAILURES below);
- a cell given an error value by VBA, CVErr(#N/A), stays empty, and an error
  value typed in a cell reads as 0 in the engine, so the probes of the
  probes suite that read an error value fail (KNOWN_FAILURES below); that
  suite passes every other assertion.
"""

from __future__ import annotations

import argparse
import importlib
import importlib.util
import json
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from collections.abc import Mapping
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(ROOT / "tools"))

import check_excel_run  # noqa: E402

# Assertions of the example suite that fail in LibreOffice only, because
# Err.Number reads 0 after a call that raised under On Error Resume Next.
KNOWN_FAILURES = frozenset({
    "raised-error/negative-input",
    "cleanup-keeps-primary-error/primary-error",
    "cleanup-partial-capture/capture-error",
    "cleanup-change-failure/change-error",
    "cleanup-primary-only/primary-error",
    "cleanup-each-restoration-failure/calculation-cleanup-error",
    "cleanup-each-restoration-failure/screen-cleanup-error",
    "cleanup-only-failure/cleanup-error",
    # The probes suite, on an error value in a cell.
    "error-values-in-inputs/issue-key-of-an-issue-id-typed-as-an-error",
    "error-values-in-inputs/issue-key-of-another-issue-id-typed-as-an-error",
    "error-values-in-inputs/net-of-the-issue-of-error-ids",
    "error-values-in-inputs/net-of-the-issue-on-its-second-row",
    "error-values-in-inputs/currency-typed-as-an-error",
    "error-values-in-inputs/position-id-typed-as-an-error",
    "reading-an-error-value/error-value-is-read-as-an-error",
    "reading-an-error-value/error-value-reads-as-the-error-text",
    # The probes suite, on protection: LibreOffice's VBA mode ignores UserInterfaceOnly, so
    # AllowWriting unprotects the sheet there, and code may write a locked cell.
    "sheet-protection/write-to-a-locked-cell-is-refused",
    "sheet-protection/sheets-that-are-not-protected",
    "sheet-protection/sheet-is-still-protected",
})
# The failure each deliberate failure mode is designed to produce, by mode.
DESIGNED = {"INJECTED_FAILURE": frozenset({"exact-double/half-plus-quarter"}),
            "SKIPPED_CASE": frozenset({"harness.completeness"})}


def entry_points(policy: check_excel_run.Policy) -> tuple[str, ...]:
    """The certification entry points, as Module.Procedure: the example suite's,
    then each project suite's."""
    return (f"{Path(policy.suite_source).stem}.{policy.entry_point}",
            *(f"{Path(item.source).stem}.{item.entry_point}" for item in policy.project_suites))


def log_names(policy: check_excel_run.Policy) -> tuple[str, ...]:
    """The run log of each mode, in run order, then the certification log; then
    the run log and the certification log of each project suite."""
    return (*policy.logs.values(), policy.certification,
            *(name for item in policy.project_suites for name in (item.log, item.certification)))


def certification_logs(policy: check_excel_run.Policy) -> frozenset[str]:
    return frozenset({policy.certification,
                      *(item.certification for item in policy.project_suites)})


def designed(policy: check_excel_run.Policy) -> dict[str, frozenset[str]]:
    """The failures each deliberate failure mode is designed to produce, by its log."""
    return {policy.logs[mode]: failures for mode, failures in DESIGNED.items()}


def free_port() -> int:
    with socket.socket() as probe:
        probe.bind(("127.0.0.1", 0))
        return int(probe.getsockname()[1])


def property_value(name: str, value: object) -> Any:
    item = importlib.import_module("uno").createUnoStruct("com.sun.star.beans.PropertyValue")
    item.Name = name
    item.Value = value
    return item


def connect(port: int) -> Any:
    local = importlib.import_module("uno").getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(120):
        try:
            return resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;"
                                    "StarOffice.ComponentContext")
        except Exception:  # noqa: BLE001 - the office is still starting
            time.sleep(0.5)
    raise RuntimeError("LibreOffice did not accept a connection")


def enable_vba(context: Any) -> None:
    provider = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", context)
    access = provider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationUpdateAccess",
        (property_value("nodepath", "/org.openoffice.Office.Calc/Filter/Import/VBA"),))
    access.setPropertyValue("Load", True)
    access.setPropertyValue("Executable", True)
    access.commitChanges()


def certify(context: Any, workbook: Path, folder: Path,
            procedures: tuple[str, ...]) -> dict[str, object]:
    """Open the workbook and run each certification entry point in turn; the summary
    holds the modules loaded and, by entry point, the error it raised or None."""
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",
                                                              context)
    try:
        # Not hidden: a suite that runs the workbook's entry points reads and sets
        # Application properties that need a view. The office itself is headless.
        document = desktop.loadComponentFromURL(
            workbook.resolve().as_uri(), "_blank", 0,
            (property_value("Hidden", False), property_value("MacroExecutionMode", 4)))
    except Exception:
        # Let the office exit on its own; killing the launcher leaves it running.
        desktop.terminate()
        raise
    summary: dict[str, object] = {"workbook": str(workbook)}
    try:
        summary["modules"] = sorted(document.BasicLibraries.getByName("VBAProject")
                                    .getElementNames())
        errors: dict[str, str | None] = {}
        for procedure in procedures:
            script = document.getScriptProvider().getScript(
                f"vnd.sun.star.script:VBAProject.{procedure}?language=Basic&location=document")
            try:
                script.invoke((str(folder), str(ROOT / "tests" / "fixtures")), (), ())
                errors[procedure] = None
            except Exception as error:  # noqa: BLE001 - reported, not raised
                errors[procedure] = str(error)
        summary["errors"] = errors
    finally:
        document.close(True)
        try:
            desktop.terminate()
        except Exception:  # noqa: BLE001 - the office may already be gone
            pass
    return summary


def written(folder: Path, name: str) -> Path | None:
    """The file an entry point saved as ``name`` in ``folder``, or None. Where the
    path separator is not the backslash the entry points join with, the file is
    beside the folder, named by the folder, a backslash and the name."""
    for path in (folder / name, folder.with_name(f"{folder.name}\\{name}")):
        if path.is_file():
            return path
    return None


def read_logs(folder: Path, policy: check_excel_run.Policy) -> dict[str, object]:
    found: dict[str, object] = {}
    for name in log_names(policy):
        path = written(folder, name)
        if path is None:
            found[name] = None
            continue
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
        if name in certification_logs(policy):
            found[name] = lines
        else:
            found[name] = {"result": lines[-1] if lines else "",
                           "failures": [line for line in lines
                                        if line.startswith(("FAILURE=", "ERROR="))]}
    return found


def read_results(folder: Path, name: str) -> object:
    """The results file, parsed; its text when it is not JSON; None when it is missing.
    The workbook's build stamp names the working tree's HEAD, so it is shown, not checked."""
    path = written(folder, name)
    if path is None:
        return None
    text = path.read_text(encoding="utf-8", errors="replace")
    try:
        return json.loads(text)
    except ValueError:
        return text


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("workbook", nargs="?", type=Path,
                        help="the workbook to run; without it, the development workbook is built")
    args = parser.parse_args(argv)
    # Checked before LibreOffice starts, which would otherwise wait on a load
    # that cannot succeed.
    if args.workbook is not None and not args.workbook.is_file():
        print(f"lo-smoke: ERROR (no workbook at {args.workbook})", file=sys.stderr)
        return 2
    try:
        policy = check_excel_run.working_policy(ROOT)
    except check_excel_run.PolicyError as error:
        print(f"lo-smoke: ERROR ({error})", file=sys.stderr)
        return 2
    soffice = shutil.which("soffice")
    if soffice is None or importlib.util.find_spec("uno") is None:
        print("lo-smoke: SKIP (needs LibreOffice and python3-uno; this is never Excel evidence)")
        return 2
    with tempfile.TemporaryDirectory(prefix="lo-smoke-") as temporary:
        base = Path(temporary)
        if args.workbook is not None:
            workbook = args.workbook
        else:
            import build_workbook
            workbook = base / "Workbook-development.xlsm"
            try:
                build_workbook.build(ROOT, "development", workbook)
            except (build_workbook.BuildError, ValueError, OSError) as error:
                # A form, a module too dense to compress or an unreadable source.
                print(f"lo-smoke: FAIL (build: {error})")
                return 1
        folder = base / "evidence"
        folder.mkdir()
        port = free_port()
        office = subprocess.Popen(
            [soffice, "--headless", "--invisible", "--norestore", "--nologo",
             f"-env:UserInstallation={(base / 'profile').as_uri()}",
             f"--accept=socket,host=127.0.0.1,port={port};urp;"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        failure = None
        try:
            context = connect(port)
            enable_vba(context)
            summary = certify(context, workbook, folder, entry_points(policy))
        except Exception as error:  # noqa: BLE001 - a UNO error: reported, not raised
            failure = f"{type(error).__name__}: {error}"
        finally:
            try:
                office.wait(timeout=60)
            except subprocess.TimeoutExpired:
                office.kill()
        if failure is not None:
            print(f"lo-smoke: FAIL (LibreOffice could not run the workbook: {failure})")
            return 1
        summary["logs"] = read_logs(folder, policy)
        results = {name: read_results(folder, name)
                   for name in check_excel_run.results_files(policy)}
    summary["results"] = results[policy.results]
    summary["project_results"] = {name: content for name, content in results.items()
                                  if name != policy.results}
    found = unexpected(summary["logs"], designed(policy))  # type: ignore[arg-type]
    found += result_problems(summary["logs"], policy)  # type: ignore[arg-type]
    found += [f"{name}: not written as a test-results file" for name, content in results.items()
              if not isinstance(content, dict) or content.get("kind") != "test-results"]
    summary["unexpected"] = found
    summary["note"] = "LibreOffice is not a supported host; this is never Excel evidence"
    print(json.dumps(summary, indent=1, default=str))
    return 1 if summary["unexpected"] else 0


def unexpected(logs: dict[str, Any], expected: Mapping[str, frozenset[str]]) -> list[str]:
    """Failures and errors that LibreOffice is not known to cause, and missing logs;
    ``expected`` holds, by log, the failures a deliberate failure mode produces."""
    found = []
    for name, content in logs.items():
        if content is None:
            found.append(f"{name}: not written")
        elif isinstance(content, dict):
            for line in content["failures"]:
                key, _, rest = line.partition("=")
                identifier = rest.split(";", 1)[0]
                if key == "ERROR" or identifier not in KNOWN_FAILURES | expected.get(name, set()):
                    found.append(f"{name}: {line}")
    return found


def result_fields(line: str) -> dict[str, str]:
    """The fields of a RESULT line of the harness, such as ``cleanup`` and ``recovery``."""
    fields = {}
    for part in line.split("; "):
        name, _, value = part.partition("=")
        fields[name] = value
    return fields


def result_problems(logs: dict[str, Any], policy: check_excel_run.Policy) -> list[str]:
    """What the RESULT line of each run log says that no failure line does. A failure
    line is judged by ``unexpected``, which knows the few that LibreOffice causes; the
    result is judged here whether or not such a line is there, because the harness
    reports a host state that was not restored, or a run that was refused, without one:

    - every run must have been executed (``operation`` is not ``NOT_RUN``), and ended
      with ``cleanup=PASS`` and ``recovery=NONE``;
    - a project suite's run must also be complete, and may end ``FAIL`` only with a
      failure line to account for it: a run that passes ends ``RESULT=PASS``.

    The deliberate failure modes of the example suite end ``FAIL``, and the skipped
    case is incomplete, by design."""
    names = [*policy.logs.values(), *(item.log for item in policy.project_suites)]
    project = {item.log for item in policy.project_suites}
    found = []
    for name in names:
        content = logs.get(name)
        if not isinstance(content, dict):
            continue
        fields = result_fields(content["result"])
        problems = []
        if fields.get("RESULT") not in ("PASS", "FAIL"):
            problems.append("no RESULT")
        if fields.get("operation") in (None, "NOT_RUN", "ERROR"):
            problems.append(f"operation={fields.get('operation')}")
        if fields.get("cleanup") != "PASS":
            problems.append(f"cleanup={fields.get('cleanup')}")
        if fields.get("recovery") != "NONE":
            problems.append(f"recovery={fields.get('recovery')}")
        if name in project:
            if fields.get("completeness") != "COMPLETE":
                problems.append(f"completeness={fields.get('completeness')}")
            if fields.get("RESULT") != "PASS" and not content["failures"]:
                problems.append("it failed and names no failure")
        if problems:
            found.append(f"{name}: {', '.join(problems)}: {content['result'] or 'no RESULT'}")
    return found


if __name__ == "__main__":
    sys.exit(main())
