#!/usr/bin/env python3
"""Run one macro of a built workbook in headless LibreOffice and save the result.

Usage: python3 tools/devbuild/lo_macro.py WORKBOOK MACRO OUTPUT

It opens WORKBOOK, a workbook that tools/devbuild/build_workbook.py built, in
LibreOffice with VBA compatibility, calls MACRO, a procedure without
arguments named as Module.Procedure, recalculates the sheets and saves their
values to OUTPUT as a macro-free .xlsx. It prints "RESULT: " and what the
macro returned, which is empty for a Sub. With "-" as MACRO it calls none: it
only recalculates and saves, which is how a workbook of formulas without a
VBA project gets its values.

It exits 0 when the macro ran, 1 when LibreOffice could not open the workbook
or the macro raised (lo-macro: FAIL (...)), and 2 when an option is unknown,
WORKBOOK does not exist, or LibreOffice or its Python UNO bridge (python3-uno)
is missing. Run it with the Python that provides the bridge.

LibreOffice is not a supported host, and this is never Excel evidence: see
lo_smoke.py for the ways it runs valid VBA differently. A macro that returns
nothing where it should return a value usually means that LibreOffice
compiled no module of the project.
"""

from __future__ import annotations

import argparse
import importlib.util
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

import lo_smoke  # noqa: E402

# The filter that saves a spreadsheet as a macro-free .xlsx.
XLSX_FILTER = "Calc MS Excel 2007 XML"
# In place of a macro: recalculate and save only.
NO_MACRO = "-"


def run(context: Any, workbook: Path, macro: str, output: Path) -> str:
    """Call ``macro`` in ``workbook``, save the recalculated sheets to ``output``
    and return what the macro returned, as text."""
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",
                                                              context)
    try:
        # Not hidden: the adapter reads and sets Application properties that
        # need a view. The office itself is headless.
        document = desktop.loadComponentFromURL(
            workbook.resolve().as_uri(), "_blank", 0,
            (lo_smoke.property_value("Hidden", False),
             lo_smoke.property_value("MacroExecutionMode", 4)))
    except Exception:
        # Let the office exit on its own; killing the launcher leaves it running.
        desktop.terminate()
        raise
    try:
        returned = None
        if macro != NO_MACRO:
            script = document.getScriptProvider().getScript(
                f"vnd.sun.star.script:VBAProject.{macro}?language=Basic&location=document")
            returned = script.invoke((), (), ())[0]
        document.calculateAll()
        document.storeToURL(output.resolve().as_uri(),
                            (lo_smoke.property_value("FilterName", XLSX_FILTER),))
    finally:
        document.close(True)
        try:
            desktop.terminate()
        except Exception:  # noqa: BLE001 - the office may already be gone
            pass
    return "" if returned is None else str(returned)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("workbook", type=Path, help="a workbook built by build_workbook.py")
    parser.add_argument("macro", help="the procedure to call, as Module.Procedure, or - for none")
    parser.add_argument("output", type=Path, help="the .xlsx to write the values to")
    args = parser.parse_args(argv)
    # Checked before LibreOffice starts, which would otherwise wait on a load
    # that cannot succeed.
    if not args.workbook.is_file():
        print(f"lo-macro: ERROR (no workbook at {args.workbook})", file=sys.stderr)
        return 2
    soffice = shutil.which("soffice")
    if soffice is None or importlib.util.find_spec("uno") is None:
        print("lo-macro: SKIP (needs LibreOffice and python3-uno; this is never Excel evidence)")
        return 2
    with tempfile.TemporaryDirectory(prefix="lo-macro-") as temporary:
        port = lo_smoke.free_port()
        office = subprocess.Popen(
            [soffice, "--headless", "--invisible", "--norestore", "--nologo",
             f"-env:UserInstallation={(Path(temporary) / 'profile').as_uri()}",
             f"--accept=socket,host=127.0.0.1,port={port};urp;"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        failure = None
        returned = ""
        try:
            context = lo_smoke.connect(port)
            lo_smoke.enable_vba(context)
            returned = run(context, args.workbook, args.macro, args.output)
        except Exception as error:  # noqa: BLE001 - a UNO error: reported, not raised
            failure = f"{type(error).__name__}: {error}"
        finally:
            try:
                office.wait(timeout=60)
            except subprocess.TimeoutExpired:
                office.kill()
    if failure is not None:
        print(f"lo-macro: FAIL (LibreOffice could not run {args.macro}: {failure})")
        return 1
    print(f"RESULT: {returned}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
