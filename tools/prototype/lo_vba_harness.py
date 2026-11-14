"""Run the exported VBA modules inside LibreOffice's VBA compatibility mode.

Usage:
  python tools/lo_vba_harness.py WORKBOOK OUT MACRO [MODULE.bas ...]

Loads WORKBOOK, inserts each .bas module into the document's Standard Basic
library, calls MACRO ("Module.Function"; "-" for none), recalculates and
saves the document to OUT as .xlsx. Prints "RESULT: <return value>" or
"RESULT: FAILED <error>".

This is a test harness for Linux and CI. It proves the logic and the
workbook wiring, not compilation in Excel. Three transformations are
applied to the source on the way in, never to the files: "Option
VBASupport 1" is prepended, vbObjectError is replaced by its value and
Debug.Print statements are commented out, because LibreOffice Basic knows
neither. Class modules are not supported.
"""
import os
import re
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

WRAPPER = '''Option VBASupport 1
Option Explicit
Public Function HarnessCall() As String
        On Error GoTo Failed
        HarnessCall = CStr({macro}())
        Exit Function
Failed:
        HarnessCall = "FAILED " & Err.Number & " | " & Err.Source & " | " & Err.Description
End Function
'''


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def connect():
    pipe = f"k_ssa_{os.getpid()}"
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    proc = subprocess.Popen(
        ["soffice", "--headless", "--invisible", "--norestore", "--nologo",
         f"-env:UserInstallation=file://{profile}",
         f"--accept=pipe,name={pipe};urp;StarOffice.ComponentContext"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(120):
        try:
            return proc, resolver.resolve(
                f"uno:pipe,name={pipe};urp;StarOffice.ComponentContext")
        except Exception:
            time.sleep(0.5)
    proc.kill()
    raise RuntimeError("LibreOffice did not start")


def basic_source(path):
    text = open(path, encoding="cp1252").read().replace("\r\n", "\n")
    if text.startswith("VERSION 1.0 CLASS"):
        raise ValueError(f"{path}: class modules are not supported by the harness")
    lines = [ln for ln in text.split("\n") if not ln.startswith("Attribute VB_")]
    body = "\n".join(lines)
    body = re.sub(r"\bvbObjectError\b", "(-2147221504)", body)
    body = re.sub(r"^(\s*)Debug\.Print", r"\1'Debug.Print", body, flags=re.M)
    return "Option VBASupport 1\n" + body


def run(workbook, out, macro, modules):
    proc, ctx = connect()
    desktop = None
    try:
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        url = uno.systemPathToFileUrl(os.path.abspath(workbook))
        # A view is needed for Application.Calculation and ScreenUpdating.
        doc = desktop.loadComponentFromURL(url, "_blank", 0,
                                           (prop("Hidden", False), prop("MacroExecutionMode", 4)))
        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        lib = libs.getByName("Standard")
        for path in modules:
            name = os.path.splitext(os.path.basename(path))[0]
            code = basic_source(path)
            if lib.hasByName(name):
                lib.replaceByName(name, code)
            else:
                lib.insertByName(name, code)
        result = ""
        if macro != "-":
            lib.insertByName("HARNESS_Wrapper", WRAPPER.format(macro=macro))
            script = doc.getScriptProvider().getScript(
                "vnd.sun.star.script:Standard.HARNESS_Wrapper.HarnessCall"
                "?language=Basic&location=document")
            result = script.invoke((), (), ())[0]
            if result == "":
                result = "FAILED no result: a module did not compile under LibreOffice"
        doc.calculateAll()
        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(out)),
                       (prop("FilterName", "Calc MS Excel 2007 XML"),))
        doc.close(True)
        return result
    finally:
        try:
            if desktop is not None:
                desktop.terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()


if __name__ == "__main__":
    wb_path, out_path, macro_name, *mods = sys.argv[1:]
    print("RESULT:", run(wb_path, out_path, macro_name, mods))
