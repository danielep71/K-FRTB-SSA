#!/usr/bin/env python3
"""Build the K-FRTB-SSA workbook in one of two forms.

Usage: python tools/prototype/build_prototype.py [OUTPUT] [--template]

Without --template the script writes the formula prototype: every output is
a live formula. With --template it writes the macro-free template for the
VBA engine: the output areas are empty and RunSSA fills them.

Both forms share layout, names, sample data and the TestCatalogue sheet, so
the formula prototype is a second implementation to compare the VBA engine
with. OUTPUT defaults to K-FRTB-SSA_Prototype.xlsx in the current directory.
The sample data and the expected test values come from ssa_reference.py.

The template is tracked as src/workbook/Template.xlsx, and
tools/devbuild/build_workbook.py builds the workbook from it. After a change
here, write it again:

    python tools/prototype/build_prototype.py src/workbook/Template.xlsx --template

tests/test_prototype.py fails while the tracked file is not what this
module writes. Two builds are byte-identical in one environment; across
environments openpyxl may serialize the same XML differently, so the test
compares the parts as XML.
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, NamedTuple

import ssa_reference as D
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string as CI
from openpyxl.utils import get_column_letter as COL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# openpyxl ships no type information: its objects are read as Any.
Book = Any
Sheet = Any
Cell = Any

DEFAULT_OUTPUT = "K-FRTB-SSA_Prototype.xlsx"
PRODUCT = "K-FRTB-SSA"
REPOSITORY = "danielep71/K-FRTB-SSA"
# The labels of the build stamp on the README sheet and the names of its cells.
STAMP = (("Commit", "build_commit"), ("Commit time", "build_time"),
         ("Configuration", "build_configuration"))
# Excel's last date serial, 31 December 9999: a larger number is not a date.
DATE_SERIAL_MAX = 2958465
# Timestamps of the package, so that two builds are byte-identical: the
# creation and modification times of its properties, and of its zip entries.
PACKAGE_TIME = "2026-01-01T00:00:00Z"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)
FONT = "Aptos"
FIRST, CAP = 5, 500
LAST = FIRST + CAP - 1
NFX = 20  # FX_Table capacity
NBANDS = 15
NBKT = 30  # Buckets capacity per block

# First and last row of IR_Ladder (one block of NBANDS rows per FX_Table row)
# and first, last and total row of IR_Zones (one row per FX_Table row).
LAD_F, LAD_L = 6, 6 + NFX * NBANDS - 1
ZF, ZL = 6, 6 + NFX - 1
ZT = ZL + 1

# Sheet order: home, inputs, outputs, tests, readme.
SHEET_ORDER = ["Home", "Params", "IRPositions", "EQPositions", "COPositions", "FXPositions", "OptionCharges", "Results",
               "IR_Zones", "IR_Ladder", "Buckets", "IRCalc", "EQCalc", "COCalc", "FXCalc", "Checks", "TestCatalogue", "README"]

# ---------------------------------------------------------------- styles
NAVY, BLUE, BAND2, GOLD = "1F3864", "2E75B6", "265A99", "FFC000"
INPUT_FILL, CONST_FILL, TOTAL_FILL = "FFF2CC", "D9D9D9", "F2F2F2"
GREY_TXT, NOTE_TXT, LINK = "595959", "7F6000", "0563C1"
TAB = {"home": "1F3864", "input": "BF9000", "output": "2E75B6", "test": "548235"}
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NF_AMT, NF_PCT, NF_DATE, NF_RATE = '#,##0;-#,##0;0', "0.00%", "yyyy-mm-dd", "0.0000##"

RED: dict[str, Any] = dict(font=Font(color="FF9C0006"), fill=PatternFill(bgColor="FFFFC7CE"))
GREEN: dict[str, Any] = dict(font=Font(color="FF375623"), fill=PatternFill(bgColor="FFE2EFDA"))
AMBER: dict[str, Any] = dict(font=Font(color="FF9C5700"), fill=PatternFill(bgColor="FFFFEB9C"))

GRADES = '"AAA to AA-,A+ to BBB-,BB+ to BB-,B+ to B-,Below B-,Unrated"'
# Status test shared by the four calculation sheets; {r} is the row number.
# The PositionIDs are compared as the engine compares them, as text without
# regard to case. COUNTIF would read "*", "?" and "~" in a PositionID as a
# pattern, and a leading "<", ">" or "=" as an operator.
DUP = 'SUMPRODUCT(--(UPPER($C$5:$C$504)=UPPER($C{r})))>1'
# The OK rows, down to row {r}, whose key in column E is that of row {r}: 1 on
# the first row of a market or of a commodity.
FIRST_OK_ROW_OF_KEY = ('SUMPRODUCT((UPPER($E$5:$E{r})=UPPER($E{r}))*($D$5:$D{r}="OK")*1)')
# True when a numeric cell is outside the serials Excel holds as dates.
NOT_A_DATE = f"OR({{cell}}<1,{{cell}}>{DATE_SERIAL_MAX})"

# The created and modified elements of docProps/core.xml.
CORE_TIME = re.compile(rb"(<dcterms:(?:created|modified)[^>]*>)([^<]*)(</dcterms:(?:created|modified)>)")

# One rule of the Checks sheet: rule, severity, sheet, count formula, reference, message.
Rule = tuple[str, str, str, str, str, str]


class Block(NamedTuple):
    """First, last and total row of one table on Buckets."""
    first: int
    last: int
    total: int


@dataclass
class Context:
    """The workbook being built and what later sheets need from earlier ones."""
    wb: Book
    template: bool
    param_last: int = 0             # last row of the parameter list on Params
    eq: Block = Block(0, 0, 0)      # equity table on Buckets
    co: Block = Block(0, 0, 0)      # commodity table on Buckets
    fx: Block = Block(0, 0, 0)      # FX table on Buckets
    rule_last: int = 0              # last rule row on Checks
    findings: int = 0               # row of the findings section on Checks


# --------------------------------------------------------------- helpers
def fill(c: str) -> Any:
    return PatternFill("solid", fgColor="FF" + c)


def font(sz: float = 9, b: bool = False, i: bool = False, color: str = "000000") -> Any:
    return Font(name=FONT, size=sz, bold=b, italic=i, color="FF" + color)


def new_sheet(wb: Book, name: str, title: str, subtitle: str, lastcol: str, tab: str,
              widths: Mapping[str, float], sub_h: float = 30.0) -> Sheet:
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = "FF" + TAB[tab]
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 1.5
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    lc = CI(lastcol)
    for c in range(2, lc + 1):
        ws.cell(1, c).fill = fill("000000")
        ws.cell(2, c).fill = fill(BAND2)
    ws["B1"].value, ws["B1"].font = PRODUCT, Font(name=FONT, size=14, color="FF" + GOLD)
    ws["B2"].value, ws["B2"].font = title, Font(name=FONT, size=18, bold=True, color="FFFFFFFF")
    ws["B1"].alignment = ws["B2"].alignment = Alignment(horizontal="left", vertical="center")
    ws["C3"].value, ws["C3"].font = subtitle, font(10, i=True, color=GREY_TXT)
    ws.row_dimensions[1].height, ws.row_dimensions[2].height = 24, 26.1
    ws.row_dimensions[3].height = sub_h
    ws.row_dimensions[4].height = 9.95
    return ws


def header(ws: Sheet, row: int, start: str, labels: Sequence[str], kind: str = "output", sz: float = 11,
           height: float = 42, center_from: int = 1) -> None:
    c0 = CI(start)
    for k, lab in enumerate(labels):
        c = ws.cell(row, c0 + k, lab)
        c.font = font(sz, b=True, color="FFFFFF")
        c.fill = fill(NAVY if kind == "input" else BLUE)
        c.border = BOX
        c.alignment = Alignment(horizontal="center" if k >= center_from else None,
                                vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = height


def section(ws: Sheet, row: int, text: str, c1: str = "C", c2: str | None = None) -> None:
    ws[f"{c1}{row}"].value = text
    med = Side(style="medium", color="FF" + NAVY)
    for c in range(CI(c1), CI(c2 or c1) + 1):
        ws.cell(row, c).font = font(14, b=True, color=NAVY)
        ws.cell(row, c).border = Border(bottom=med)
    ws.row_dimensions[row].height = 30


def put(ws: Sheet, ref: str, value: Any, nf: str | None = None, kind: str | None = "calc", b: bool = False,
        align: str | None = None, color: str = "000000", sz: float = 9) -> Cell:
    c = ws[ref]
    c.value = value
    c.font = font(sz, b=b, color=color)
    if kind == "input":
        c.fill, c.border = fill(INPUT_FILL), BOX
    elif kind == "const":
        c.fill, c.border = fill(CONST_FILL), BOX
    elif kind == "calc":
        c.border = BOX
    elif kind == "total":
        c.fill, c.border = fill(TOTAL_FILL), BOX
        c.font = font(sz, b=True, color=color)
    if nf:
        c.number_format = nf
    if align:
        c.alignment = Alignment(horizontal=align, vertical="center")
    return c


def name(wb: Book, nm: str, ref: str) -> None:
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def link(ws: Sheet, ref: str, sheet: str, text: str | None = None) -> None:
    c = ws[ref]
    c.value = text or sheet
    c.hyperlink = f"#'{sheet}'!A1"
    c.font = font(9, color=LINK)


def dv_list(ws: Sheet, rng: str, items: str) -> None:
    dv = DataValidation(type="list", formula1=items, allow_blank=True)
    dv.error, dv.errorTitle, dv.showErrorMessage = "Pick a value from the list.", PRODUCT, True
    ws.add_data_validation(dv)
    dv.add(rng)


def cf_status(ws: Sheet, rng: str, col: str, first: int) -> None:
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(${col}{first},7)="INVALID"'], **RED))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(${col}{first},8)="EXCLUDED"'], **AMBER))


def text_block(ws: Sheet, row: int, items: Sequence[tuple[str, str]], wide: bool = False) -> int:
    """Two-column label / text rows in the Home and README style."""
    dashed = Side(style="dashed", color="FFBFBFBF")
    for k, (lab, txt) in enumerate(items):
        last = k == len(items) - 1
        bd = Border(bottom=THIN if last else dashed)
        a, b = ws.cell(row, 3, lab), ws.cell(row, 4, txt)
        a.font, b.font = font(10, b=True), font(9)
        a.border = b.border = bd
        a.alignment = b.alignment = Alignment(vertical="center", wrap_text=True)
        lines = max(1, len(txt) // (150 if wide else 125) + 1)
        ws.row_dimensions[row].height = max(30, 14 * lines + 8)
        row += 1
    return row


def blank(ws: Sheet, rows: Sequence[int], cols: Sequence[int], keep: tuple[str, ...] = ()) -> None:
    for r_ in rows:
        for c_ in cols:
            ref = f"{COL(c_)}{r_}"
            if ref not in keep:
                ws[ref].value = None


# ------------------------------------------------------- range references
def R(c: str) -> str:
    """Column c of the position rows, on the sheet that holds the formula."""
    return f"${c}${FIRST}:${c}${LAST}"


# The text the engine reads an error value in a cell as.
ERROR_TEXT = "#ERROR"
# The column of IRCalc, right of its output columns, that holds the key the
# rows of an issue are netted on. Only the formula workbook has it.
NETTING_KEY = "Z"


# ----------------------------------------------------- literal comparison
# The formula workbook groups, counts and looks up rows by typed keys as the
# engine does: as text, without regard to case. COUNTIF, SUMIFS and MATCH
# would read "*", "?" and "~" in a key as a pattern, and a leading "<", ">"
# or "=" as an operator.
def same(cells: str, key: str) -> str:
    """An array expression: TRUE for the cells whose text is key, without regard to case."""
    return f"(UPPER({cells})=UPPER({key}))"


def count_where(*conditions: str) -> str:
    """The number of rows where every condition holds."""
    return "SUMPRODUCT(" + "*".join(conditions) + "*1)"


def sum_where(values: str, *conditions: str) -> str:
    """The sum of values over the rows where every condition holds.

    The values are a second argument of SUMPRODUCT, which reads a text among
    them as zero."""
    return "SUMPRODUCT(" + "*".join(conditions) + f"*1,{values})"


def ok(status: str) -> str:
    """The rows whose status is OK."""
    return f'({status}="OK")'


def positive(cells: str) -> str:
    return f"({cells}>0)"


def negative(cells: str) -> str:
    return f"({cells}<0)"


def no_currency(code: str) -> str:
    """TRUE when a currency code is blank or not in FX_Table."""
    return f'OR({code}="",{count_where(same("FX_Ccy", code))}=0)'


def fx_rate(code: str) -> str:
    """The rate of a currency code of FX_Table."""
    return sum_where("FX_Rate", same("FX_Ccy", code))


def text(cell: str) -> str:
    """A typed text as the engine reads it: trimmed, and ERROR_TEXT for an error value.

    An error value read as it is would make an error of every formula that
    compares the whole column, whatever the row it is on."""
    return f'IFERROR(TRIM({cell}),"{ERROR_TEXT}")'


def typed(cell: str) -> str:
    """A typed text that is compared as it is: ERROR_TEXT for an error value."""
    return f'IFERROR({cell},"{ERROR_TEXT}")'


def neither(cell: str, first: str, second: str) -> str:
    """TRUE when a typed text is neither of two words."""
    return f'AND({typed(cell)}<>"{first}",{typed(cell)}<>"{second}")'


def issue_rows(last: str) -> str:
    """The interest rate rows, down to row last, of the issue of row {r}.

    The Issue key shown is not enough: a position without an Issue ID shows
    "POS:" and its PositionID, which another row can hold as its entered Issue
    ID. The two never net, so the rows are matched on the netting key."""
    return same(f"${NETTING_KEY}${FIRST}:${NETTING_KEY}{last}", f"${NETTING_KEY}{{r}}")


def IC(c: str) -> str:
    """Column c of the position rows of IRCalc."""
    return f"IRCalc!${c}${FIRST}:${c}${LAST}"


def EC(c: str) -> str:
    """Column c of the position rows of EQCalc."""
    return f"EQCalc!${c}${FIRST}:${c}${LAST}"


def CC(c: str) -> str:
    """Column c of the position rows of COCalc."""
    return f"COCalc!${c}${FIRST}:${c}${LAST}"


def FC(c: str) -> str:
    """Column c of the position rows of FXCalc."""
    return f"FXCalc!${c}${FIRST}:${c}${LAST}"


def LD(c: str) -> str:
    """Column c of the band rows of IR_Ladder."""
    return f"IR_Ladder!${c}${LAD_F}:${c}${LAD_L}"


def SUMR(sh: str) -> str:
    """Count of the rejected rows of calculation sheet sh."""
    return f'=COUNTIF({sh}!$D${FIRST}:$D${LAST},"INVALID*")'


# =================================================================== Home
def build_home_frame(ctx: Context) -> None:
    """Create the Home sheet first; fill_home writes its content last."""
    new_sheet(ctx.wb, "Home", "Home", "", "E", "home", {"C": 30, "D": 130, "E": 4.8}, 33.95)


# ================================================================= Params
def build_params(ctx: Context) -> None:
    ws = new_sheet(ctx.wb, "Params", "Parameters",
                   "Each value is read through the workbook name equal to its code in column C. Yellow cells are inputs; "
                   "grey cells are MAR40 constants. Change a grey cell only for a what-if run: Checks flags every value that "
                   "differs from column G.", "K", "input",
                   {"C": 25.6, "D": 52, "E": 15.6, "F": 34, "G": 12.6, "H": 11.6, "I": 14, "J": 14, "K": 30})
    header(ws, 5, "C", ["Code", "Description", "Value", "Reference", "MAR40 value", "Changed"], "input", 12, 30, 2)
    ctx.param_last = params_list(ctx.wb, ws)
    ws.conditional_formatting.add(f"H6:H{ctx.param_last}", CellIsRule(operator="equal", formula=['"CHANGED"'], **AMBER))
    ws.freeze_panes = "A6"
    r = params_sr_table(ctx.wb, ws, ctx.param_last + 2)
    r = params_ladder(ctx.wb, ws, r)
    params_fx_table(ctx.wb, ws, r)


def params_list(wb: Book, ws: Sheet) -> int:
    """Write the parameter list from row 6 and return its last row."""
    params: list[tuple[str, str, date | str | float, str, str, str, str | float | None]] = [
        ("AsOfDate", "Reporting (calculation) date", D.AS_OF, "input", NF_DATE, "", None),
        ("ReportingCcy", "Reporting currency (all amounts converted into it)", D.REPORTING_CCY, "input", "@", "", None),
        ("DaysPerYear", "Calendar days per year for residual maturities", 365, "const", "0", "Workbook convention (Act/365)", 365),
        ("CouponThreshold", "Coupon below which the low-coupon time bands apply", 0.03, "const", NF_PCT, "MAR40.26 Table 4", 0.03),
        ("ScalingIR", "Scaling factor, interest rate risk", 1.30, "const", "0.00", "MAR40.2", 1.30),
        ("ScalingEQ", "Scaling factor, equity risk", 3.50, "const", "0.00", "MAR40.2", 3.50),
        ("ScalingCO", "Scaling factor, commodity risk", 1.90, "const", "0.00", "MAR40.2", 1.90),
        ("ScalingFX", "Scaling factor, foreign exchange risk", 1.20, "const", "0.00", "MAR40.2", 1.20),
        ("RWAMultiplier", "RWA = requirement x multiplier", 12.5, "const", "0.0", "MAR40.1", 12.5),
        ("VerticalDisallowance", "Vertical disallowance on matched weighted positions per time band", 0.10, "const", NF_PCT, "MAR40.27", 0.10),
        ("HorizontalZone1", "Horizontal disallowance within zone 1", 0.40, "const", NF_PCT, "MAR40.28 Table 5", 0.40),
        ("HorizontalZone2", "Horizontal disallowance within zone 2", 0.30, "const", NF_PCT, "MAR40.28 Table 5", 0.30),
        ("HorizontalZone3", "Horizontal disallowance within zone 3", 0.30, "const", NF_PCT, "MAR40.28 Table 5", 0.30),
        ("HorizontalAdjacent", "Horizontal disallowance between adjacent zones (1-2, 2-3)", 0.40, "const", NF_PCT, "MAR40.28 Table 5", 0.40),
        ("HorizontalZone13", "Horizontal disallowance between zones 1 and 3", 1.00, "const", NF_PCT, "MAR40.28 Table 5", 1.00),
        ("EQSpecificRate", "Equity specific risk, single names (on net position per stock)", 0.08, "const", NF_PCT, "MAR40.43, 40.46", 0.08),
        ("EQIndexRate", "Equity further charge on net position per index contract", 0.02, "const", NF_PCT, "MAR40.47", 0.02),
        ("EQGeneralRate", "Equity general market risk (on net position per market)", 0.08, "const", NF_PCT, "MAR40.43", 0.08),
        ("FXChargeRate", "FX and gold, on the overall net open position", 0.08, "const", NF_PCT, "MAR40.61", 0.08),
        ("GoldCode", "Currency code treated as gold", "XAU", "const", "@", "MAR40.60", "XAU"),
        ("COMDirectionalRate", "Commodity directional charge (net position per commodity)", 0.15, "const", NF_PCT, "MAR40.72", 0.15),
        ("COMGrossRate", "Commodity charge on gross positions at spot", 0.03, "const", NF_PCT, "MAR40.73", 0.03),
    ]
    r = 6
    for code, desc, val, kind, nf, ref, mar in params:
        put(ws, f"C{r}", code, b=True, kind=None, sz=10)
        put(ws, f"D{r}", desc, kind=None)
        put(ws, f"E{r}", val, nf, kind, align="center")
        put(ws, f"F{r}", ref, kind=None, color=GREY_TXT)
        if mar is not None:
            put(ws, f"G{r}", mar, nf, "calc", align="center", color=GREY_TXT)
            put(ws, f"H{r}", f'=IF(E{r}=G{r},"","CHANGED")', kind="calc", align="center")
        name(wb, code, f"Params!$E${r}")
        ws.row_dimensions[r].height = 20.1
        r += 1
    return r - 1


def params_sr_table(wb: Book, ws: Sheet, r: int) -> int:
    """Write the IR specific risk table from row r and return the row of the next section."""
    section(ws, r, "Interest rate specific risk charges (MAR40.6 Table 1)", "C", "H")
    r += 1
    header(ws, r, "C", ["SR_Key", "Category", "Rating grade", "Residual maturity", "Charge", "Reference"], "input", 12, 30, 1)
    r += 1
    sr_first = r
    for (cat, grd, mb), ch in D.SR.items():
        put(ws, f"C{r}", f"{cat}|{grd}|{mb}", kind=None, b=True, sz=10)
        put(ws, f"D{r}", cat, kind=None, align="center")
        put(ws, f"E{r}", grd, kind=None, align="center")
        put(ws, f"F{r}", {"Any": "Any", "<=6m": "6 months or less", "6m-24m": "over 6 to 24 months", ">24m": "over 24 months"}[mb], kind=None, align="center")
        put(ws, f"G{r}", ch, NF_PCT, "const", align="center")
        put(ws, f"H{r}", "MAR40.6", kind=None, color=GREY_TXT)
        r += 1
    sr_last = r - 1
    name(wb, "SR_Key", f"Params!$C${sr_first}:$C${sr_last}")
    name(wb, "SR_Charge", f"Params!$G${sr_first}:$G${sr_last}")
    put(ws, f"C{r}", "GOV = central governments and central banks; QUAL = qualifying (investment-grade PSEs, MDBs, corporates); "
        "OTHER = all other issuers; NONE = derivative legs and positions without issuer risk. A category / grade pair not in "
        "this table is rejected (for example OTHER with an investment-grade rating: use QUAL).", kind=None, color=GREY_TXT)
    return r + 2


def params_ladder(wb: Book, ws: Sheet, r: int) -> int:
    """Write the maturity ladder from row r and return the row of the next section."""
    section(ws, r, "General market risk: maturity method time bands (MAR40.26 Table 4)", "C", "J")
    r += 1
    header(ws, r, "C", ["Band", "Zone", "Upper bound, coupon >= 3% (years)", "Upper bound, coupon < 3% (years)",
                        "Time band, coupon >= 3%", "Time band, coupon < 3%", "Risk weight", "Assumed yield change (pp)"],
           "input", 11, 42, 1)
    r += 1
    lab_hi = ["1 month or less", "1 to 3 months", "3 to 6 months", "6 to 12 months", "1 to 2 years", "2 to 3 years",
              "3 to 4 years", "4 to 5 years", "5 to 7 years", "7 to 10 years", "10 to 15 years", "15 to 20 years",
              "over 20 years", "", ""]
    lab_lo = ["1 month or less", "1 to 3 months", "3 to 6 months", "6 to 12 months", "1.0 to 1.9 years",
              "1.9 to 2.8 years", "2.8 to 3.6 years", "3.6 to 4.3 years", "4.3 to 5.7 years", "5.7 to 7.3 years",
              "7.3 to 9.3 years", "9.3 to 10.6 years", "10.6 to 12 years", "12 to 20 years", "over 20 years"]
    ayc = [1.0, 1.0, 1.0, 1.0, 0.90, 0.80, 0.75, 0.75, 0.70, 0.65, 0.60, 0.60, 0.60, 0.60, 0.60]
    lad_first = r
    for k, (b, z, uh, ul, w) in enumerate(D.LADDER):
        put(ws, f"C{r}", b, "0", None, b=True, align="center", sz=10)
        put(ws, f"D{r}", z, "0", "const", align="center")
        put(ws, f"E{r}", uh, "0.0000", "const", align="center")
        put(ws, f"F{r}", ul, "0.0000", "const", align="center")
        put(ws, f"G{r}", lab_hi[k], kind=None, align="center")
        put(ws, f"H{r}", lab_lo[k], kind=None, align="center")
        put(ws, f"I{r}", w, NF_PCT, "const", align="center")
        put(ws, f"J{r}", ayc[k], "0.00", "const", align="center")
        r += 1
    lad_last = r - 1
    for nm, c in [("Ladder_Band", "C"), ("Ladder_Zone", "D"), ("Ladder_UpperHigh", "E"), ("Ladder_UpperLow", "F"),
                  ("Ladder_LabelHigh", "G"), ("Ladder_LabelLow", "H"), ("Ladder_RW", "I")]:
        name(wb, nm, f"Params!${c}${lad_first}:${c}${lad_last}")
    put(ws, f"C{r}", "Upper bounds are inclusive (\"1 month or less\"). A blank upper bound means the band is open-ended. "
        "Bands 14-15 apply to coupons below 3% only. The assumed yield changes are shown for the duration method, "
        "which this prototype does not implement.", kind=None, color=GREY_TXT)
    return r + 2


def params_fx_table(wb: Book, ws: Sheet, r: int) -> None:
    """Write the FX table from row r."""
    section(ws, r, "FX rates (units of reporting currency per unit of currency)", "C", "F")
    r += 1
    header(ws, r, "C", ["Currency", "Description", "Rate to reporting ccy", "Note"], "input", 12, 30, 2)
    r += 1
    fx_first = r
    for k in range(NFX):
        c: str | None
        rate: float | None
        desc: str | None
        if k < len(D.FX):
            c, rate, desc = D.FX[k]
        else:
            c, rate, desc = None, None, None
        put(ws, f"C{r}", c, "@", "input", b=True, align="center")
        put(ws, f"D{r}", desc, None, "input")
        put(ws, f"E{r}", rate, NF_RATE, "input", align="center")
        if c == "XAU":
            put(ws, f"F{r}", "EUR per troy ounce; FX positions in XAU are entered in ounces", kind=None, color=GREY_TXT)
        if c in ("USD", "GBP", "CAD", "SEK", "NOK"):
            put(ws, f"F{r}", "Illustrative; rounded to reproduce MAR40 Table 9", kind=None, color=GREY_TXT)
        r += 1
    fx_last = r - 1
    name(wb, "FX_Ccy", f"Params!$C${fx_first}:$C${fx_last}")
    name(wb, "FX_Rate", f"Params!$E${fx_first}:$E${fx_last}")
    put(ws, f"C{r}", "The reporting currency must be listed with rate 1. Rates are illustrative: replace them with the "
        "reporting-date fixing.", kind=None, color=GREY_TXT)


# ============================================================ input sheets
def input_sheet(wb: Book, name_: str, title: str, subtitle: str, cols: Sequence[tuple[str, float]],
                rows: Sequence[Sequence[object]], lastcol: str, dvs: Mapping[str, str],
                numfmts: Mapping[str, str], freeze: str = "D5") -> Sheet:
    widths = {c: w for c, (_, w) in zip([COL(3 + k) for k in range(len(cols))], cols)}
    ws = new_sheet(wb, name_, title, subtitle, lastcol, "input", widths)
    header(ws, 4, "C", [lab for lab, _ in cols], "input", 11, 50.1, 2)
    name(wb, f"{name_}_Header", f"{name_}!$C$4")
    for i in range(CAP):
        rr = FIRST + i
        vals: Sequence[object] = rows[i] if i < len(rows) else [None] * len(cols)
        for k, v in enumerate(vals):
            cl = COL(3 + k)
            c = ws.cell(rr, 3 + k, v)
            c.font, c.fill, c.border = font(9), fill(INPUT_FILL), BOX
            c.number_format = numfmts.get(cl, "General")
            if cl not in ("C", "D") and k < len(vals) - 1 and numfmts.get(cl) in ("@", None):
                c.alignment = Alignment(horizontal="center")
        ws.row_dimensions[rr].height = 20.1 if i < 60 else None
    for cl, items in dvs.items():
        dv_list(ws, f"{cl}{FIRST}:{cl}{LAST}", items)
    ws.freeze_panes = freeze
    return ws


def build_ir_positions(ctx: Context) -> None:
    input_sheet(
        ctx.wb, "IRPositions", "Interest rate positions",
        "One row per debt position or derivative leg (MAR40.24-40.40). Swaps, futures, FRAs and forwards are entered as two "
        "notional legs. Position value is positive; Direction gives the sign. Floating legs are slotted by the next repricing date.",
        [("PositionID", 15.6), ("Portfolio", 14.6), ("Instrument", 12.6), ("Currency", 9.6), ("Direction", 9.6),
         ("Position value (ccy)", 15.6), ("Rate type", 10.6), ("Coupon (decimal)", 10.6), ("Final maturity date", 12.6),
         ("Next repricing date (floating)", 12.6), ("SR category", 10.6), ("Rating grade", 13.6),
         ("Issue ID (specific-risk netting)", 22.6), ("Comment", 44)],
        [[p[0], p[1], p[2], p[3], p[4], p[5], p[6], p[7], p[8], p[9], p[10], p[11], p[12], p[13]] for p in D.IR],
        "P", {"E": '"Bond,FRN,Swap leg,Future leg,FRA leg,Forward leg,Other"', "F": "=FX_Ccy", "G": '"Long,Short"',
              "I": '"Fixed,Floating"', "M": '"GOV,QUAL,OTHER,NONE"', "N": GRADES},
        {"C": "@", "D": "@", "E": "@", "F": "@", "G": "@", "H": "#,##0", "I": "@", "J": "0.000%", "K": NF_DATE,
         "L": NF_DATE, "M": "@", "N": "@", "O": "@", "P": "@"})


def build_eq_positions(ctx: Context) -> None:
    input_sheet(
        ctx.wb, "EQPositions", "Equity positions",
        "One row per equity or index position, delta-equivalent for derivatives (MAR40.41-40.50). Market groups positions "
        "for general market risk; positions in the same stock or index in a market net before the specific charge.",
        [("PositionID", 15.6), ("Portfolio", 14.6), ("Market", 9.6), ("Type", 12.6), ("Stock or index identifier", 20.6),
         ("Currency", 9.6), ("Direction", 9.6), ("Market value (ccy)", 15.6), ("Comment", 44)],
        [list(p) for p in D.EQ], "K",
        {"E": '"Single name,Index"', "H": "=FX_Ccy", "I": '"Long,Short"'},
        {"C": "@", "D": "@", "E": "@", "F": "@", "G": "@", "H": "@", "I": "@", "J": "#,##0", "K": "@"})


def build_co_positions(ctx: Context) -> None:
    input_sheet(
        ctx.wb, "COPositions", "Commodity positions",
        "One row per commodity position, delta-equivalent for derivatives (MAR40.64-40.73). Value = quantity x spot price, "
        "converted at the FX table. Positions with the same commodity name net. The maturity date is kept for the "
        "maturity-ladder approach, which this prototype does not implement.",
        [("PositionID", 15.6), ("Portfolio", 14.6), ("Commodity", 18.6), ("Direction", 9.6), ("Quantity", 12.6),
         ("Unit", 8.6), ("Spot price (ccy per unit)", 12.6), ("Currency", 9.6), ("Maturity date", 12.6), ("Comment", 40)],
        [list(p) for p in D.CO], "L",
        {"F": '"Long,Short"', "J": "=FX_Ccy"},
        {"C": "@", "D": "@", "E": "@", "F": "@", "G": "#,##0.##", "H": "@", "I": "#,##0.00", "J": "@", "K": NF_DATE, "L": "@"})


def build_fx_positions(ctx: Context) -> None:
    input_sheet(
        ctx.wb, "FXPositions", "FX and gold positions",
        "Net positions per currency from all sources (spot, forwards, delta of options, structural positions where "
        "included), in units of the currency; gold (XAU) in troy ounces (MAR40.53-40.61). Positions in the reporting "
        "currency carry no FX risk and are excluded.",
        [("PositionID", 15.6), ("Portfolio", 14.6), ("Currency", 9.6), ("Direction", 9.6), ("Amount (ccy)", 17.6),
         ("Comment", 44)],
        [list(p) for p in D.FXPOS], "H",
        {"E": "=FX_Ccy", "F": '"Long,Short"'},
        {"C": "@", "D": "@", "E": "@", "F": "@", "G": "#,##0", "H": "@"})


# --- OptionCharges
def build_option_charges(ctx: Context) -> None:
    ws = new_sheet(ctx.wb, "OptionCharges", "Option charges",
                   "Non-delta option charges per risk class, calculated outside the workbook and entered in reporting "
                   "currency (MAR40.74-40.84). They are added to the class before scaling (MAR40.2). Delta equivalents go on "
                   "the position sheets.", "G", "input", {"C": 14.6, "D": 30.6, "E": 18.6, "F": 18.6, "G": 60})
    header(ws, 4, "C", ["Risk class", "Description", "Method", "Charge (rep ccy)", "Comment"], "input", 11, 42, 2)
    oc_rows = [("IR", "Interest rate options"), ("EQ", "Equity options"), ("CO", "Commodity options"),
               ("FX", "FX and gold options")]
    for k, (rc, desc) in enumerate(oc_rows):
        rr = 5 + k
        put(ws, f"C{rr}", rc, "@", None, b=True, sz=10)
        put(ws, f"D{rr}", desc, kind=None)
        put(ws, f"E{rr}", "None", "@", "input", align="center")
        put(ws, f"F{rr}", 0, NF_AMT, "input", align="center")
        put(ws, f"G{rr}", "", "@", "input")
        name(ctx.wb, f"Options{rc}", f"OptionCharges!$F${rr}")
        ws.row_dimensions[rr].height = 20.1
    dv_list(ws, "E5:E8", '"None,Simplified,Delta-plus,Scenario"')


# ============================================================== calc sheets
def calc_sheet(wb: Book, name_: str, title: str, subtitle: str, cols: Sequence[tuple[str, float, str]],
               lastcol: str, formulas: Sequence[str], freeze: str = "E5") -> Sheet:
    """One formula per column; {r} stands for the row number and {p} for the row above."""
    widths = {COL(3 + k): w for k, (_, w, _) in enumerate(cols)}
    ws = new_sheet(wb, name_, title, subtitle, lastcol, "output", widths)
    header(ws, 4, "C", [lab for lab, _, _ in cols], "output", 11, 42, 2)
    name(wb, f"{name_}_Header", f"{name_}!$C$4")
    for i in range(CAP):
        rr = FIRST + i
        for k, (lab, w, nf) in enumerate(cols):
            c = ws.cell(rr, 3 + k, formulas[k].replace("{r}", str(rr)).replace("{p}", str(rr - 1)))
            c.font, c.border = font(9), BOX
            c.number_format = nf
            if nf in ("@", "0", "General") and k > 0:
                c.alignment = Alignment(horizontal="center")
    st = COL(4)
    cf_status(ws, f"C{FIRST}:{lastcol}{LAST}", st, FIRST)
    ws.freeze_panes = freeze
    return ws


def build_ir_calc(ctx: Context) -> None:
    ir_status = (
        '=IF($C{r}="","",IF(' + DUP + ',"INVALID: duplicate PositionID",'
        'IF(' + no_currency("$E{r}") + ',"INVALID: currency not in FX_Table",'
        'IF($E{r}=GoldCode,"INVALID: gold is FX risk; enter it on FXPositions",'
        'IF(' + neither("IRPositions!G{r}", "Long", "Short") + ',"INVALID: Direction must be Long or Short",'
        'IF(NOT(ISNUMBER(IRPositions!H{r})),"INVALID: Position value must be a number",'
        'IF(IRPositions!H{r}<0,"INVALID: Position value must be >= 0 (Direction gives the sign)",'
        'IF(' + neither("IRPositions!I{r}", "Fixed", "Floating") + ',"INVALID: Rate type must be Fixed or Floating",'
        'IF(NOT(ISNUMBER(IRPositions!K{r})),"INVALID: Final maturity date missing",'
        'IF(' + NOT_A_DATE.format(cell="IRPositions!K{r}") + ',"INVALID: Final maturity date is not a date",'
        'IF(IRPositions!K{r}<AsOfDate,"INVALID: final maturity before AsOfDate",'
        'IF(AND(IRPositions!I{r}="Fixed",NOT(ISNUMBER(IRPositions!J{r}))),"INVALID: Coupon required for a fixed-rate position",'
        'IF(AND(IRPositions!I{r}="Floating",NOT(ISNUMBER(IRPositions!L{r}))),"INVALID: Next repricing date required for a floating-rate position",'
        'IF(AND(IRPositions!I{r}="Floating",' + NOT_A_DATE.format(cell="IRPositions!L{r}") + '),"INVALID: Next repricing date is not a date",'
        'IF(AND(IRPositions!I{r}="Floating",OR(IRPositions!L{r}<AsOfDate,IRPositions!L{r}>IRPositions!K{r})),"INVALID: repricing date outside AsOfDate to final maturity",'
        'IF(AND($O{r}<>"GOV",$O{r}<>"QUAL",$O{r}<>"OTHER",$O{r}<>"NONE"),"INVALID: SR category must be GOV, QUAL, OTHER or NONE",'
        'IF($Q{r}="NO MATCH","INVALID: no MAR40.6 charge for this category and rating grade",'
        '"OK")))))))))))))))))')
    grd = text("IRPositions!N{r}")
    issue = text("IRPositions!O{r}")
    k1 = '$O{r}&"|"&' + grd + '&"|"&$P{r}'
    k2 = '$O{r}&"|"&' + grd + '&"|Any"'
    calc_sheet(
        ctx.wb, "IRCalc", "Interest rate position calculation",
        "Every intermediate per IRPositions row (same row number): time band, risk weight, weighted position and the "
        "specific-risk charge. Rejected rows show the reason in Status and contribute nothing.",
        [("PositionID", 15.6, "@"), ("Status", 30.6, "@"), ("Ccy", 7.6, "@"), ("FX rate", 9.6, NF_RATE),
         ("Signed value (rep ccy)", 15.6, NF_AMT), ("Slotting date", 11.6, NF_DATE), ("Residual maturity (years)", 10.6, "0.0000"),
         ("Coupon table", 8.6, "@"), ("Time band", 7.6, "0"), ("Zone", 6.6, "0"), ("Risk weight", 8.6, NF_PCT),
         ("Weighted position", 15.6, NF_AMT), ("SR category", 9.6, "@"), ("SR maturity band", 10.6, "@"),
         ("SR key", 28.6, "@"), ("SR rate", 8.6, NF_PCT), ("Issue key", 22.6, "@"), ("Issue net value", 15.6, NF_AMT),
         ("First row of issue", 9.6, "General"), ("Specific risk charge", 14.6, NF_AMT), ("Issue mixes SR keys or ccys", 11.6, "@"),
         ("Error #", 6.6, "0")],
        "X",
        ['=' + text("IRPositions!C{r}"),
         ir_status,
         '=IF($C{r}="","",UPPER(' + text("IRPositions!F{r}") + '))',
         '=IF($C{r}="","",IF(' + no_currency("$E{r}") + ',"",' + fx_rate("$E{r}") + '))',
         '=IF($D{r}="OK",IF(IRPositions!G{r}="Long",1,-1)*IRPositions!H{r}*$F{r},0)',
         '=IF($D{r}="OK",IF(IRPositions!I{r}="Floating",IRPositions!L{r},IRPositions!K{r}),"")',
         '=IF($D{r}="OK",($H{r}-AsOfDate)/DaysPerYear,"")',
         '=IF($D{r}="OK",IF(OR(IRPositions!I{r}="Floating",IRPositions!J{r}>=CouponThreshold),"High","Low"),"")',
         '=IF($D{r}="OK",IF($J{r}="High",SUMPRODUCT((Ladder_UpperHigh<>"")*(Ladder_UpperHigh<$I{r})),'
         'SUMPRODUCT((Ladder_UpperLow<>"")*(Ladder_UpperLow<$I{r})))+1,"")',
         '=IF($D{r}="OK",INDEX(Ladder_Zone,$K{r}),"")',
         '=IF($D{r}="OK",INDEX(Ladder_RW,$K{r}),"")',
         '=IF($D{r}="OK",$G{r}*$M{r},0)',
         '=IF($C{r}="","",UPPER(' + text("IRPositions!M{r}") + '))',
         '=IF($C{r}="","",IFERROR(IF((IRPositions!K{r}-AsOfDate)/DaysPerYear<=0.5,"<=6m",'
         'IF((IRPositions!K{r}-AsOfDate)/DaysPerYear<=2,"6m-24m",">24m")),""))',
         '=IF($C{r}="","",IF($O{r}="NONE","NONE",IF(' + count_where(same("SR_Key", k1)) + '>0,' + k1 +
         ',IF(' + count_where(same("SR_Key", k2)) + '>0,' + k2 + ',"NO MATCH"))))',
         '=IF(AND($D{r}="OK",$Q{r}<>"NONE"),' + sum_where("SR_Charge", same("SR_Key", "$Q{r}")) + ',0)',
         '=IF(AND($D{r}="OK",$O{r}<>"NONE"),IF(' + issue + '="","POS:"&$C{r},UPPER(' + issue + ')),"")',
         '=IF($S{r}="","",' + sum_where(R("G"), issue_rows(f"${LAST}")) + ')',
         '=IF($S{r}="",FALSE,' + count_where(issue_rows("{r}")) + '=1)',
         '=IF($U{r},ABS($T{r})*$R{r},0)',
         '=IF($S{r}="","",IF(' + count_where(issue_rows(f"${LAST}"), "(" + R("Q") + "<>$Q{r})") + '+'
         + count_where(issue_rows(f"${LAST}"), "(" + R("E") + "<>$E{r})") + '>0,"MIXED",""))',
         '=IF(LEFT($D{r},7)="INVALID",COUNTIF($D$5:$D{r},"INVALID*"),"")'])
    if not ctx.template:
        # The engine's own key: the Issue key shown for a position without an
        # Issue ID, and "ISSUE:" before an entered one. It reads the Issue ID
        # of a row only when the row has an Issue key.
        ws = ctx.wb["IRCalc"]
        ws.column_dimensions[NETTING_KEY].width = 24.6
        header(ws, 4, NETTING_KEY, ["Netting key (formula workbook only)"], "output", 11, 42, 2)
        for rr in range(FIRST, LAST + 1):
            c = ws[f"{NETTING_KEY}{rr}"]
            c.value = (f'=IF($S{rr}="","",IF(' + issue.replace("{r}", str(rr))
                       + f'="",$S{rr},"ISSUE:"&$S{rr}))')
            c.font, c.border, c.number_format = font(9), BOX, "@"
            c.alignment = Alignment(horizontal="center")


def build_eq_calc(ctx: Context) -> None:
    eq_currency = text("EQPositions!H{r}")
    eq_status = (
        '=IF($C{r}="","",IF(' + DUP + ',"INVALID: duplicate PositionID",'
        'IF($E{r}="","INVALID: Market missing",'
        'IF(AND($F{r}<>"Single name",$F{r}<>"Index"),"INVALID: Type must be Single name or Index",'
        'IF($G{r}="","INVALID: stock or index identifier missing",'
        'IF(OR(' + no_currency(eq_currency) + ',UPPER(' + eq_currency + ')=GoldCode),"INVALID: currency not in FX_Table",'
        'IF(' + neither("EQPositions!I{r}", "Long", "Short") + ',"INVALID: Direction must be Long or Short",'
        'IF(NOT(ISNUMBER(EQPositions!J{r})),"INVALID: Market value must be a number",'
        'IF(EQPositions!J{r}<0,"INVALID: Market value must be >= 0 (Direction gives the sign)",'
        '"OK")))))))))')
    calc_sheet(
        ctx.wb, "EQCalc", "Equity position calculation",
        "Per EQPositions row: signed value, net value of the same stock or index in the market, and the specific or index "
        "charge (booked on the first row of each stock or index). Market aggregation is on Buckets.",
        [("PositionID", 15.6, "@"), ("Status", 30.6, "@"), ("Market", 9.6, "@"), ("Type", 11.6, "@"),
         ("Identifier", 16.6, "@"), ("FX rate", 9.6, NF_RATE), ("Signed value (rep ccy)", 15.6, NF_AMT),
         ("Net key", 28.6, "@"), ("Net value of stock / index", 15.6, NF_AMT), ("Specific / index charge", 14.6, NF_AMT),
         ("Market #", 7.6, "0"), ("Error #", 6.6, "0")],
        "N",
        ['=' + text("EQPositions!C{r}"),
         eq_status,
         '=IF($C{r}="","",UPPER(' + text("EQPositions!E{r}") + '))',
         '=IF($C{r}="","",' + text("EQPositions!F{r}") + ')',
         '=IF($C{r}="","",UPPER(' + text("EQPositions!G{r}") + '))',
         '=IF($D{r}="OK",' + fx_rate(eq_currency) + ',"")',
         '=IF($D{r}="OK",IF(EQPositions!I{r}="Long",1,-1)*EQPositions!J{r}*$H{r},0)',
         '=IF($D{r}="OK",$E{r}&"|"&$F{r}&"|"&$G{r},"")',
         '=IF($J{r}="","",' + sum_where(R("I"), same(R("J"), "$J{r}")) + ')',
         '=IF($J{r}="",0,IF(' + count_where(same(f"$J${FIRST}:$J{{r}}", "$J{r}"))
         + '=1,IF($F{r}="Index",EQIndexRate,EQSpecificRate)*ABS($K{r}),0))',
         '=IF($D{r}<>"OK","",IF(' + FIRST_OK_ROW_OF_KEY + '=1,MAX($M$4:$M{p})+1,""))',
         '=IF(LEFT($D{r},7)="INVALID",COUNTIF($D$5:$D{r},"INVALID*"),"")'])


def build_co_calc(ctx: Context) -> None:
    co_currency = text("COPositions!J{r}")
    co_status = (
        '=IF($C{r}="","",IF(' + DUP + ',"INVALID: duplicate PositionID",'
        'IF($E{r}="","INVALID: Commodity missing",'
        'IF(' + neither("COPositions!F{r}", "Long", "Short") + ',"INVALID: Direction must be Long or Short",'
        'IF(NOT(ISNUMBER(COPositions!G{r})),"INVALID: Quantity must be a number",'
        'IF(COPositions!G{r}<0,"INVALID: Quantity must be >= 0 (Direction gives the sign)",'
        'IF(NOT(ISNUMBER(COPositions!I{r})),"INVALID: Spot price must be a number",'
        'IF(OR(' + no_currency(co_currency) + ',UPPER(' + co_currency + ')=GoldCode),"INVALID: currency not in FX_Table",'
        '"OK"))))))))')
    calc_sheet(
        ctx.wb, "COCalc", "Commodity position calculation",
        "Per COPositions row: value at spot in reporting currency. Aggregation per commodity (net and gross) is on Buckets.",
        [("PositionID", 15.6, "@"), ("Status", 30.6, "@"), ("Commodity", 18.6, "@"), ("FX rate", 9.6, NF_RATE),
         ("Signed value (rep ccy)", 15.6, NF_AMT), ("Gross value (rep ccy)", 15.6, NF_AMT), ("Commodity #", 9.6, "0"),
         ("Error #", 6.6, "0")],
        "J",
        ['=' + text("COPositions!C{r}"),
         co_status,
         '=IF($C{r}="","",UPPER(' + text("COPositions!E{r}") + '))',
         '=IF($D{r}="OK",' + fx_rate(co_currency) + ',"")',
         '=IF($D{r}="OK",IF(COPositions!F{r}="Long",1,-1)*COPositions!G{r}*COPositions!I{r}*$F{r},0)',
         '=ABS($G{r})',
         '=IF($D{r}<>"OK","",IF(' + FIRST_OK_ROW_OF_KEY + '=1,MAX($I$4:$I{p})+1,""))',
         '=IF(LEFT($D{r},7)="INVALID",COUNTIF($D$5:$D{r},"INVALID*"),"")'])


def build_fx_calc(ctx: Context) -> None:
    fx_status = (
        '=IF($C{r}="","",IF(' + DUP + ',"INVALID: duplicate PositionID",'
        'IF(' + no_currency("$E{r}") + ',"INVALID: currency not in FX_Table",'
        'IF(' + neither("FXPositions!F{r}", "Long", "Short") + ',"INVALID: Direction must be Long or Short",'
        'IF(NOT(ISNUMBER(FXPositions!G{r})),"INVALID: Amount must be a number",'
        'IF(FXPositions!G{r}<0,"INVALID: Amount must be >= 0 (Direction gives the sign)",'
        'IF($E{r}=ReportingCcy,"EXCLUDED: reporting currency carries no FX risk",'
        '"OK")))))))')
    calc_sheet(
        ctx.wb, "FXCalc", "FX position calculation",
        "Per FXPositions row: signed amount in currency and in reporting currency. The net open position is built on Buckets.",
        [("PositionID", 15.6, "@"), ("Status", 34.6, "@"), ("Ccy", 7.6, "@"), ("FX rate", 11.6, NF_RATE),
         ("Signed amount (ccy)", 17.6, NF_AMT), ("Signed value (rep ccy)", 17.6, NF_AMT), ("Error #", 6.6, "0")],
        "I",
        ['=' + text("FXPositions!C{r}"),
         fx_status,
         '=IF($C{r}="","",UPPER(' + text("FXPositions!E{r}") + '))',
         '=IF($D{r}="OK",' + fx_rate("$E{r}") + ',"")',
         '=IF($D{r}="OK",IF(FXPositions!F{r}="Long",1,-1)*FXPositions!G{r},0)',
         '=IF($D{r}="OK",$G{r}*$F{r},0)',
         '=IF(LEFT($D{r},7)="INVALID",COUNTIF($D$5:$D{r},"INVALID*"),"")'])


# =============================================================== IR_Ladder
def build_ir_ladder(ctx: Context) -> None:
    ws = new_sheet(ctx.wb, "IR_Ladder", "Interest rate maturity ladder",
                   "One block of 15 time bands per FX_Table currency (MAR40.24-40.27). Weighted positions = signed value x "
                   "risk weight; the vertical disallowance is 10% of the smaller of weighted long and weighted short in each band.",
                   "O", "output", {"C": 8.6, "D": 6.6, "E": 6.6, "F": 16.6, "G": 16.6, "H": 9.6, "I": 15.6, "J": 15.6,
                                   "K": 15.6, "L": 15.6, "M": 15.6, "N": 14.6, "O": 15.6})
    header(ws, 5, "C", ["Ccy", "Band", "Zone", "Time band, coupon >= 3%", "Time band, coupon < 3%", "Risk weight",
                        "Long (rep ccy)", "Short (rep ccy)", "Weighted long", "Weighted short", "Matched in band",
                        "Vertical disallowance", "Net weighted position"], "output", 11, 42, 1)
    name(ctx.wb, "IR_Ladder_Header", "IR_Ladder!$C$5")
    for i in range(NFX * NBANDS):
        rr = LAD_F + i
        ci, b = i // NBANDS + 1, i % NBANDS + 1
        # The rows of IRCalc in this row's currency and time band.
        slot = (same(IC("E"), f"$C{rr}"), f'({IC("K")}=$D{rr})')
        f: dict[str, str | int] = {
            "C": f'=IF(OR(INDEX(FX_Ccy,{ci})="",INDEX(FX_Ccy,{ci})=GoldCode),"",INDEX(FX_Ccy,{ci}))',
            "D": b, "E": f"=INDEX(Ladder_Zone,$D{rr})",
            "F": f'=IF(INDEX(Ladder_LabelHigh,$D{rr})="","n/a",INDEX(Ladder_LabelHigh,$D{rr}))',
            "G": f"=INDEX(Ladder_LabelLow,$D{rr})", "H": f"=INDEX(Ladder_RW,$D{rr})",
            "I": f'=IF($C{rr}="",0,{sum_where(IC("G"), *slot, positive(IC("G")))})',
            "J": f'=IF($C{rr}="",0,-{sum_where(IC("G"), *slot, negative(IC("G")))})',
            "K": f'=IF($C{rr}="",0,{sum_where(IC("N"), *slot, positive(IC("N")))})',
            "L": f'=IF($C{rr}="",0,-{sum_where(IC("N"), *slot, negative(IC("N")))})',
            "M": f"=MIN($K{rr},$L{rr})", "N": f"=VerticalDisallowance*$M{rr}", "O": f"=$K{rr}-$L{rr}",
        }
        for cl, v in f.items():
            nf = {"D": "0", "E": "0", "H": NF_PCT}.get(cl, NF_AMT if cl >= "I" else "@")
            put(ws, f"{cl}{rr}", v, nf, "calc", align="center" if cl <= "H" and cl != "C" else None)
        if b == NBANDS:
            for c in range(3, 16):
                ws.cell(rr, c).border = Border(left=THIN, right=THIN, top=THIN, bottom=Side(style="medium", color="FF" + NAVY))
    ws.conditional_formatting.add(f"C{LAD_F}:O{LAD_L}", FormulaRule(formula=[f'$C{LAD_F}=""'], font=Font(color="FFBFBFBF")))
    ws.conditional_formatting.add(f"C{LAD_F}:O{LAD_L}",
                                  FormulaRule(formula=[f'AND($C{LAD_F}<>"",OR($I{LAD_F}<>0,$J{LAD_F}<>0))'],
                                              fill=PatternFill(bgColor="FFDDEBF7")))
    ws.freeze_panes = "F6"


# =============================================================== IR_Zones
def build_ir_zones(ctx: Context) -> None:
    ws = new_sheet(ctx.wb, "IR_Zones", "Interest rate zones and charges",
                   "Per currency (MAR40.28): net band positions offset within each zone, then between zones 1-2, 2-3 and 1-3 "
                   "in that order; the residual net position is charged at 100%. No offsetting across currencies.",
                   "AF", "output",
                   {**{COL(c): 12.6 for c in range(4, 33)}, "C": 8.6, "D": 9.6})
    labels = ["Ccy", "Positions used", "Vertical disallowance",
              "Z1 net long", "Z1 net short", "Z1 matched", "Z1 charge (40%)", "Z1 net",
              "Z2 net long", "Z2 net short", "Z2 matched", "Z2 charge (30%)", "Z2 net",
              "Z3 net long", "Z3 net short", "Z3 matched", "Z3 charge (30%)", "Z3 net",
              "Matched Z1-Z2", "Z1 after", "Z2 after", "Matched Z2-Z3", "Z2 after", "Z3 after", "Matched Z1-Z3",
              "Between-zone charge", "Net position charge (100%)", "General market risk", "Specific risk",
              "IR requirement (unscaled)"]
    header(ws, 5, "C", labels, "output", 10, 54, 1)
    name(ctx.wb, "IR_Zones_Header", "IR_Zones!$C$5")
    for i in range(NFX):
        rr = ZF + i
        for cl, v in zone_formulas(i + 1, rr).items():
            put(ws, f"{cl}{rr}", v, "@" if cl == "C" else ("0" if cl == "D" else NF_AMT), "calc",
                align="center" if cl in ("C", "D") else None)
    put(ws, f"C{ZT}", "Total", None, "total")
    for c in range(4, 33):
        cl = COL(c)
        put(ws, f"{cl}{ZT}", f"=SUM({cl}{ZF}:{cl}{ZL})" if cl not in ("J", "O", "T", "V", "W", "Y", "Z") else None,
            "0" if cl == "D" else NF_AMT, "total")
    for cl in ("AD", "AE", "AF"):
        for rr in range(ZF, ZT + 1):
            ws[f"{cl}{rr}"].font = font(9, b=True)
    ws.conditional_formatting.add(f"C{ZF}:AF{ZL}", FormulaRule(formula=[f'$C{ZF}=""'], font=Font(color="FFBFBFBF")))
    ws.freeze_panes = "D6"
    name(ctx.wb, "IR_GMR_Total", f"IR_Zones!$AD${ZT}")
    name(ctx.wb, "IR_SR_Total", f"IR_Zones!$AE${ZT}")


def zone_formulas(ccy: int, rr: int) -> dict[str, str]:
    """The formulas of row rr of IR_Zones, for the currency in row ccy of FX_Table, by column."""
    currency = f"$C{rr}"
    f = {"C": f'=IF(OR(INDEX(FX_Ccy,{ccy})="",INDEX(FX_Ccy,{ccy})=GoldCode),"",INDEX(FX_Ccy,{ccy}))',
         "D": f'=IF($C{rr}="",0,{count_where(same(IC("E"), currency), ok(IC("D")))})',
         "E": f'=IF($C{rr}="",0,{sum_where(LD("N"), same(LD("C"), currency))})'}
    for z, (a, bcol, cc, d, e) in zip((1, 2, 3), [("F", "G", "H", "I", "J"), ("K", "L", "M", "N", "O"),
                                                    ("P", "Q", "R", "S", "T")]):
        in_zone = (same(LD("C"), currency), f'({LD("E")}={z})')
        f[a] = f'=IF($C{rr}="",0,{sum_where(LD("O"), *in_zone, positive(LD("O")))})'
        f[bcol] = f'=IF($C{rr}="",0,-{sum_where(LD("O"), *in_zone, negative(LD("O")))})'
        f[cc] = f"=MIN(${a}{rr},${bcol}{rr})"
        f[d] = f"=HorizontalZone{z}*${cc}{rr}"
        f[e] = f"=${a}{rr}-${bcol}{rr}"
    f.update({
        "U": f"=IF($J{rr}*$O{rr}<0,MIN(ABS($J{rr}),ABS($O{rr})),0)",
        "V": f"=$J{rr}-SIGN($J{rr})*$U{rr}", "W": f"=$O{rr}-SIGN($O{rr})*$U{rr}",
        "X": f"=IF($W{rr}*$T{rr}<0,MIN(ABS($W{rr}),ABS($T{rr})),0)",
        "Y": f"=$W{rr}-SIGN($W{rr})*$X{rr}", "Z": f"=$T{rr}-SIGN($T{rr})*$X{rr}",
        "AA": f"=IF($V{rr}*$Z{rr}<0,MIN(ABS($V{rr}),ABS($Z{rr})),0)",
        "AB": f"=HorizontalAdjacent*($U{rr}+$X{rr})+HorizontalZone13*$AA{rr}",
        "AC": f"=ABS($J{rr}+$O{rr}+$T{rr})",
        "AD": f"=$E{rr}+$I{rr}+$N{rr}+$S{rr}+$AB{rr}+$AC{rr}",
        "AE": f'=IF($C{rr}="",0,{sum_where(IC("V"), same(IC("E"), currency))})',
        "AF": f"=$AD{rr}+$AE{rr}"})
    return f


# ================================================================ Buckets
def build_buckets(ctx: Context) -> None:
    ws = new_sheet(ctx.wb, "Buckets", "Equity, commodity and FX buckets",
                   "Equity by market (MAR40.43-40.47), commodities by commodity (MAR40.72-40.73) and the FX net open "
                   "position (MAR40.60-40.61). Markets and commodities are listed in order of first valid appearance.",
                   "L", "output", {"C": 7.6, "D": 22.6, "E": 10.6, "F": 15.6, "G": 15.6, "H": 15.6, "I": 15.6,
                                   "J": 15.6, "K": 15.6, "L": 15.6})
    ctx.eq = buckets_equity(ctx.wb, ws)
    ctx.co = buckets_commodities(ctx.wb, ws, ctx.eq.total + 3)
    ctx.fx = buckets_fx(ctx.wb, ws, ctx.co.total + 3)
    ws.conditional_formatting.add(f"C{ctx.fx.first}:K{ctx.fx.last}",
                                  FormulaRule(formula=[f'$E{ctx.fx.first}="Reporting ccy"'], font=Font(color="FFBFBFBF")))
    ws.freeze_panes = "A5"


def buckets_equity(wb: Book, ws: Sheet) -> Block:
    section(ws, 5, "Equity by market", "C", "L")
    header(ws, 6, "C", ["#", "Market", "Positions", "Long (rep ccy)", "Short (rep ccy)", "Net position",
                        "General market risk (8% of |net|)", "Specific risk, single names (8%)", "Index charge (2%)", "Total"],
           "output", 11, 42, 2)
    EQ_F, EQ_L = 7, 7 + NBKT - 1
    name(wb, "Buckets_EQ_Header", "Buckets!$C$6")
    single_name, index = f'({EC("F")}="Single name")', f'({EC("F")}="Index")'
    for i in range(NBKT):
        rr = EQ_F + i
        market = same(EC("E"), f"$D{rr}")
        f: dict[str, str | int] = {
            "C": i + 1,
            "D": f'=IFERROR(INDEX({EC("E")},MATCH($C{rr},{EC("M")},0)),"")',
            "E": f'=IF($D{rr}="","",{count_where(market, ok(EC("D")))})',
            "F": f'=IF($D{rr}="","",{sum_where(EC("I"), market, positive(EC("I")))})',
            "G": f'=IF($D{rr}="","",-{sum_where(EC("I"), market, negative(EC("I")))})',
            "H": f'=IF($D{rr}="","",$F{rr}-$G{rr})',
            "I": f'=IF($D{rr}="","",EQGeneralRate*ABS($H{rr}))',
            "J": f'=IF($D{rr}="","",{sum_where(EC("L"), market, single_name)})',
            "K": f'=IF($D{rr}="","",{sum_where(EC("L"), market, index)})',
            "L": f'=IF($D{rr}="","",$I{rr}+$J{rr}+$K{rr})'}
        for cl, v in f.items():
            put(ws, f"{cl}{rr}", v, {"C": "0", "D": "@", "E": "0"}.get(cl, NF_AMT), "calc",
                align="center" if cl in ("C", "E") else None)
    EQ_T = EQ_L + 1
    put(ws, f"D{EQ_T}", "Total equity", None, "total")
    for cl in "EFGHIJKL":
        put(ws, f"{cl}{EQ_T}", f"=SUM({cl}{EQ_F}:{cl}{EQ_L})" if cl != "H" else None, "0" if cl == "E" else NF_AMT,
            "total", align="center" if cl == "E" else None)
    put(ws, f"C{EQ_T}", None, None, "total")
    name(wb, "EQ_General", f"Buckets!$I${EQ_T}")
    name(wb, "EQ_Specific", f"Buckets!$J${EQ_T}")
    name(wb, "EQ_Index", f"Buckets!$K${EQ_T}")
    return Block(EQ_F, EQ_L, EQ_T)


def buckets_commodities(wb: Book, ws: Sheet, CS: int) -> Block:
    section(ws, CS, "Commodities by commodity (simplified approach)", "C", "L")
    header(ws, CS + 1, "C", ["#", "Commodity", "Positions", "Long (rep ccy)", "Short (rep ccy)", "Net position",
                             "Gross position", "Directional (15% of |net|)", "Gross charge (3%)", "Total"],
           "output", 11, 42, 2)
    CO_F, CO_L = CS + 2, CS + 2 + NBKT - 1
    name(wb, "Buckets_CO_Header", f"Buckets!$C${CS + 1}")
    for i in range(NBKT):
        rr = CO_F + i
        commodity = same(CC("E"), f"$D{rr}")
        f: dict[str, str | int] = {
            "C": i + 1,
            "D": f'=IFERROR(INDEX({CC("E")},MATCH($C{rr},{CC("I")},0)),"")',
            "E": f'=IF($D{rr}="","",{count_where(commodity, ok(CC("D")))})',
            "F": f'=IF($D{rr}="","",{sum_where(CC("G"), commodity, positive(CC("G")))})',
            "G": f'=IF($D{rr}="","",-{sum_where(CC("G"), commodity, negative(CC("G")))})',
            "H": f'=IF($D{rr}="","",$F{rr}-$G{rr})',
            "I": f'=IF($D{rr}="","",{sum_where(CC("H"), commodity, ok(CC("D")))})',
            "J": f'=IF($D{rr}="","",COMDirectionalRate*ABS($H{rr}))',
            "K": f'=IF($D{rr}="","",COMGrossRate*$I{rr})',
            "L": f'=IF($D{rr}="","",$J{rr}+$K{rr})'}
        for cl, v in f.items():
            put(ws, f"{cl}{rr}", v, {"C": "0", "D": "@", "E": "0"}.get(cl, NF_AMT), "calc",
                align="center" if cl in ("C", "E") else None)
    CO_T = CO_L + 1
    put(ws, f"C{CO_T}", None, None, "total")
    put(ws, f"D{CO_T}", "Total commodities", None, "total")
    for cl in "EFGHIJKL":
        put(ws, f"{cl}{CO_T}", f"=SUM({cl}{CO_F}:{cl}{CO_L})" if cl != "H" else None, "0" if cl == "E" else NF_AMT,
            "total", align="center" if cl == "E" else None)
    name(wb, "CO_Directional", f"Buckets!$J${CO_T}")
    name(wb, "CO_Gross", f"Buckets!$K${CO_T}")
    return Block(CO_F, CO_L, CO_T)


def buckets_fx(wb: Book, ws: Sheet, FS: int) -> Block:
    section(ws, FS, "FX net open position (shorthand method)", "C", "L")
    header(ws, FS + 1, "C", ["#", "Currency", "Treatment", "Net amount (ccy)", "FX rate", "Net position (rep ccy)",
                             "Net long", "Net short", "Gold (|net|)"], "output", 11, 42, 2)
    FX_F, FX_L = FS + 2, FS + 2 + NFX - 1
    name(wb, "Buckets_FX_Header", f"Buckets!$C${FS + 1}")
    for i in range(NFX):
        rr = FX_F + i
        currency = same(FC("E"), f"$D{rr}")
        f: dict[str, str | int] = {
            "C": i + 1,
            "D": f'=IF(INDEX(FX_Ccy,{i + 1})="","",INDEX(FX_Ccy,{i + 1}))',
            "E": f'=IF($D{rr}="","",IF($D{rr}=ReportingCcy,"Reporting ccy",IF($D{rr}=GoldCode,"Gold","Currency")))',
            "F": f'=IF($D{rr}="",0,{sum_where(FC("G"), currency, ok(FC("D")))})',
            "G": f'=IF($D{rr}="","",INDEX(FX_Rate,{i + 1}))',
            "H": f'=IF($D{rr}="",0,{sum_where(FC("H"), currency, ok(FC("D")))})',
            "I": f'=IF($E{rr}="Currency",MAX($H{rr},0),0)',
            "J": f'=IF($E{rr}="Currency",MAX(-$H{rr},0),0)',
            "K": f'=IF($E{rr}="Gold",ABS($H{rr}),0)'}
        for cl, v in f.items():
            put(ws, f"{cl}{rr}", v, {"C": "0", "D": "@", "E": "@", "G": NF_RATE}.get(cl, NF_AMT), "calc",
                align="center" if cl in ("C", "D", "E", "G") else None)
    FX_T = FX_L + 1
    put(ws, f"C{FX_T}", None, None, "total")
    put(ws, f"D{FX_T}", "Totals", None, "total")
    for cl in "EFGH":
        put(ws, f"{cl}{FX_T}", None, None, "total")
    for cl in "IJK":
        put(ws, f"{cl}{FX_T}", f"=SUM({cl}{FX_F}:{cl}{FX_L})", NF_AMT, "total")
    summary = [("Overall net open position: max(net long, net short) + |gold|", f"=MAX(I{FX_T},J{FX_T})+K{FX_T}", "FX_ONOP"),
               ("FX requirement: 8% x overall net open position", f"=FXChargeRate*I{FX_T + 2}", "FX_Charge")]
    for k, (lab, fm, nm) in enumerate(summary):
        rr = FX_T + 2 + k
        put(ws, f"D{rr}", lab, None, None, b=True)
        put(ws, f"I{rr}", fm, NF_AMT, "total")
        name(wb, nm, f"Buckets!$I${rr}")
    return Block(FX_F, FX_L, FX_T)


# ================================================================ Results
def build_results(ctx: Context) -> None:
    wb = ctx.wb
    ws = new_sheet(wb, "Results", "Results",
                   "Own funds requirement under the simplified standardised approach: each risk class with its non-delta "
                   "option charge, multiplied by its MAR40.2 scaling factor, then summed; RWA = 12.5 x requirement (MAR40.1).",
                   "P", "output", {"C": 18.6, "D": 46.6, "E": 15.6, "F": 15.6, "G": 15.6, "H": 13.6, "I": 15.6, "J": 9.6,
                                   "K": 16.6, "L": 17.6, "M": 9.6, "N": 9.6, "O": 9.6, "P": 26.6}, 33.95)
    header(ws, 6, "C", ["Risk class", "Components", "Specific risk", "General market risk", "Other charge",
                        "Non-delta options", "Requirement before scaling", "Scaling factor", "Requirement after scaling",
                        "RWA", "Share of total", "Positions used", "Rows rejected", "Reference"], "output", 11, 42, 2)
    results_classes(ws)
    put(ws, "C11", "Total", None, "total", sz=10)
    put(ws, "D11", None, None, "total")
    for cl in "EFGHIKLNO":
        put(ws, f"{cl}11", f"=SUM({cl}7:{cl}10)", "0" if cl in "NO" else NF_AMT, "total",
            align="center" if cl in "NO" else None)
    put(ws, "J11", None, None, "total")
    put(ws, "M11", "=SUM(M7:M10)", "0.0%", "total", align="center")
    put(ws, "P11", "MAR40.1-40.2", None, "total", color=GREY_TXT)
    ws.row_dimensions[11].height = 24
    name(wb, "SSA_Requirement", "Results!$K$11")
    name(wb, "Results_Header", "Results!$C$6")
    name(wb, "Results_Status", "Results!$C$4")
    name(wb, "SSA_RWA", "Results!$L$11")
    ws["C4"].value = ('=IF(Checks!$D$4>0,"INCOMPLETE: "&Checks!$D$4&" error(s) on Checks; rejected rows are excluded '
                      'from the totals.","OK: "&$N$11&" positions used, no errors. Every figure is a live formula.")')
    ws["C4"].font = font(10, i=True, color=NOTE_TXT)
    ws.row_dimensions[4].height = 20
    ws.conditional_formatting.add("C4", FormulaRule(formula=['LEFT($C$4,10)="INCOMPLETE"'], **RED))
    section(ws, 13, "Headline", "C", "P")
    heads = [("Own funds requirement", "=SSA_Requirement", "1.30 IR + 3.50 EQ + 1.90 CO + 1.20 FX, each including non-delta option charges"),
             ("Risk-weighted assets", "=SSA_RWA", "Requirement x 12.5"),
             ("Reporting currency", "=ReportingCcy", "All amounts in this currency, unscaled units"),
             ("Reporting date", "=AsOfDate", "")]
    for k, (lab, fm, note) in enumerate(heads):
        rr = 14 + k
        put(ws, f"C{rr}", lab, None, None, b=True, sz=10)
        put(ws, f"E{rr}", fm, NF_DATE if "Date" in fm else NF_AMT, "total", align="center" if k >= 2 else None, sz=11)
        put(ws, f"F{rr}", note, None, None, color=GREY_TXT)
        ws.row_dimensions[rr].height = 22
    section(ws, 19, "Where the numbers come from", "C", "P")
    for k, (sh, txt) in enumerate([("IR_Zones", "Interest rate charges per currency, with every offset"),
                                   ("IR_Ladder", "Weighted positions per currency and time band"),
                                   ("Buckets", "Equity by market, commodities by commodity, FX net open position"),
                                   ("Checks", "Rejected rows and parameter findings")]):
        link(ws, f"C{20 + k}", sh)
        put(ws, f"D{20 + k}", txt, None, None)
    ws.freeze_panes = "D7"


def results_classes(ws: Sheet) -> None:
    """Rows 7 to 10 of Results: one row per risk class."""
    rows: list[tuple[str, str, str | int, str | int, str | int, str, str, str, str]] = [
        ("Interest rate", "Specific risk; maturity-method general market risk", "=IR_SR_Total", "=IR_GMR_Total", 0,
         "=OptionsIR", "=ScalingIR", "IRCalc", "MAR40.4-40.40"),
        ("Equity", "Specific (single names); general per market; other = index 2%", "=EQ_Specific", "=EQ_General",
         "=EQ_Index", "=OptionsEQ", "=ScalingEQ", "EQCalc", "MAR40.41-40.50"),
        ("Commodity", "General = 15% directional; other = 3% gross", 0, "=CO_Directional", "=CO_Gross",
         "=OptionsCO", "=ScalingCO", "COCalc", "MAR40.64-40.73"),
        ("Foreign exchange", "8% of the overall net open position incl. gold", 0, "=FX_Charge", 0, "=OptionsFX",
         "=ScalingFX", "FXCalc", "MAR40.53-40.61"),
    ]
    for k, (rc, comp, e, f, g, h, j, cs, ref) in enumerate(rows):
        rr = 7 + k
        put(ws, f"C{rr}", rc, None, "calc", b=True, sz=10)
        put(ws, f"D{rr}", comp, None, "calc")
        ws[f"D{rr}"].alignment = Alignment(vertical="center", wrap_text=True)
        for cl, v in (("E", e), ("F", f), ("G", g), ("H", h)):
            put(ws, f"{cl}{rr}", v, NF_AMT, "calc")
        put(ws, f"I{rr}", f"=SUM(E{rr}:H{rr})", NF_AMT, "calc", b=True)
        put(ws, f"J{rr}", j, "0.00", "calc", align="center")
        put(ws, f"K{rr}", f"=I{rr}*J{rr}", NF_AMT, "calc", b=True)
        put(ws, f"L{rr}", f"=K{rr}*RWAMultiplier", NF_AMT, "calc")
        put(ws, f"M{rr}", f"=IF($K$11=0,0,K{rr}/$K$11)", "0.0%", "calc", align="center")
        put(ws, f"N{rr}", f'=COUNTIF({cs}!$D${FIRST}:$D${LAST},"OK")', "0", "calc", align="center")
        put(ws, f"O{rr}", f'=COUNTIF({cs}!$D${FIRST}:$D${LAST},"INVALID*")', "0", "calc", align="center")
        put(ws, f"P{rr}", ref, None, "calc", color=GREY_TXT)
        ws.row_dimensions[rr].height = 30


# ================================================================= Checks
def build_checks(ctx: Context) -> None:
    wb = ctx.wb
    ws = new_sheet(wb, "Checks", "Checks",
                   "ERROR: an input row is rejected (it contributes nothing) or a parameter is unusable. WARNING: the run "
                   "is complete but departs from MAR40 or from the workbook's capacity. INFO: for the record.",
                   "J", "output", {"C": 8.6, "D": 10.6, "E": 16.6, "F": 9.6, "G": 10.6, "H": 22.6, "I": 90, "J": 4})
    ws["C4"].value = "Errors:"
    ws["C4"].font = font(10, b=True)
    ws.row_dimensions[4].height = 20
    header(ws, 6, "C", ["Rule", "Severity", "Sheet", "Count", "Result", "Reference", "Message"], "output", 11, 30, 1)
    rules = check_rules(ctx)
    for k, (rid, sev, sh, fm, ref, msg) in enumerate(rules):
        rr = 7 + k
        put(ws, f"C{rr}", rid, None, "calc", b=True)
        put(ws, f"D{rr}", sev, None, "calc", align="center")
        put(ws, f"E{rr}", sh, None, "calc")
        put(ws, f"F{rr}", fm, "0", "calc", align="center")
        put(ws, f"G{rr}", f'=IF(F{rr}=0,"OK",D{rr})', None, "calc", align="center", b=True)
        put(ws, f"H{rr}", ref, None, "calc", color=GREY_TXT)
        put(ws, f"I{rr}", msg, None, "calc")
    RULE_L = ctx.rule_last = 7 + len(rules) - 1
    ws["D4"].value = f'=SUMIF($D$7:$D${RULE_L},"ERROR",$F$7:$F${RULE_L})'
    ws["D4"].font = font(10, b=True)
    ws["E4"].value = "Warnings:"
    ws["E4"].font = font(10, b=True)
    ws["F4"].value = f'=SUMIF($D$7:$D${RULE_L},"WARNING",$F$7:$F${RULE_L})'
    ws["F4"].font = font(10, b=True)
    rngG = f"G7:G{RULE_L}"
    ws.conditional_formatting.add(rngG, CellIsRule(operator="equal", formula=['"ERROR"'], **RED))
    ws.conditional_formatting.add(rngG, CellIsRule(operator="equal", formula=['"WARNING"'], **AMBER))
    ws.conditional_formatting.add(rngG, CellIsRule(operator="equal", formula=['"OK"'], **GREEN))
    DS = ctx.findings = RULE_L + 3
    section(ws, DS, "Findings: rejected rows and unusable Params values" if ctx.template
            else "Rejected rows (first 15 per sheet)", "C", "I")
    name(wb, "Checks_Header", "Checks!$C$6")
    name(wb, "Checks_Errors", "Checks!$D$4")
    name(wb, "Checks_Warnings", "Checks!$F$4")
    name(wb, "Checks_Findings_Header", f"Checks!$C${DS + 1}")
    header(ws, DS + 1, "C", ["#", "Sheet", "PositionID / code" if ctx.template else "PositionID", "Row", "", "",
                             "Reason"], "output", 11, 30, 1)
    rr = checks_findings(ws, DS + 2)
    ws.conditional_formatting.add(f"C{DS + 2}:I{rr - 1}", FormulaRule(formula=[f'$D{DS + 2}=""'], font=Font(color="FFD9D9D9")))
    ws.freeze_panes = "A7"


def check_rules(ctx: Context) -> list[Rule]:
    """The rules of the Checks sheet, in row order; the template has its own messages and rule C14."""
    PARAM_LAST = ctx.param_last
    rules: list[Rule] = [
        ("C01", "ERROR", "IRPositions", SUMR("IRCalc"), "IRCalc Status", "Interest rate rows rejected; reason per row below and in IRCalc."),
        ("C02", "ERROR", "EQPositions", SUMR("EQCalc"), "EQCalc Status", "Equity rows rejected."),
        ("C03", "ERROR", "COPositions", SUMR("COCalc"), "COCalc Status", "Commodity rows rejected."),
        ("C04", "ERROR", "FXPositions", SUMR("FXCalc"), "FXCalc Status", "FX rows rejected."),
        ("C05", "ERROR", "Params", f'=IF({fx_rate("TRIM(ReportingCcy)")}=1,0,1)', "FX_Table",
         "The reporting currency must be listed in FX_Table with rate 1."),
        ("C06", "ERROR", "Params", "=" + "+".join(
            f'(INDEX(FX_Ccy,{i})<>"")*({count_where(same("FX_Ccy", f"INDEX(FX_Ccy,{i})"))}>1)'
            for i in range(1, NFX + 1)), "FX_Table",
         "Currency codes listed more than once in FX_Table."),
        ("C07", "ERROR", "Params", '=SUMPRODUCT((FX_Ccy<>"")*(1-ISNUMBER(FX_Rate)))+SUMPRODUCT((FX_Ccy<>"")*ISNUMBER(FX_Rate)*(FX_Rate<=0))',
         "FX_Table", "Listed currencies whose rate is missing, not a number or not positive."),
        ("C08", "WARNING", "Params", f'=COUNTIF(Params!$H$6:$H${PARAM_LAST},"CHANGED")', "MAR40",
         "Constants that differ from the MAR40 value (what-if run): the result is not the regulatory figure."),
        ("C09", "WARNING", "IRPositions", f'=COUNTIF(IRCalc!$W${FIRST}:$W${LAST},"MIXED")', "MAR40.6",
         "Rows whose Issue ID is shared with a different currency or SR key: the issue is charged at the rate of its first row."),
        ("C10", "WARNING", "Buckets", f'=MAX(0,MAX(EQCalc!$M${FIRST}:$M${LAST})-{NBKT})', "Capacity",
         f"Equity markets beyond the {NBKT} rows of Buckets: their charges are missing from Results."),
        ("C11", "WARNING", "Buckets", f'=MAX(0,MAX(COCalc!$I${FIRST}:$I${LAST})-{NBKT})', "Capacity",
         f"Commodities beyond the {NBKT} rows of Buckets: their charges are missing from Results."),
        ("C12", "INFO", "FXPositions", f'=COUNTIF(FXCalc!$D${FIRST}:$D${LAST},"EXCLUDED*")', "MAR40.53",
         "FX rows in the reporting currency, excluded (no FX risk)."),
        ("C13", "INFO", "OptionCharges", '=COUNTIF(OptionCharges!$F$5:$F$8,"<>0")', "MAR40.74-40.84",
         "Risk classes with a non-delta option charge entered by hand."),
    ]
    if ctx.template:
        rules.append(("C14", "ERROR", "Params", "=0", "MAR40",
                      "Unusable Params values (parameter, ladder, specific-risk table, option charge): the run "
                      "stops and lists them below."))
        rules[4] = rules[4][:5] + ("The reporting currency must be listed in FX_Table with rate 1 (the run stops).",)
        rules[5] = rules[5][:5] + ("Currency codes listed more than once in FX_Table (the run stops).",)
        rules[6] = rules[6][:5] + ("Listed currencies whose rate is missing, not a number or not positive (the run stops).",)
        rules[9] = rules[9][:5] + (f"Equity markets beyond the {NBKT} rows of Buckets are not listed; the totals include them.",)
        rules[10] = rules[10][:5] + (f"Commodities beyond the {NBKT} rows of Buckets are not listed; the totals include them.",)
    return rules


def checks_findings(ws: Sheet, rr: int) -> int:
    """Write the 15 finding rows of each position sheet from row rr and return the row after them."""
    for sh, cs, errcol in [("IRPositions", "IRCalc", "X"), ("EQPositions", "EQCalc", "N"),
                           ("COPositions", "COCalc", "J"), ("FXPositions", "FXCalc", "I")]:
        for k in range(1, 16):
            m = f"MATCH({k},{cs}!${errcol}${FIRST}:${errcol}${LAST},0)"
            put(ws, f"C{rr}", k, "0", "calc", align="center")
            put(ws, f"D{rr}", f'=IF(ISNA({m}),"","{sh}")', None, "calc", align="center")
            put(ws, f"E{rr}", f'=IF(ISNA({m}),"",INDEX({cs}!$C${FIRST}:$C${LAST},{m}))', None, "calc")
            put(ws, f"F{rr}", f'=IF(ISNA({m}),"",{m}+{FIRST - 1})', "0", "calc", align="center")
            for cl in "GH":
                put(ws, f"{cl}{rr}", None, None, "calc")
            put(ws, f"I{rr}", f'=IF(ISNA({m}),"",INDEX({cs}!$D${FIRST}:$D${LAST},{m}))', None, "calc")
            rr += 1
    return rr


# ========================================================== TestCatalogue
def build_test_catalogue(ctx: Context) -> None:
    wb = ctx.wb
    FX_T = ctx.fx.total
    calc = D.calc_all()
    ws = new_sheet(wb, "TestCatalogue", "Reference-test catalogue",
                   "Each test reads one output cell. T01-T03 reproduce the published MAR40.61 Table 9 example; the others "
                   "are an independent re-calculation (Python) of the shipped sample data, so they FAIL by design once the "
                   "sample positions or parameters are changed.", "N", "test",
                   {"C": 6.8, "D": 14.8, "E": 46.8, "F": 26.8, "G": 24.8, "H": 12.8, "I": 16.8, "J": 16.8, "K": 11.8,
                    "L": 9.8, "M": 9.8, "N": 4}, 33.95)
    header(ws, 6, "C", ["Test", "Area", "Description", "Source", "Source status", "Sheet", "Actual", "Expected",
                        "Abs diff", "Tolerance", "Result"], "input", 11, 30, 5)
    pub, ind = "Official published example", "Independent re-calculation"
    ir = calc["ir"]
    zrow = {c: ZF + i for i, (c, _, _) in enumerate(D.FX)}
    tests: list[tuple[str, str, str, str, str, str, float]] = [
        ("FX", "Overall net open position, Table 9 mapped to EUR", "MAR40.61 Table 9", pub, "Buckets", "=FX_ONOP", calc["fx"]["onop"]),
        ("FX", "FX requirement 8% x (300 + 35)m = 26.8m", "MAR40.61 Table 9", pub, "Buckets", "=FX_Charge", calc["fx"]["charge"]),
        ("FX", "Sum of net long positions (300m)", "MAR40.61 Table 9", pub, "Buckets", f"=Buckets!I{FX_T}", calc["fx"]["longs"]),
        ("IR", "EUR vertical disallowance", "MAR40.27", ind, "IR_Zones", f"=IR_Zones!E{zrow['EUR']}", ir["EUR"]["vd"]),
        ("IR", "EUR within-zone horizontal disallowance", "MAR40.28", ind, "IR_Zones",
         f"=IR_Zones!I{zrow['EUR']}+IR_Zones!N{zrow['EUR']}+IR_Zones!S{zrow['EUR']}", ir["EUR"]["within"]),
        ("IR", "EUR between-zone disallowance", "MAR40.28", ind, "IR_Zones", f"=IR_Zones!AB{zrow['EUR']}", ir["EUR"]["between"]),
        ("IR", "EUR net position charge", "MAR40.28", ind, "IR_Zones", f"=IR_Zones!AC{zrow['EUR']}", ir["EUR"]["net"]),
        ("IR", "USD general market risk", "MAR40.26-40.28", ind, "IR_Zones", f"=IR_Zones!AD{zrow['USD']}", ir["USD"]["gmr"]),
        ("IR", "GBP general market risk (low-coupon ladder)", "MAR40.26", ind, "IR_Zones", f"=IR_Zones!AD{zrow['GBP']}", ir["GBP"]["gmr"]),
        ("IR", "Specific risk, all currencies (issue netting)", "MAR40.6", ind, "IR_Zones", "=IR_SR_Total",
         sum(v["sr"] for v in ir.values())),
        ("EQ", "Specific risk, single names", "MAR40.43, 40.46", ind, "Buckets", "=EQ_Specific", calc["eq"]["specific"]),
        ("EQ", "General market risk per market", "MAR40.43", ind, "Buckets", "=EQ_General", calc["eq"]["general"]),
        ("EQ", "Index charge", "MAR40.47", ind, "Buckets", "=EQ_Index", calc["eq"]["index"]),
        ("CO", "Directional charge", "MAR40.72", ind, "Buckets", "=CO_Directional", calc["co"]["directional"]),
        ("CO", "Gross charge", "MAR40.73", ind, "Buckets", "=CO_Gross", calc["co"]["gross"]),
        ("Total", "Own funds requirement after scaling", "MAR40.2", ind, "Results", "=SSA_Requirement", calc["total"]),
        ("Total", "RWA", "MAR40.1", ind, "Results", "=SSA_RWA", calc["rwa"]),
        ("Invariant", "Ladder weighted positions equal IRCalc weighted positions (difference = 0)", "Workbook", "Invariant",
         "IR_Ladder", f'=SUM(IR_Ladder!K{LAD_F}:K{LAD_L})-SUM(IR_Ladder!L{LAD_F}:L{LAD_L})-SUM({IC("N")})', 0),
        ("Invariant", "Results total equals the MAR40.2 formula (difference = 0)", "MAR40.2", "Invariant", "Results",
         "=SSA_Requirement-(ScalingIR*Results!I7+ScalingEQ*Results!I8+ScalingCO*Results!I9+ScalingFX*Results!I10)", 0),
    ]
    for k, (area, desc, src, sst, sh, fm, exp) in enumerate(tests):
        rr = 7 + k
        put(ws, f"C{rr}", f"T{k + 1:02d}", None, None, b=True, sz=10)
        put(ws, f"D{rr}", area, None, None, sz=10)
        put(ws, f"E{rr}", desc, None, None, sz=10)
        put(ws, f"F{rr}", src, None, None, sz=10)
        put(ws, f"G{rr}", sst, None, None, sz=10)
        link(ws, f"H{rr}", sh)
        put(ws, f"I{rr}", fm, "#,##0.00", "calc")
        put(ws, f"J{rr}", round(exp, 6), NF_AMT, "const")
        put(ws, f"K{rr}", f"=ABS(I{rr}-J{rr})", "0.0000", "calc")
        put(ws, f"L{rr}", 0.01, "0.00", "const", align="center")
        put(ws, f"M{rr}", f'=IF(K{rr}<=L{rr},"PASS","FAIL")', None, "calc", b=True, align="center")
        ws.row_dimensions[rr].height = 20.1
    TL = 7 + len(tests) - 1
    name(wb, "Tests_Result", f"TestCatalogue!$M$7:$M${TL}")
    ws["C4"].value = f'=COUNTIF(M7:M{TL},"PASS")&" of "&COUNTA(M7:M{TL})&" tests PASS"'
    ws["C4"].font = font(10, b=True, color=NOTE_TXT)
    ws.row_dimensions[4].height = 20
    ws.conditional_formatting.add(f"M7:M{TL}", CellIsRule(operator="equal", formula=['"PASS"'], **GREEN))
    ws.conditional_formatting.add(f"M7:M{TL}", CellIsRule(operator="equal", formula=['"FAIL"'], **RED))
    ws.freeze_panes = "D7"


# ================================================================= README
def build_readme(ctx: Context) -> None:
    TEMPLATE = ctx.template
    ws = new_sheet(ctx.wb, "README", "ReadMe",
                   "Basel MAR40 simplified standardised approach for market risk (the recalibrated Basel II standardised "
                   "approach), " + ("with a VBA engine (RunSSA)." if TEMPLATE
                    else "as a macro-free prototype."), "E", "home", {"C": 40, "D": 150, "E": 4.8})
    r = 5
    section(ws, r, "How to use", "C", "D")
    r = text_block(ws, r + 1, [
        ("1. Parameters", "Params: reporting date, reporting currency and FX rates (yellow). The grey cells are MAR40 constants; "
         "changing one makes the run a what-if, flagged on Checks."),
        ("2. Positions", "IRPositions, EQPositions, COPositions and FXPositions: one row per position, up to 500 per sheet. "
         "Values are positive; Direction gives the sign. Derivatives enter as their delta-equivalent legs."),
        ("3. Options", "OptionCharges: non-delta (gamma, vega or simplified-approach) charges per risk class, calculated outside "
         "the workbook."),
        ("4. Run" if TEMPLATE else "4. Read",
         ("Import the modules of src/ (README), save as .xlsm, then Alt+F8 > RunSSA. The outputs are values written by the "
          "engine; rerun after any change. A Params value the engine cannot use stops the run and is listed on Checks." if TEMPLATE
          else "No Run button: every output is a live formula.") +
         " Results gives the requirement and RWA; Checks lists every rejected row and parameter finding; TestCatalogue "
         "shows the reference tests."),
        ("5. Drill down", "IRCalc, EQCalc, COCalc and FXCalc repeat each input row with every intermediate. IR_Ladder and IR_Zones "
         "show the maturity method step by step; Buckets aggregates equities, commodities and FX."),
    ], True)
    r += 1
    section(ws, r, "Methodology", "C", "D")
    r = text_block(ws, r + 1, [
        ("Aggregation (MAR40.1-40.2)", "Requirement = 1.30 x IR + 3.50 x EQ + 1.90 x CO + 1.20 x FX, each class including its "
         "non-delta option charge. RWA = 12.5 x requirement."),
        ("IR specific risk (MAR40.6)", "Charge = |net value per issue| x Table 1 rate, by category (GOV, QUAL, OTHER), rating "
         "grade and residual maturity to final maturity. Long and short positions in the same Issue ID net first. "
         "Derivative legs (NONE) carry no specific risk. Securitisation positions are out of scope."),
        ("IR general market risk (MAR40.24-40.28)", "Maturity method, one ladder per currency. Fixed-rate positions are slotted "
         "by residual maturity, floating-rate positions by next repricing date; coupons below 3% use the low-coupon bands. "
         "Weighted position = value x risk weight. Charge = 10% of the matched weighted position in each band + 40% / 30% / 30% "
         "of the matched band positions within zones 1 / 2 / 3 + 40% of the zone 1-2 and zone 2-3 offsets + 100% of the zone "
         "1-3 offset + 100% of the residual net position. No offsetting across currencies."),
        ("Equity (MAR40.41-40.50)", "Positions in the same stock (or index) in a market net. Specific risk = 8% x |net| per "
         "single name; index contracts carry a further 2% x |net| per index; general market risk = 8% x |net| per market, "
         "single names and indices together."),
        ("Commodities (MAR40.72-40.73)", "Simplified approach: 15% x |net position| per commodity + 3% x gross position, at "
         "spot prices converted to the reporting currency. Each distinct commodity name is a separate commodity."),
        ("FX and gold (MAR40.53-40.61)", "Shorthand method: 8% x (max(sum of net long, sum of net short) + |net gold|). "
         "Positions in the reporting currency are excluded."),
    ], True)
    r += 1
    section(ws, r, "Conventions", "C", "D")
    r = text_block(ws, r + 1, [
        ("Colours", "Yellow fill = input. Grey fill = MAR40 constant or expected test value. No fill = formula. Navy headers = "
         "inputs and reference tables; blue headers = outputs."),
        ("Amounts", "Reporting currency, unscaled units. FX rates are units of reporting currency per unit of currency."),
        ("Dates", "Residual maturity = (date - AsOfDate) / DaysPerYear. Time-band upper bounds are inclusive."),
        ("Row status", "OK = used; INVALID: reason = rejected, contributes nothing, listed on Checks; EXCLUDED = deliberately "
         "out of scope (FX in reporting currency)."),
        ("Names", "Every parameter and table is a workbook name (Formulas > Name Manager)."),
    ], True)
    r += 1
    section(ws, r, "Scope and limitations", "C", "D")
    r = text_block(ws, r + 1, [
        ("Not implemented", "Duration method for IR general market risk; commodity maturity-ladder approach; delta-plus and "
         "scenario option methods (enter their non-delta charges on OptionCharges); securitisation and nth-to-default "
         "specific risk; index-arbitrage carve-outs (MAR40.48-40.50); the 4% specific-risk option for liquid diversified "
         "portfolios; the FX de minimis exemption (MAR40.62); structural FX position exclusions."),
        ("Simplifications", "Positions are entered already decomposed into legs. Specific-risk hedges other than identical "
         "issues are not recognised. Immaterial-currency aggregation (MAR40.24) is not offered."),
        ("Regime", "Basel MAR40 only. EU CRR3 provisions on the simplified standardised approach, and their application date, "
         "are not modelled and should be checked against the Official Journal before use."),
        ("Status", "Prototype: not validated for regulatory reporting. Sample data are illustrative."),
    ], True)
    r += 1
    section(ws, r, "References", "C", "D")
    r = text_block(ws, r + 1, [
        ("MAR40", "Basel Committee on Banking Supervision, Basel Framework, MAR40 Simplified standardised approach "
         "(bis.org/basel_framework/chapter/MAR/40.htm)."),
        ("Source", f"{REPOSITORY} on GitHub: the exported VBA source, the guides and the tools that build this workbook."),
    ], True)
    r += 1
    # The build stamp: tools/devbuild/build_workbook.py writes the commit, its
    # time and the configuration into these cells, which it finds by name.
    section(ws, r, "Build", "C", "D")
    text_block(ws, r + 1, [(label, "") for label, _ in STAMP], True)
    for k, (_, stamp_name) in enumerate(STAMP, r + 1):
        ws.cell(k, 4).value = None
        name(ctx.wb, stamp_name, f"README!$D${k}")


# ================================================================= Home (fill)
def fill_home(ctx: Context) -> None:
    TEMPLATE = ctx.template
    home = ctx.wb["Home"]
    home["C3"].value = (f"Start here: what {PRODUCT} calculates, how to use it and a link to every sheet. " +
                        ("The VBA engine writes the outputs: import the modules, then run RunSSA." if TEMPLATE
                         else "No macros: every figure is a live formula."))
    r = 5
    section(home, r, f"What {PRODUCT} does", "C", "D")
    r = text_block(home, r + 1, [
        ("Calculates", "The market risk own funds requirement and RWA under the Basel MAR40 simplified standardised approach: "
         "interest rate (specific risk and maturity-method general market risk), equity, commodity (simplified approach) and "
         "FX including gold, scaled by 1.30 / 3.50 / 1.90 / 1.20."),
        ("Does not cover", "Duration method, commodity maturity ladder, option delta-plus and scenario methods (non-delta "
         "charges are an input), securitisation specific risk, the full FRTB sensitivities-based method. README gives the list."),
        ("Status", "Prototype. Sample data reproduce the MAR40 Table 9 FX example and are otherwise illustrative."),
    ])
    r += 1
    section(home, r, "How to use", "C", "D")
    r = text_block(home, r + 1, [
        ("1. Check Params", "Set the reporting date, the reporting currency and the FX rates."),
        ("2. Enter the positions", "One row per position on IRPositions, EQPositions, COPositions and FXPositions; non-delta "
         "option charges on OptionCharges."),
        ("3. Run" if TEMPLATE else "3. Read the results",
         "Alt+F8 > RunSSA calculates and writes every output sheet (ClearOutputs empties them). If the Results status "
         "line says INCOMPLETE or STOPPED, Checks lists every rejected row and unusable parameter." if TEMPLATE else
         "Results updates live. If its status line says INCOMPLETE, Checks lists every rejected row with its reason."),
        ("4. Drill down", "IRCalc / EQCalc / COCalc / FXCalc per row; IR_Ladder and IR_Zones for the maturity method; "
         "Buckets per market, commodity and currency."),
    ])
    r += 1
    section(home, r, "Sheets", "C", "D")
    sheet_desc = [
        ("Params", "Inputs. Reporting date, currency, MAR40 constants, specific-risk table, time bands and FX rates."),
        ("IRPositions", "Inputs. Debt positions and interest rate derivative legs."),
        ("EQPositions", "Inputs. Equity and index positions."),
        ("COPositions", "Inputs. Commodity positions."),
        ("FXPositions", "Inputs. Net currency and gold positions."),
        ("OptionCharges", "Inputs. Non-delta option charges per risk class."),
        ("Results", "Outputs. Requirement per risk class, scaling, total and RWA."),
        ("IR_Zones", "Outputs. Interest rate charges per currency, every offset shown."),
        ("IR_Ladder", "Outputs. Weighted positions per currency and time band."),
        ("Buckets", "Outputs. Equity by market, commodities by commodity, FX net open position."),
        ("IRCalc", "Outputs. Every intermediate per interest rate row."),
        ("EQCalc", "Outputs. Every intermediate per equity row."),
        ("COCalc", "Outputs. Every intermediate per commodity row."),
        ("FXCalc", "Outputs. Every intermediate per FX row."),
        ("Checks", "Outputs. Errors, warnings and information lines; rejected rows with reasons."),
        ("TestCatalogue", "Reference and tests. Published example and independent re-calculation."),
        ("README", "Reference. Methodology, conventions, scope and limitations."),
    ]
    for k, (sh, txt) in enumerate(sheet_desc):
        rr = r + 1 + k
        link(home, f"C{rr}", sh)
        c = home.cell(rr, 4, txt)
        c.font = font(9)
        bd = Border(bottom=THIN if k == len(sheet_desc) - 1 else Side(style="dashed", color="FFBFBFBF"))
        home.cell(rr, 3).border = c.border = bd
        home.row_dimensions[rr].height = 18
    home.freeze_panes = "A4"


# ============================================================== workbook
def finish_workbook(wb: Book) -> None:
    """Sheet order, calculation flag, package properties and sheet code names."""
    wb._sheets = [wb[n] for n in SHEET_ORDER]
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    # No personal metadata in the package (repository binary rules).
    wb.properties.creator = ""
    wb.properties.lastModifiedBy = ""
    # Code names "sh" + tab name in PascalCase (house style, protection matrix).
    for ws_ in wb.worksheets:
        ws_.sheet_properties.codeName = "sh" + "".join(
            part[:1].upper() + part[1:]
            for part in ws_.title.split("_"))


def blank_outputs(ctx: Context) -> None:
    """Template only. Output areas are written by RunSSA: no formulas or constants in them."""
    wb = ctx.wb
    for sh, width in (("IRCalc", 22), ("EQCalc", 12), ("COCalc", 8), ("FXCalc", 7)):
        blank(wb[sh], range(FIRST, LAST + 1), range(3, 3 + width))
    blank(wb["IR_Ladder"], range(LAD_F, LAD_L + 1), range(3, 16))
    blank(wb["IR_Zones"], range(ZF, ZT + 1), range(3, 33), keep=(f"C{ZT}",))
    buckets = wb["Buckets"]
    blank(buckets, range(ctx.eq.first, ctx.eq.total + 1), range(3, 13), keep=(f"D{ctx.eq.total}",))
    blank(buckets, range(ctx.co.first, ctx.co.total + 1), range(3, 13), keep=(f"D{ctx.co.total}",))
    blank(buckets, range(ctx.fx.first, ctx.fx.total + 1), range(3, 12), keep=(f"D{ctx.fx.total}",))
    blank(buckets, range(ctx.fx.total + 2, ctx.fx.total + 4), [9])
    results = wb["Results"]
    blank(results, range(7, 12), range(5, 16))
    results["C4"].value = "Not calculated: run RunSSA."
    checks = wb["Checks"]
    blank(checks, range(7, ctx.rule_last + 1), [6, 7])
    blank(checks, [4], [4, 6])
    blank(checks, range(ctx.findings + 2, ctx.findings + 62), range(3, 10))
    checks.conditional_formatting.add(f"G7:G{ctx.rule_last}", CellIsRule(operator="equal", formula=['"NOT RUN"'],
                                      font=Font(color="FF7F7F7F")))


def fix_package_times(path: Path) -> None:
    """Rewrite the package with fixed timestamps, so that two builds are byte-identical.

    openpyxl stamps the properties with the time of the save, and the zip
    entries with the clock."""
    with zipfile.ZipFile(path) as package:
        parts = [(info.filename, package.read(info.filename)) for info in package.infolist()]
    with zipfile.ZipFile(path, "w") as package:
        for filename, data in parts:
            if filename == "docProps/core.xml":
                data = CORE_TIME.sub(rb"\g<1>" + PACKAGE_TIME.encode("ascii") + rb"\g<3>", data)
            info = zipfile.ZipInfo(filename, date_time=ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            # ZipInfo takes the creating system from sys.platform; fix it, so
            # that the package is the same wherever it is built.
            info.create_system = 3
            info.external_attr = 0o644 << 16
            package.writestr(info, data)


def build(output: Path, template: bool) -> None:
    """Build the workbook and save it to output: the template if template is set, else the formula prototype."""
    wb = Workbook()
    wb.remove(wb.active)
    ctx = Context(wb, template)
    # The order of the calls is the order in which openpyxl numbers styles,
    # strings and names: keep it.
    build_home_frame(ctx)
    build_params(ctx)
    build_ir_positions(ctx)
    build_eq_positions(ctx)
    build_co_positions(ctx)
    build_fx_positions(ctx)
    build_option_charges(ctx)
    build_ir_calc(ctx)
    build_eq_calc(ctx)
    build_co_calc(ctx)
    build_fx_calc(ctx)
    build_ir_ladder(ctx)
    build_ir_zones(ctx)
    build_buckets(ctx)
    build_results(ctx)
    build_checks(ctx)
    build_test_catalogue(ctx)
    build_readme(ctx)
    fill_home(ctx)
    finish_workbook(wb)
    if template:
        blank_outputs(ctx)
    wb.save(output)
    fix_package_times(output)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Build the K-FRTB-SSA workbook: the formula prototype, or with --template the "
                    "macro-free template for the VBA engine.")
    parser.add_argument("output", nargs="?", type=Path, default=Path(DEFAULT_OUTPUT),
                        help=f"workbook to write (default: {DEFAULT_OUTPUT})")
    parser.add_argument("--template", action="store_true",
                        help="write the template (output areas empty) instead of the formula prototype")
    args = parser.parse_args(argv)
    build(args.output, args.template)
    print("saved", args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
