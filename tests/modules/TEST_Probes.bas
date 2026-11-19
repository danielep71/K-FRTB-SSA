Attribute VB_Name = "TEST_Probes"
'==============================================================================
' MODULE: TEST_Probes
'------------------------------------------------------------------------------
' PURPOSE
'   The probes suite of the workbook's tests: what the prototype relies on
'   Excel to do, as cases and assertions of TEST_Harness. It reads an error
'   value from an input cell, writes a text into a cell whose number format
'   is not text, takes a range as a table, shows the status line on the
'   other output sheets, checks the save state after the status check and
'   checks that the template protects its sheets as the workbook relies on.
'   Each case is the probe of a behaviour of the Host behaviour table of
'   docs/methodology/VALIDATION_PLAN.md. A pass measures the behaviour on
'   the host that ran it; it validates no rule of the calculation.
'
' PUBLIC SURFACE
'   RunProbeTests is the documented test entry point. Test_RunProbes runs
'   the same suite and returns its result line where a caller cannot read
'   an error, as tools/devbuild/lo_macro.py. CertifyProbes certifies the run
'   and saves its log and its certification log as evidence of an Excel
'   run. Option Private Module keeps them out of the workbook's external
'   surface. Never part of the production import set.
'
' DEPENDENCIES
'   TEST_Harness; M_Main, CORE_Config and CORE_Util; the workbook's names
'   and sheets, with the sample data the template ships.
'
' STATE OWNERSHIP
'   Owns the rows it appends under the four position tables, which it
'   removes at the end of each case that appends them, and the rows each
'   table had before. Changes ScalingIR and ThisWorkbook.Saved for one
'   check and puts them back. Rewrites the output sheets through
'   RunSSA_Silent, and the last case leaves them cleared.
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START. An unexpected error inside a case is
'   recorded as a failure of that case, and the run continues.
'   Test_RunProbes contains those errors; CertifyProbes records them, saves
'   the log, and raises only when the certification did not pass.
'
' USAGE
'   In a development workbook, type TEST_Probes.RunProbeTests in the
'   Immediate window. To record an Excel run, type
'   TEST_Probes.CertifyProbes "<folder>", "<checkout>\tests\fixtures"
'   (INSTALLATION.md, Recording an Excel run).
'
' COMPATIBILITY
'   Excel VBA. LibreOffice's VBA mode runs it, and differs from Excel on the
'   probes of an error value in a cell: tools/devbuild/lo_smoke.py lists the
'   assertions it is known to fail there. That is never Excel evidence.
'
' REFERENCE
'   The expected texts and numbers of error-values-in-inputs are those of
'   error_keys and error_values of tools/prototype/verify.py, which
'   checks the formula workbook of the same rows.
'
' UPDATED
'   2026-11-19
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the tests inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once.
        Private Const SUITE_NAME          As String = "probes"
        Private Const EXPECTED_CASES      As Long = 7
        Private Const EXPECTED_ASSERTIONS As Long = 49
    'The error values a cell holds: the codes of CVErr.
        Private Const ERR_NA              As Long = 2042   '#N/A
        Private Const ERR_VALUE           As Long = 2015   '#VALUE!
        Private Const ERR_REF             As Long = 2023   '#REF!
    'The error Excel raises for a write to a locked cell of a protected sheet.
        Private Const ERR_PROTECTED_WRITE As Long = 1004
    'The text that stands for an error value in an output.
        Private Const ERROR_TEXT          As String = "#ERROR"

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mOriginal(1 To 4) As Long   'Rows under the header of a table before the case
        Private mFirst(1 To 4)    As Long   'Row, under the header, of the first row appended
        Private mAppended(1 To 4) As Long   'Rows appended to the table



'------------------------------------------------------------------------------
'                                 ENTRY POINTS
'------------------------------------------------------------------------------

Public Sub RunProbeTests()
'
'==============================================================================
'                                 RunProbeTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run the probes suite. A passing run ends with
'   RESULT=PASS; suite=probes; operation=PASS; completeness=COMPLETE;
'   cases=7; assertions=49; failures=0; cleanup=PASS; recovery=NONE.
'
' STATE OWNERSHIP
'   Rewrites the output sheets and leaves them cleared.
'
' ERROR POLICY
'   As TEST_Harness.EndRun: raises when the run does not pass.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
        RunSuite

End Sub


Public Function Test_RunProbes() As String
'
'==============================================================================
'                                Test_RunProbes
'------------------------------------------------------------------------------
' PURPOSE
'   Run the probes suite for a caller that reads a result and not an error.
'
' RETURNS
'   The RESULT line of the run, as TEST_Harness.LastResult gives it.
'
' STATE OWNERSHIP
'   Rewrites the output sheets and leaves them cleared.
'
' ERROR POLICY
'   Contains the error of a run that does not pass, or of a refused start:
'   the result line says so.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
        On Error Resume Next
        RunSuite
        On Error GoTo 0
        Test_RunProbes = TEST_Harness.LastResult()

End Function


Public Sub CertifyProbes( _
    ByVal Folder As String, _
    ByVal Fixtures As String)
'
'==============================================================================
'                                 CertifyProbes
'------------------------------------------------------------------------------
' PURPOSE
'   Certify the probes suite's run as evidence of an Excel run: run it, save
'   its log as probes.log, and write probes-certification.log.
'
' INPUTS
'   Folder: an existing folder for the logs, without a trailing separator.
'   Fixtures: the full path of tests/fixtures/ in the checkout; not read,
'   and taken so that every certification entry point has the same two
'   arguments.
'
' STATE OWNERSHIP
'   Rewrites the output sheets and leaves them cleared.
'
' ERROR POLICY
'   The error of the run is recorded, not raised, so that the log is always
'   saved. Raises ERR_TEST_FAILED after writing the certification log when
'   the certification did not pass. An error of BeginCertification or
'   SaveLog propagates; run TEST_Harness.ResetTests before trying again.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim RunErrorNumber        As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription   As String   'Its description

        TEST_Harness.BeginCertification SUITE_NAME
        On Error Resume Next
        RunSuite
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
        On Error GoTo 0
        TEST_Harness.SaveLog Folder & "\probes.log"
        TEST_Harness.Echo "SAVED=" & Folder & "\probes.log; mode=NORMAL; error=" & _
            CStr(RunErrorNumber) & IIf(RunErrorNumber = 0, vbNullString, " " & RunErrorDescription)
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
        TEST_Harness.EndCertification Folder & "\probes-certification.log"

End Sub


'------------------------------------------------------------------------------
'                                    RUNNER
'------------------------------------------------------------------------------

Private Sub RunSuite()
'
'==============================================================================
'                                   RunSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. The order
'   matters: the last case clears the output sheets.
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it
'   and re-raises it unchanged. BeginRun's refusal propagates directly.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim ErrNumber        As Long     'Runner error number; 0 for none
    Dim ErrSource        As String   'Runner error source
    Dim ErrDescription   As String   'Runner error description

        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS
        On Error GoTo RunFailed
        AllowWritingToSheets
        CaseErrorValuesInInputs
        CaseReadingAnErrorValue
        CaseTextWrittenIntoACell
        CaseAreasAsTables
        CaseSheetProtection
        CaseStatusOnTheOtherSheets
        CaseSaveStateAfterTheCheck

Finish:
        On Error GoTo 0
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription
        Exit Sub

RunFailed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume Finish

End Sub


'------------------------------------------------------------------------------
'                                     CASES
'------------------------------------------------------------------------------

Private Sub CaseErrorValuesInInputs()
'
'==============================================================================
'                            CaseErrorValuesInInputs
'------------------------------------------------------------------------------
' PURPOSE
'   An error value typed in an input cell, as an Issue ID, a currency, a
'   PositionID or a market, a commodity and an amount code, reads as the
'   text #ERROR where an output shows that text, and no output cell holds
'   an error value. HB-08.
'
' REFERENCE
'   Rows of error_keys of tools/prototype/verify.py. Two Issue IDs that
'   both read #ERROR are one issue: its key is #ERROR and the net of its
'   long 3,000,000 and short 1,000,000 is 2,000,000.
'
' STATE OWNERSHIP
'   Appends rows under the position tables and removes them at the end.
'
' RETURNS
'   Nothing; the run's own result is not compared.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim ResultText   As String    'RunSSA_Silent result line, not compared
    Dim RestoreDue   As Boolean   'True once rows are appended
    Dim Detail       As String    'Why the appended rows could not be removed

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "error-values-in-inputs"
        RestoreDue = True
        AppendRows 1, 3
        AppendRows 2, 3
        AppendRows 3, 2
        AppendRows 4, 1
        PutRow 1, 0, Array("IR-K1", "T", "Bond", "EUR", "Long", 3000000, "Fixed", 0.05, _
            DateSerial(2030, 1, 1), Empty, "OTHER", "BB+ to BB-", CVErr(ERR_NA))
        PutRow 1, 1, Array("IR-K2", "T", "Bond", "EUR", "Short", 1000000, "Fixed", 0.05, _
            DateSerial(2030, 1, 1), Empty, "OTHER", "BB+ to BB-", CVErr(ERR_VALUE))
        PutRow 1, 2, Array("IR-K3", "T", "Bond", CVErr(ERR_NA), "Long", 1000000, "Fixed", 0.05, _
            DateSerial(2030, 1, 1), Empty, "NONE")
        PutRow 2, 0, Array(CVErr(ERR_NA), "T", "IT", "Single name", "ENI", "EUR", "Short", 1000000)
        PutRow 2, 1, Array("EQ-K2", "T", "IT", "Single name", CVErr(ERR_REF), "EUR", "Long", 2000000)
        PutRow 2, 2, Array("EQ-K3", "T", CVErr(ERR_NA), "Index", "DAX", "EUR", "Long", 1000000)
        PutRow 3, 0, Array("CO-K1", "T", CVErr(ERR_NA), "Long", 100, "t", 250, "EUR")
        PutRow 3, 1, Array("CO-K2", "T", "Wheat", "Long", 100, "t", 250, CVErr(ERR_NA))
        PutRow 4, 0, Array("FX-K1", "FX", CVErr(ERR_NA), "Long", 1000000)
        ResultText = RunSSA_Silent()
        TEST_Harness.AssertEqualText "issue-key-of-an-issue-id-typed-as-an-error", _
            ERROR_TEXT, PlainText(RowCell(NM_IR_CALC_HEADER, mFirst(1), 16))
        TEST_Harness.AssertEqualText "issue-key-of-another-issue-id-typed-as-an-error", _
            ERROR_TEXT, PlainText(RowCell(NM_IR_CALC_HEADER, mFirst(1) + 1, 16))
        TEST_Harness.AssertEqualDouble "net-of-the-issue-of-error-ids", 2000000, _
            NumberOf(RowCell(NM_IR_CALC_HEADER, mFirst(1), 17))
        TEST_Harness.AssertEqualDouble "net-of-the-issue-on-its-second-row", 2000000, _
            NumberOf(RowCell(NM_IR_CALC_HEADER, mFirst(1) + 1, 17))
        TEST_Harness.AssertEqualText "currency-typed-as-an-error", _
            ERROR_TEXT, PlainText(RowCell(NM_IR_CALC_HEADER, mFirst(1) + 2, 2))
        TEST_Harness.AssertEqualText "status-of-a-currency-typed-as-an-error", _
            "INVALID: currency not in FX_Table", _
            PlainText(RowCell(NM_IR_CALC_HEADER, mFirst(1) + 2, 1))
        TEST_Harness.AssertEqualText "position-id-typed-as-an-error", _
            ERROR_TEXT, PlainText(RowCell(NM_EQ_CALC_HEADER, mFirst(2), 0))
        TEST_Harness.AssertEqualDouble "output-cells-that-hold-an-error-value", 0, ErrorCells()

CleanUp:
        On Error GoTo 0
        If RestoreDue Then
            Detail = RemoveAppended()
        End If
        TEST_Harness.AssertTrue "appended-rows-removed", Len(Detail) = 0, Detail
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub


Private Sub CaseReadingAnErrorValue()
'
'==============================================================================
'                            CaseReadingAnErrorValue
'------------------------------------------------------------------------------
' PURPOSE
'   A cell given an error value holds one: its Value is an error, which
'   PlainText reads as the text #ERROR, while the text #N/A written into a
'   cell that is formatted as text stays a text. HB-08.
'
' STATE OWNERSHIP
'   Appends rows under the position tables and removes them at the end.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim Target       As Range     'Scratch cell: the amount of an appended row
    Dim RestoreDue   As Boolean   'True once rows are appended
    Dim Detail       As String    'Why the appended rows could not be removed

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "reading-an-error-value"
        RestoreDue = True
        AppendRows 4, 1
        Set Target = PositionHeader(4).Offset(mFirst(4), 4)
        Target.Value = CVErr(ERR_NA)
        TEST_Harness.AssertTrue "error-value-is-read-as-an-error", IsError(Target.Value), _
            "a cell given CVErr(#N/A) must hold an error value"
        TEST_Harness.AssertEqualText "error-value-reads-as-the-error-text", _
            ERROR_TEXT, PlainText(Target.Value)
        Set Target = PositionHeader(4).Offset(mFirst(4), 0)
        Target.Value = "#N/A"
        TEST_Harness.AssertTrue "text-in-a-text-cell-stays-a-text", _
            Not IsError(Target.Value) And VarType(Target.Value) = vbString, _
            "#N/A written into a text-formatted cell must stay a text"

CleanUp:
        On Error GoTo 0
        If RestoreDue Then
            Detail = RemoveAppended()
        End If
        TEST_Harness.AssertTrue "appended-rows-removed", Len(Detail) = 0, Detail
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub


Private Sub CaseTextWrittenIntoACell()
'
'==============================================================================
'                           CaseTextWrittenIntoACell
'------------------------------------------------------------------------------
' PURPOSE
'   A text that reads as a number, written as it is into a cell that is not
'   formatted as text, becomes a number; written with the cell formatted as
'   text, and the cell given its format back, it stays a text with its
'   digits. This is how the cases suite writes a text. HB-09.
'
' STATE OWNERSHIP
'   Appends rows under the position tables and removes them at the end.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim Target          As Range     'Scratch cell: the amount of an appended row
    Dim Saved           As String    'Number format of that cell
    Dim FormatChanged   As Boolean   'True while the cell is formatted as text
    Dim RestoreDue      As Boolean   'True once rows are appended
    Dim Detail          As String    'Why the appended rows could not be removed

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "text-written-into-a-cell"
        RestoreDue = True
        AppendRows 4, 1
        Set Target = PositionHeader(4).Offset(mFirst(4), 4)
        Saved = CStr(Target.NumberFormat)
        TEST_Harness.AssertTrue "amount-cell-is-not-text-formatted", Saved <> "@", _
            "the scratch cell must not be formatted as text; it is " & Saved
        Target.Value = "00123"
        TEST_Harness.AssertTrue "text-written-directly-becomes-a-number", _
            VarType(Target.Value) = vbDouble, _
            "00123 written as it is must become a number"
        Target.ClearContents
        FormatChanged = True
        Target.NumberFormat = "@"
        Target.Value = "00123"
        Target.NumberFormat = Saved
        FormatChanged = False
        TEST_Harness.AssertEqualText "text-written-through-the-format-keeps-its-digits", _
            "00123", CStr(Target.Value)
        TEST_Harness.AssertTrue "text-written-through-the-format-is-a-text", _
            VarType(Target.Value) = vbString, _
            "the value must be a text, not a number"
        TEST_Harness.AssertEqualText "number-format-is-back", Saved, CStr(Target.NumberFormat)

CleanUp:
        On Error GoTo 0
        If FormatChanged Then Target.NumberFormat = Saved
        If RestoreDue Then
            Detail = RemoveAppended()
        End If
        TEST_Harness.AssertTrue "appended-rows-removed", Len(Detail) = 0, Detail
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub


Private Sub CaseAreasAsTables()
'
'==============================================================================
'                               CaseAreasAsTables
'------------------------------------------------------------------------------
' PURPOSE
'   The Value and the Formula of a range of more than one cell are tables
'   with bounds from 1, and those of one cell are not: the cases suite keeps
'   and restores the input areas this way.
'
' STATE OWNERSHIP
'   Appends rows under the position tables and removes them at the end.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim Area         As Range     'Two rows and two columns of an appended table
    Dim One          As Range     'One cell of it
    Dim RestoreDue   As Boolean   'True once rows are appended
    Dim Detail       As String    'Why the appended rows could not be removed

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "areas-as-tables"
        RestoreDue = True
        AppendRows 4, 2
        Set Area = PositionHeader(4).Offset(mFirst(4), 0).Resize(2, 2)
        Set One = Area.Cells(1, 1)
        TEST_Harness.AssertTrue "value-of-an-area-is-a-table", IsArray(Area.Value), _
            "Value of two rows and two columns must be an array"
        TEST_Harness.AssertTrue "formula-of-an-area-is-a-table", IsArray(Area.Formula), _
            "Formula of two rows and two columns must be an array"
        TEST_Harness.AssertEqualDouble "formula-of-an-area-has-two-rows", 2, UBound(Area.Formula, 1)
        TEST_Harness.AssertEqualDouble "formula-of-an-area-starts-at-one", 1, LBound(Area.Formula, 1)
        TEST_Harness.AssertTrue "value-of-one-cell-is-not-a-table", Not IsArray(One.Value), _
            "Value of one cell must not be an array"
        TEST_Harness.AssertTrue "formula-of-one-cell-is-not-a-table", Not IsArray(One.Formula), _
            "Formula of one cell must not be an array"

CleanUp:
        On Error GoTo 0
        If RestoreDue Then
            Detail = RemoveAppended()
        End If
        TEST_Harness.AssertTrue "appended-rows-removed", Len(Detail) = 0, Detail
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub


Private Sub CaseSheetProtection()
'
'==============================================================================
'                             CaseSheetProtection
'------------------------------------------------------------------------------
' PURPOSE
'   Every sheet of the template is protected, and only the input cells are
'   unlocked. A write by code to a locked cell is refused on a protected
'   sheet, and accepted once AllowWriting has applied UserInterfaceOnly to
'   it, which leaves the sheet protected. HB-15.
'
' STATE OWNERSHIP
'   Protects Results again for one write and applies UserInterfaceOnly to it
'   again; changes the label of its header for that write and puts it back.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim Sheet       As Worksheet   'Output sheet used for the writes
    Dim Target      As Range       'A locked label on it
    Dim KeptValue   As Variant     'Its text before the case
    Dim Written     As Boolean     'True while the label is not its own
    Dim Raised      As Long        'Err.Number right after a write

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "sheet-protection"
        Set Target = ThisWorkbook.Names(NM_RESULTS_HEADER).RefersToRange
        Set Sheet = Target.Worksheet
        KeptValue = Target.Value
        TEST_Harness.AssertEqualDouble "sheets-that-are-not-protected", 0, UnprotectedSheets()
        TEST_Harness.AssertTrue "input-cell-of-params-is-unlocked", _
            Not ThisWorkbook.Names("AsOfDate").RefersToRange.Locked, _
            "AsOfDate must be unlocked: users type in it"
        TEST_Harness.AssertTrue "position-cell-is-unlocked", _
            Not PositionHeader(1).Offset(1, 0).Locked, _
            "the first cell of IRPositions under its header must be unlocked"
        TEST_Harness.AssertTrue "label-is-locked", Target.Locked, _
            "the header of Results must be locked"

        'Protected as the template protects it, without UserInterfaceOnly.
        Sheet.Protect Contents:=True
        Written = True
        On Error Resume Next
        Target.Value = "probe"
        Raised = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "write-to-a-locked-cell-is-refused", ERR_PROTECTED_WRITE, Raised
        AllowWriting Sheet
        On Error Resume Next
        Target.Value = "probe"
        Raised = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "write-after-allow-writing-succeeds", 0, Raised
        TEST_Harness.AssertTrue "sheet-is-still-protected", Sheet.ProtectContents, _
            "AllowWriting must leave the sheet protected"

CleanUp:
        On Error GoTo 0
        If Written Then Target.Value = KeptValue
        AllowWriting Sheet
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub


Private Sub CaseStatusOnTheOtherSheets()
'
'==============================================================================
'                          CaseStatusOnTheOtherSheets
'------------------------------------------------------------------------------
' PURPOSE
'   After a run, E1 of each calculation sheet and of Checks shows the status
'   line of Results. HB-10.
'
' STATE OWNERSHIP
'   Rewrites the output sheets; the case that follows rewrites them again.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim ResultText   As String   'RunSSA_Silent result line

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "status-on-the-other-sheets"
        ResultText = RunSSA_Silent()
        TEST_Harness.AssertEqualText "run-outcome", "RESULT=OK", Left$(ResultText, 9)
        TEST_Harness.AssertEqualText "status-line-of-the-run", "OK: ", Left$(StatusLine(), 4)
        TEST_Harness.AssertEqualText "status-on-ir-zones", StatusLine(), _
            PlainText(ThisWorkbook.Worksheets("IR_Zones").Range("E1").Value)
        TEST_Harness.AssertEqualText "status-on-ir-ladder", StatusLine(), _
            PlainText(ThisWorkbook.Worksheets("IR_Ladder").Range("E1").Value)
        TEST_Harness.AssertEqualText "status-on-buckets", StatusLine(), _
            PlainText(ThisWorkbook.Worksheets("Buckets").Range("E1").Value)
        TEST_Harness.AssertEqualText "status-on-ircalc", StatusLine(), _
            PlainText(ThisWorkbook.Worksheets("IRCalc").Range("E1").Value)
        TEST_Harness.AssertEqualText "status-on-eqcalc", StatusLine(), _
            PlainText(ThisWorkbook.Worksheets("EQCalc").Range("E1").Value)
        TEST_Harness.AssertEqualText "status-on-cocalc", StatusLine(), _
            PlainText(ThisWorkbook.Worksheets("COCalc").Range("E1").Value)
        TEST_Harness.AssertEqualText "status-on-fxcalc", StatusLine(), _
            PlainText(ThisWorkbook.Worksheets("FXCalc").Range("E1").Value)
        TEST_Harness.AssertEqualText "status-on-checks", StatusLine(), _
            PlainText(ThisWorkbook.Worksheets("Checks").Range("E1").Value)
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseSaveStateAfterTheCheck()
'
'==============================================================================
'                          CaseSaveStateAfterTheCheck
'------------------------------------------------------------------------------
' PURPOSE
'   The status check that the workbook runs when it is opened and before
'   it is saved, CheckStatus_OnEvent, leaves a workbook with no unsaved
'   change as it was when the results are current, and leaves it with an
'   unsaved change when it writes OUTDATED: before the status line. HB-11.
'   The events themselves, Workbook_Open and Workbook_BeforeSave, are
'   checked by hand: INSTALLATION.md.
'
' STATE OWNERSHIP
'   Changes ScalingIR and ThisWorkbook.Saved, and puts both back. Rewrites
'   the output sheets and ends with them cleared.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim WasSaved     As Boolean   'ThisWorkbook.Saved at the start
    Dim Target       As Range     'The IR scaling factor
    Dim KeptValue    As Variant   'Its value before the case
    Dim Changed      As Boolean   'True while the factor is not its own
    Dim Cleared      As String    'ClearOutputs_Silent result line
    Dim ResultText   As String    'RunSSA_Silent result line, not compared

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "save-state-after-the-check"
        WasSaved = ThisWorkbook.Saved
        ResultText = RunSSA_Silent()
        ThisWorkbook.Saved = True
        CheckStatus_OnEvent
        TEST_Harness.AssertTrue "saved-after-a-check-of-current-results", ThisWorkbook.Saved, _
            "a check that writes nothing must leave the workbook saved"
        TEST_Harness.AssertEqualText "status-after-a-check-of-current-results", "OK: ", _
            Left$(StatusLine(), 4)
        Set Target = ThisWorkbook.Names("ScalingIR").RefersToRange
        KeptValue = Target.Value
        Changed = True
        Target.Value = KeptValue + 0.1
        ThisWorkbook.Saved = True
        CheckStatus_OnEvent
        TEST_Harness.AssertEqualText "status-marked-outdated-by-the-check", "OUTDATED: OK: ", _
            Left$(StatusLine(), 14)
        TEST_Harness.AssertTrue "unsaved-after-the-status-was-written", Not ThisWorkbook.Saved, _
            "a check that writes OUTDATED: must leave an unsaved change"
        Target.Value = KeptValue
        Changed = False
        CheckStatus_OnEvent
        TEST_Harness.AssertEqualText "status-current-after-the-input-is-put-back", "OK: ", _
            Left$(StatusLine(), 4)

CleanUp:
        On Error GoTo 0
        If Changed Then Target.Value = KeptValue
        Cleared = ClearOutputs_Silent()
        ThisWorkbook.Saved = WasSaved
        TEST_Harness.AssertEqualText "outputs-cleared", "RESULT=CLEARED", Left$(Cleared, 14)
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub


'------------------------------------------------------------------------------
'                                    HELPERS
'------------------------------------------------------------------------------

Private Sub AllowWritingToSheets()
'
'==============================================================================
'                            AllowWritingToSheets
'------------------------------------------------------------------------------
' PURPOSE
'   Let the suite write the sheets of its workbook, which the template
'   protects: apply UserInterfaceOnly to each of them, as AllowWriting does
'   for the sheets that RunSSA writes.
'
' STATE OWNERSHIP
'   Changes the protection mode of every sheet, not who may change it.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim Sheet       As Worksheet   'One sheet

        For Each Sheet In ThisWorkbook.Worksheets
            AllowWriting Sheet
        Next Sheet

End Sub


Private Function PositionHeader( _
    ByVal Index As Long) _
    As Range
'
'==============================================================================
'                                PositionHeader
'------------------------------------------------------------------------------
' PURPOSE
'   The header cell of one of the four position sheets.
'
' INPUTS
'   Index   1 interest rate, 2 equity, 3 commodity, 4 foreign exchange.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
        Select Case Index
            Case 1: Set PositionHeader = ThisWorkbook.Names(NM_IR_POS_HEADER).RefersToRange
            Case 2: Set PositionHeader = ThisWorkbook.Names(NM_EQ_POS_HEADER).RefersToRange
            Case 3: Set PositionHeader = ThisWorkbook.Names(NM_CO_POS_HEADER).RefersToRange
            Case 4: Set PositionHeader = ThisWorkbook.Names(NM_FX_POS_HEADER).RefersToRange
        End Select

End Function


Private Function LastOffset( _
    ByVal Index As Long) _
    As Long
'
'==============================================================================
'                                  LastOffset
'------------------------------------------------------------------------------
' PURPOSE
'   The rows of a position table: from the row under the header to the last
'   row that holds a PositionID, as M_Main reads it.
'
' INPUTS
'   Index   As PositionHeader.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim Header   As Range       'Header cell
    Dim Sheet    As Worksheet   'Sheet of the table

        Set Header = PositionHeader(Index)
        Set Sheet = Header.Worksheet
        LastOffset = Sheet.Cells(Sheet.Rows.Count, Header.Column).End(xlUp).Row - Header.Row

End Function


Private Sub AppendRows( _
    ByVal Index As Long, _
    ByVal Count As Long)
'
'==============================================================================
'                                  AppendRows
'------------------------------------------------------------------------------
' PURPOSE
'   Make room for rows under a position table: remember how many rows it has
'   and where the first appended row is.
'
' INPUTS
'   Index   As PositionHeader.
'   Count   Rows that the case writes.
'
' STATE OWNERSHIP
'   Sets mOriginal, mFirst and mAppended of the table.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
        mOriginal(Index) = LastOffset(Index)
        mFirst(Index) = mOriginal(Index) + 1
        mAppended(Index) = Count

End Sub


Private Sub PutRow( _
    ByVal Index As Long, _
    ByVal Row As Long, _
    ByVal Values As Variant)
'
'==============================================================================
'                                    PutRow
'------------------------------------------------------------------------------
' PURPOSE
'   Write one appended row of a position table.
'
' INPUTS
'   Index    As PositionHeader.
'   Row      0 for the first appended row, 1 for the next.
'   Values   The cells of the row from the PositionID; Empty leaves one empty,
'            and an error value from CVErr is written as it is.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim Target   As Range   'First cell of the row
    Dim c        As Long    'Cell of the row

        Set Target = PositionHeader(Index).Offset(mFirst(Index) + Row, 0)
        For c = LBound(Values) To UBound(Values)
            If Not IsEmpty(Values(c)) Then
                Target.Offset(0, c - LBound(Values)).Value = Values(c)
            End If
        Next c

End Sub


Private Function RemoveAppended() As String
'
'==============================================================================
'                                RemoveAppended
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the rows a case appended, and check that each table has the rows
'   it had.
'
' RETURNS
'   An empty text when every table is as it was; otherwise what is left.
'
' ERROR POLICY
'   Contains its own error and returns it, so that the case asserts the
'   result on the success and the error path.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim i        As Long      'Position table
    Dim Left     As Long      'Rows that remain
    Dim Widths   As Variant   'Columns of each table

        On Error GoTo Failed
        Widths = Array(IR_INPUT_WIDTH, EQ_INPUT_WIDTH, CO_INPUT_WIDTH, FX_INPUT_WIDTH)
        For i = 1 To 4
            If mAppended(i) > 0 Then
                PositionHeader(i).Offset(mFirst(i), 0).Resize(mAppended(i), Widths(i - 1)).ClearContents
                If LastOffset(i) <> mOriginal(i) Then Left = Left + 1
                mAppended(i) = 0
            End If
        Next i
        If Left <> 0 Then RemoveAppended = CStr(Left) & " table(s) do not have the rows they had"
        Exit Function

Failed:
        RemoveAppended = "error " & CStr(Err.Number) & " while removing rows: " & Err.Description

End Function


Private Function RowCell( _
    ByVal HeaderName As String, _
    ByVal Index As Long, _
    ByVal ColumnOffset As Long) _
    As Variant
'
'==============================================================================
'                                    RowCell
'------------------------------------------------------------------------------
' PURPOSE
'   The value of one cell of a calculation sheet.
'
' INPUTS
'   HeaderName     Workbook name of the sheet's header cell.
'   Index          Row under the header, from 1: the position's row.
'   ColumnOffset   Columns to the right of the PositionID.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
        RowCell = ThisWorkbook.Names(HeaderName).RefersToRange.Offset(Index, ColumnOffset).Value

End Function


Private Function ErrorCells() As Long
'
'==============================================================================
'                                  ErrorCells
'------------------------------------------------------------------------------
' PURPOSE
'   Count the cells that hold an error value in the output sheets.
'
' RETURNS
'   The number of such cells, in the used range of each sheet.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
    Dim Sheets   As Variant   'Names of the output sheets
    Dim Item     As Variant   'One name
    Dim Cells    As Variant   'Values of a used range
    Dim r        As Long      'Row
    Dim c        As Long      'Column
    Dim Count    As Long      'Cells that hold an error value

        Sheets = Array("IR_Zones", "IR_Ladder", "Buckets", "IRCalc", "EQCalc", "COCalc", _
                       "FXCalc", "Results", "Checks")
        For Each Item In Sheets
            Cells = ThisWorkbook.Worksheets(CStr(Item)).UsedRange.Value
            If IsArray(Cells) Then
                For r = LBound(Cells, 1) To UBound(Cells, 1)
                    For c = LBound(Cells, 2) To UBound(Cells, 2)
                        If IsError(Cells(r, c)) Then Count = Count + 1
                    Next c
                Next r
            ElseIf IsError(Cells) Then
                Count = Count + 1
            End If
        Next Item
        ErrorCells = Count

End Function


Private Function UnprotectedSheets() As Long
'
'==============================================================================
'                             UnprotectedSheets
'------------------------------------------------------------------------------
' PURPOSE
'   Count the sheets of the workbook that are not protected.
'
' RETURNS
'   The number of sheets whose contents are not protected.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim Sheet       As Worksheet   'One sheet
    Dim Count       As Long        'Sheets that are not protected

        For Each Sheet In ThisWorkbook.Worksheets
            If Not Sheet.ProtectContents Then Count = Count + 1
        Next Sheet
        UnprotectedSheets = Count

End Function


Private Function StatusLine() As String
'
'==============================================================================
'                                  StatusLine
'------------------------------------------------------------------------------
' PURPOSE
'   Return the status line of Results.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
        StatusLine = NamedText(NM_RESULTS_STATUS)

End Function


Private Function NamedText( _
    ByVal CellName As String) _
    As String
'
'==============================================================================
'                                   NamedText
'------------------------------------------------------------------------------
' PURPOSE
'   Return the text of the cell a workbook name refers to; empty for an
'   empty cell.
'
' INPUTS
'   CellName   A workbook name that refers to one cell.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
        NamedText = PlainText(ThisWorkbook.Names(CellName).RefersToRange.Value)

End Function


Private Function NumberOf( _
    ByVal Value As Variant) _
    As Double
'
'==============================================================================
'                                   NumberOf
'------------------------------------------------------------------------------
' PURPOSE
'   The number a cell holds, for an assertion that compares numbers.
'
' ERROR POLICY
'   Raises error 13 when Value is not a number, so that an empty cell, a
'   text or an error value is never compared as 0; the case records it.
'
' UPDATED
'   2026-11-18
'==============================================================================
'
        If Not IsNumberValue(Value) Then
            Err.Raise 13, "TEST_Probes.NumberOf", "The result is not a number."
        End If
        NumberOf = CDbl(Value)

End Function
