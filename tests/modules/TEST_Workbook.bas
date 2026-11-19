Attribute VB_Name = "TEST_Workbook"
'==============================================================================
' MODULE: TEST_Workbook
'------------------------------------------------------------------------------
' PURPOSE
'   The workbook suite of the workbook's tests, as cases and assertions of
'   TEST_Harness: the whole workbook on the shipped sample data, runs that
'   fail, and inputs changed after a run. Every case needs the workbook and
'   runs the calculation through its entry points; the tests that need no
'   sheet are in TEST_Suite.
'
' PUBLIC SURFACE
'   RunWorkbookTests is the documented test entry point. Test_RunWorkbook
'   runs the same suite and returns its result line where a caller cannot
'   read an error, as tools/devbuild/lo_macro.py. CertifyWorkbook certifies
'   the run and saves its log and its certification log as evidence of an
'   Excel run. Option Private Module keeps them out of the workbook's
'   external surface. Never part of the production import set.
'
' DEPENDENCIES
'   TEST_Harness; M_Main, CORE_Config and CORE_Util; the workbook's names
'   and sheets, with the sample data the template ships.
'
' STATE OWNERSHIP
'   Rewrites the output sheets through RunSSA_Silent, and leaves them
'   cleared: the last case ends with ClearOutputs_Silent, so no result of a
'   test run stays in the workbook. Two cases change a cell or a workbook
'   name for one operation and put it back before they assert. Excel's
'   settings are those the entry points restore; the harness compares them
'   with their values at the start of the run.
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START. An unexpected error inside a case is
'   recorded as a failure of that case, and the run continues.
'   Test_RunWorkbook contains those errors; CertifyWorkbook records them,
'   saves the log, and raises only when the certification did not pass.
'
' USAGE
'   In a development workbook, type TEST_Workbook.RunWorkbookTests in the
'   Immediate window. To record an Excel run, type
'   TEST_Workbook.CertifyWorkbook "<folder>", "<checkout>\tests\fixtures"
'   (INSTALLATION.md, Recording an Excel run).
'
' COMPATIBILITY
'   Excel VBA; also runs under LibreOffice's VBA mode through the tools of
'   tools/devbuild/, which is never Excel evidence.
'
' REFERENCE
'   The requirement and the risk-weighted assets of the sample data are
'   those of the test case illustrative-sample-requirement
'   (tests/expected/), which tests/derivations/ssa.py derives.
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
        Private Const SUITE_NAME          As String = "workbook"
        Private Const EXPECTED_CASES      As Long = 4
        Private Const EXPECTED_ASSERTIONS As Long = 41
    'Tolerances of a derived value (docs/methodology/CALCULATION_CONTRACTS.md,
    'Decisions).
        Private Const TOL_ABSOLUTE        As Double = 0.000001      'Absolute tolerance
        Private Const TOL_RELATIVE        As Double = 0.000000001   'Relative tolerance
    'The sample data: test case illustrative-sample-requirement.
        Private Const SAMPLE_REQUIREMENT  As Double = 48267441#     'Own funds requirement
        Private Const SAMPLE_RWA          As Double = 603343012.5   'Risk-weighted assets
        Private Const SAMPLE_POSITIONS    As Long = 35              'Sample rows with status OK



'------------------------------------------------------------------------------
'                                 ENTRY POINTS
'------------------------------------------------------------------------------

Public Sub RunWorkbookTests()
'
'==============================================================================
'                               RunWorkbookTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run the workbook suite. A passing run ends with
'   RESULT=PASS; suite=workbook; operation=PASS; completeness=COMPLETE;
'   cases=4; assertions=41; failures=0; cleanup=PASS; recovery=NONE.
'
' STATE OWNERSHIP
'   Rewrites the output sheets and leaves them cleared.
'
' ERROR POLICY
'   As TEST_Harness.EndRun: raises when the run does not pass.
'
' UPDATED
'   2026-11-17
'==============================================================================
'
        RunSuite

End Sub


Public Function Test_RunWorkbook() As String
'
'==============================================================================
'                               Test_RunWorkbook
'------------------------------------------------------------------------------
' PURPOSE
'   Run the workbook suite for a caller that reads a result and not an
'   error.
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
'   2026-11-17
'==============================================================================
'
        On Error Resume Next
        RunSuite
        On Error GoTo 0
        Test_RunWorkbook = TEST_Harness.LastResult()

End Function


Public Sub CertifyWorkbook( _
    ByVal Folder As String, _
    ByVal Fixtures As String)
'
'==============================================================================
'                                CertifyWorkbook
'------------------------------------------------------------------------------
' PURPOSE
'   Certify the workbook suite's run as evidence of an Excel run: run it,
'   save its log as workbook.log, and write workbook-certification.log.
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
'   2026-11-17
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
        TEST_Harness.SaveLog Folder & "\workbook.log"
        TEST_Harness.Echo "SAVED=" & Folder & "\workbook.log; mode=NORMAL; error=" & _
            CStr(RunErrorNumber) & IIf(RunErrorNumber = 0, vbNullString, " " & RunErrorDescription)
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
        TEST_Harness.EndCertification Folder & "\workbook-certification.log"

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
'   2026-11-17
'==============================================================================
'
    Dim ErrNumber        As Long     'Runner error number; 0 for none
    Dim ErrSource        As String   'Runner error source
    Dim ErrDescription   As String   'Runner error description

        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS
        On Error GoTo RunFailed
        AllowWritingToSheets
        CaseSamplePortfolio
        CaseFailedRun
        CaseVolume
        CaseOutdatedResults

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

Private Sub CaseSamplePortfolio()
'
'==============================================================================
'                              CaseSamplePortfolio
'------------------------------------------------------------------------------
' PURPOSE
'   Run the workbook on the shipped sample data and compare the totals with
'   those of the test case; every test of TestCatalogue must be PASS.
'
' STATE OWNERSHIP
'   Rewrites the output sheets.
'
' REFERENCE
'   Test case illustrative-sample-requirement: requirement 48,267,441 and
'   risk-weighted assets 603,343,012.5. The sample data has 35 rows that a
'   calculation uses.
'
' UPDATED
'   2026-11-17
'==============================================================================
'
    Dim ResultText   As String    'RunSSA_Silent result line
    Dim Results      As Variant   'TestCatalogue Result column
    Dim Failures     As Long      'Tests that are not PASS
    Dim Item         As Variant   'One result

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "sample-portfolio"
        ResultText = RunSSA_Silent()
        TEST_Harness.AssertEqualText "run-outcome", "RESULT=OK", Left$(ResultText, 9)
        TEST_Harness.AssertNear "requirement", SAMPLE_REQUIREMENT, _
            NumberOf(ThisWorkbook.Names("SSA_Requirement").RefersToRange.Value), _
            TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "risk-weighted-assets", SAMPLE_RWA, _
            NumberOf(ThisWorkbook.Names("SSA_RWA").RefersToRange.Value), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertEqualDouble "positions-used", SAMPLE_POSITIONS, _
            Val(Split(Split(ResultText, "positions_used=")(1), ";")(0))
        Application.Calculate
        Results = ThisWorkbook.Names("Tests_Result").RefersToRange.Value
        For Each Item In Results
            If CStr(Item) <> "PASS" Then Failures = Failures + 1
        Next Item
        TEST_Harness.AssertEqualDouble "catalogue-tests-not-pass", 0, Failures
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseFailedRun()
'
'==============================================================================
'                                 CaseFailedRun
'------------------------------------------------------------------------------
' PURPOSE
'   A run that raises an error, while it writes or before, leaves no result
'   of that run or of the one before, and a status line that says FAILED:
'   the status line, the requirement, the first calculation row, the error
'   count of Checks, the fingerprint and what the check returns.
'
' STATE OWNERSHIP
'   Rewrites the output sheets. Changes the first rule identifier of Checks
'   for one run, and the range of the name Ladder_RW for another, and puts
'   both back before it asserts.
'
' UPDATED
'   2026-11-17
'==============================================================================
'
    Dim Target       As Range     'Cell changed to make a run fail
    Dim Saved        As Variant   'Its value before the test
    Dim Ladder       As Object    'Workbook name changed to make a run fail
    Dim SavedRange   As String    'What the name referred to
    Dim ErrNumber    As Long      'Error number of the failed run

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "failed-run"
        'While writing: the rules of Checks are written after every result
        'table, and the first rule is not found under its header.
        TEST_Harness.AssertEqualText "run-before-the-failure-while-writing", "RESULT=OK", _
            Left$(RunSSA_Silent(), 9)
        Set Target = ThisWorkbook.Names(NM_CHECKS_HEADER).RefersToRange.Offset(1, 0)
        Saved = Target.Value
        Target.Value = "X01"
        ErrNumber = RunError()
        Target.Value = Saved
        TEST_Harness.AssertRaised "while-writing-error", ERR_LAYOUT, ErrNumber
        TEST_Harness.AssertEqualText "while-writing-status-line", "FAILED: ", Left$(StatusLine(), 8)
        TEST_Harness.AssertEqualText "while-writing-requirement", vbNullString, NamedText("SSA_Requirement")
        TEST_Harness.AssertEqualText "while-writing-first-calculation-row", vbNullString, _
            PlainText(ThisWorkbook.Names(NM_IR_CALC_HEADER).RefersToRange.Offset(1, 0).Value)
        TEST_Harness.AssertEqualText "while-writing-error-count-of-checks", vbNullString, _
            NamedText(NM_CHECKS_ERRORS)
        TEST_Harness.AssertEqualText "while-writing-fingerprint", vbNullString, NamedText(NM_FINGERPRINT)
        TEST_Harness.AssertEqualText "while-writing-check", "RESULT=FAILED; cleanup=PASS", CheckOutcome()

        'Before writing: a column of the maturity ladder is one row short.
        TEST_Harness.AssertEqualText "run-before-the-failure-while-reading", "RESULT=OK", _
            Left$(RunSSA_Silent(), 9)
        Set Ladder = ThisWorkbook.Names(NM_LADDER_RW)
        SavedRange = Ladder.RefersTo
        Set Target = Ladder.RefersToRange
        Ladder.RefersTo = "='" & Target.Worksheet.Name & "'!" & _
                          Target.Resize(Target.Rows.Count - 1, 1).Address
        ErrNumber = RunError()
        Ladder.RefersTo = SavedRange
        TEST_Harness.AssertRaised "while-reading-error", ERR_LAYOUT, ErrNumber
        TEST_Harness.AssertEqualText "while-reading-status-line", "FAILED: ", Left$(StatusLine(), 8)
        TEST_Harness.AssertEqualText "while-reading-requirement", vbNullString, NamedText("SSA_Requirement")
        TEST_Harness.AssertEqualText "while-reading-first-calculation-row", vbNullString, _
            PlainText(ThisWorkbook.Names(NM_IR_CALC_HEADER).RefersToRange.Offset(1, 0).Value)
        TEST_Harness.AssertEqualText "while-reading-error-count-of-checks", vbNullString, _
            NamedText(NM_CHECKS_ERRORS)
        TEST_Harness.AssertEqualText "while-reading-fingerprint", vbNullString, NamedText(NM_FINGERPRINT)
        TEST_Harness.AssertEqualText "while-reading-check", "RESULT=FAILED; cleanup=PASS", CheckOutcome()
        TEST_Harness.AssertEqualText "run-after-the-failures", "RESULT=OK", Left$(RunSSA_Silent(), 9)
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseVolume()
'
'==============================================================================
'                                  CaseVolume
'------------------------------------------------------------------------------
' PURPOSE
'   A position sheet that holds more rows than the workbook supports stops
'   the run: a PositionID on the row after the last supported one makes the
'   table 501 rows, and a run on exactly 500 rows is not stopped. The stop
'   leaves no result, counts one error on Checks and lists the sheet and the
'   rows found.
'
' STATE OWNERSHIP
'   Rewrites the output sheets. Writes one PositionID below the sample rows
'   of IRPositions, twice, and clears the cell before it asserts.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim Header       As Range     'First header cell of IRPositions
    Dim Findings     As Range     'Header of the Checks findings table

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "volume"
        Set Header = ThisWorkbook.Names(NM_IR_POS_HEADER).RefersToRange
        Header.Offset(MAX_POSITION_ROWS, 0).Value = "V500"
        TEST_Harness.AssertEqualText "run-on-500-rows-is-not-stopped", "RESULT=INCOMPLETE", _
            Left$(RunSSA_Silent(), 17)
        Header.Offset(MAX_POSITION_ROWS, 0).ClearContents
        Header.Offset(MAX_POSITION_ROWS + 1, 0).Value = "V501"
        TEST_Harness.AssertEqualText "run-on-501-rows-is-stopped", "RESULT=STOPPED;", _
            Left$(RunSSA_Silent(), 15)
        Header.Offset(MAX_POSITION_ROWS + 1, 0).ClearContents
        TEST_Harness.AssertEqualText "status-line", "STOPPED: 1 unusable input(s)", Left$(StatusLine(), 28)
        TEST_Harness.AssertEqualText "no-result", vbNullString, NamedText("SSA_Requirement")
        TEST_Harness.AssertEqualText "error-count-of-checks", "1", NamedText(NM_CHECKS_ERRORS)
        Set Findings = ThisWorkbook.Names(NM_FINDINGS_HEADER).RefersToRange
        TEST_Harness.AssertEqualText "finding-sheet", SH_IR_POS, PlainText(Findings.Offset(1, 1).Value)
        TEST_Harness.AssertEqualText "finding-reason", _
            ST_INVALID & "501 rows found; the workbook supports at most 500", _
            PlainText(Findings.Offset(1, 6).Value)
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseOutdatedResults()
'
'==============================================================================
'                              CaseOutdatedResults
'------------------------------------------------------------------------------
' PURPOSE
'   After a run, a changed parameter or position makes the check mark the
'   results OUTDATED; the value put back makes it take the mark away.
'
' STATE OWNERSHIP
'   Rewrites the output sheets and ends with them cleared. Changes the IR
'   scaling factor and the value of the first interest rate position for
'   one check each, and puts both back.
'
' UPDATED
'   2026-11-17
'==============================================================================
'
    Dim Target   As Range     'Input cell changed
    Dim Saved    As Variant   'Its value before the test

        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "outdated-results"
        TEST_Harness.AssertEqualText "run-before-the-changes", "RESULT=OK", Left$(RunSSA_Silent(), 9)
        TEST_Harness.AssertEqualText "check-after-the-run", "RESULT=CURRENT; cleanup=PASS", _
            CheckOutcome()
        TEST_Harness.AssertEqualText "status-after-the-run", "OK: ", Left$(StatusLine(), 4)

        'A parameter.
        Set Target = ThisWorkbook.Names("ScalingIR").RefersToRange
        Saved = Target.Value
        Target.Value = Saved + 0.1
        TEST_Harness.AssertEqualText "check-after-a-parameter-changed", _
            "RESULT=OUTDATED; cleanup=PASS", CheckOutcome()
        TEST_Harness.AssertEqualText "status-after-a-parameter-changed", "OUTDATED: OK: ", _
            Left$(StatusLine(), 14)
        TEST_Harness.AssertEqualText "check-again-still-changed", "RESULT=OUTDATED; cleanup=PASS", _
            CheckOutcome()
        TEST_Harness.AssertEqualText "status-is-marked-once", "OUTDATED: OK: ", Left$(StatusLine(), 14)
        Target.Value = Saved
        TEST_Harness.AssertEqualText "check-after-the-parameter-is-put-back", _
            "RESULT=CURRENT; cleanup=PASS", CheckOutcome()
        TEST_Harness.AssertEqualText "status-after-the-parameter-is-put-back", "OK: ", _
            Left$(StatusLine(), 4)

        'A position, through the entry point of the event procedures.
        Set Target = ThisWorkbook.Names(NM_IR_POS_HEADER).RefersToRange.Offset(1, IR_VALUE - 1)
        Saved = Target.Value
        Target.Value = Saved + 1
        CheckStatus_OnEvent
        TEST_Harness.AssertEqualText "status-after-a-position-changed", "OUTDATED: OK: ", _
            Left$(StatusLine(), 14)
        Target.Value = Saved
        CheckStatus_OnEvent
        TEST_Harness.AssertEqualText "status-after-the-position-is-put-back", "OK: ", _
            Left$(StatusLine(), 4)

        'No result to check; the output sheets stay cleared.
        ClearOutputs_Silent
        TEST_Harness.AssertEqualText "check-without-results", _
            "RESULT=NOT_CALCULATED; cleanup=PASS", CheckOutcome()
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

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


Private Function RunError() As Long
'
'==============================================================================
'                                   RunError
'------------------------------------------------------------------------------
' PURPOSE
'   Run RunSSA_Silent where it is expected to raise an error.
'
' RETURNS
'   The error number; 0 when the run raised none.
'
' ERROR POLICY
'   Contains the error of the run, so that the caller can put back what it
'   changed.
'
' UPDATED
'   2026-11-17
'==============================================================================
'
    Dim ResultText   As String   'Result line, not used

        On Error GoTo Failed
        ResultText = RunSSA_Silent()
        RunError = 0
        Exit Function

Failed:
        RunError = Err.Number

End Function


Private Function CheckOutcome() As String
'
'==============================================================================
'                                 CheckOutcome
'------------------------------------------------------------------------------
' PURPOSE
'   Run CheckStatus_Silent between a change of an input and its undoing.
'
' RETURNS
'   The result line, or "ERROR", the number and the description.
'
' ERROR POLICY
'   Contains the error of the check, so that the caller can put back what
'   it changed.
'
' UPDATED
'   2026-11-17
'==============================================================================
'
        On Error GoTo Failed
        CheckOutcome = CheckStatus_Silent()
        Exit Function

Failed:
        CheckOutcome = "ERROR " & Err.Number & " " & Err.Description

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
'   2026-11-17
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
'   2026-11-17
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
' INPUTS
'   Value   The value of a cell.
'
' ERROR POLICY
'   Raises error 13 when Value is not a number, so that an empty cell, a
'   text or an error value is never compared as 0; the case records it.
'
' UPDATED
'   2026-11-17
'==============================================================================
'
        If Not IsNumberValue(Value) Then
            Err.Raise 13, "TEST_Workbook.NumberOf", "The result is not a number."
        End If
        NumberOf = CDbl(Value)

End Function
