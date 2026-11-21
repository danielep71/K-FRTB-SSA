Attribute VB_Name = "TEST_Workbook"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Workbook
'==============================================================================
' PURPOSE
'   Workbook suite of the test harness: runs the whole workbook on the shipped
'   sample data, runs that fail, and inputs changed after a run, as cases and
'   assertions of TEST_Harness. Every case needs the workbook and runs the
'   calculation through its entry points; the tests that need no sheet are in
'   TEST_Suite
'
' WHY THIS EXISTS
'   Holds the tests that need the workbook, its sheets and its names, apart from
'   the tests of TEST_Suite that need no sheet. The cases run the calculation
'   through the entry points of the workbook and assert what it leaves on the
'   sheets
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunWorkbookTests is the documented test entry point. Test_RunWorkbook runs
'   the same suite and returns its result line where a caller cannot read an
'   error, as tools/devbuild/lo_macro.py. CertifyWorkbook certifies the run and
'   saves its log and its certification log as evidence of an Excel run
'
'   RunSuite runs five cases in a fixed order between BeginRun and EndRun:
'     - CaseSamplePortfolio: the sample data against the totals of the test case
'     - CaseFailedRun: runs that fail while they write and before they write
'     - CaseVolume: a position sheet with more rows than the workbook supports
'     - CaseEligibility: the months of Eligibility under the Basel regime and
'       under the CRR, and an invalid month
'     - CaseOutdatedResults: a parameter and a position changed after a run
'
'   The private helpers are AllowWritingToSheets, RunError, CheckOutcome,
'   StatusLine, NamedText and NumberOf
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED, or
'   the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START. An unexpected error inside a case is
'   recorded as a failure of that case, and the run continues.
'   Test_RunWorkbook contains those errors; CertifyWorkbook records them, saves
'   the log, and raises only when the certification did not pass
'
' DEPENDENCIES
'   TEST_Harness; M_Main, CORE_Config and CORE_Util; the workbook's names and
'   sheets, with the sample data the template ships
'
'   Reference: the requirement and the risk-weighted assets of the sample data
'   are those of the test case illustrative-sample-requirement (tests/expected/),
'   which tests/derivations/ssa.py derives
'
' NOTES
'   Option Explicit requires explicit declarations, and Option Private Module
'   keeps the tests inside this project, out of the workbook's external
'   surface. Never part of the production import set
'
'   State ownership: rewrites the output sheets through RunSSA_Silent, and
'   leaves them cleared: the last case ends with ClearOutputs_Silent, so no
'   result of a test run stays in the workbook. Two cases change a cell or a
'   workbook name for one operation and put it back before they assert;
'   CaseEligibility changes three input cells for two runs and puts them back
'   Excel's settings are those the entry points restore; the harness compares
'   them with their values at the start of the run
'
'   Usage: in a development workbook, type this in the Immediate window:
'
'       TEST_Workbook.RunWorkbookTests
'
'   To record an Excel run, type this (INSTALLATION.md, Recording an Excel
'   run):
'
'       TEST_Workbook.CertifyWorkbook "<folder>", "<checkout>\tests\fixtures"
'
'   Compatibility: Excel VBA; also runs under LibreOffice's VBA mode through the
'   tools of tools/devbuild/, which is never Excel evidence
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-21
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME          As String = "workbook"    'Name of the suite in its records
    Private Const EXPECTED_CASES      As Long = 5               'Cases of a complete run
    Private Const EXPECTED_ASSERTIONS As Long = 56              'Assertions of a complete run

    'Tolerances of a derived value (docs/methodology/CALCULATION_CONTRACTS.md, Decisions)
    Private Const TOL_ABSOLUTE        As Double = 0.000001      'Absolute tolerance
    Private Const TOL_RELATIVE        As Double = 0.000000001   'Relative tolerance

    'The sample data: test case illustrative-sample-requirement
    Private Const SAMPLE_REQUIREMENT  As Double = 48267441#     'Own funds requirement
    Private Const SAMPLE_RWA          As Double = 603343012.5   'Risk-weighted assets
    Private Const SAMPLE_POSITIONS    As Long = 35              'Sample rows with status OK


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunWorkbookTests()

'
'------------------------------------------------------------------------------
'                              RUN WORKBOOK SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Run the workbook suite. A passing run ends with
'   RESULT=PASS; suite=workbook; operation=PASS; completeness=COMPLETE;
'   cases=5; assertions=56; failures=0; cleanup=PASS; recovery=NONE
'
' WHY THIS EXISTS
'   Gives the workbook suite one documented entry point that raises when the
'   run does not pass; Test_RunWorkbook is the variant that returns the result
'   line instead
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls RunSuite
'
' ERROR POLICY
'   As TEST_Harness.EndRun: raises when the run does not pass
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   Rewrites the output sheets and leaves them cleared
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run every case of the suite between BeginRun and EndRun
        RunSuite

End Sub

Public Function Test_RunWorkbook() As String

'
'------------------------------------------------------------------------------
'                             RUN SUITE FOR RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Run the workbook suite for a caller that reads a result and not an
'   error
'
' WHY THIS EXISTS
'   A caller such as tools/devbuild/lo_macro.py cannot read an error, and reads
'   the result line instead
'
' INPUTS
'   None
'
' RETURNS
'   The RESULT line of the run, as TEST_Harness.LastResult gives it
'
' BEHAVIOR
'   Runs the suite with the error of the run contained, then returns the
'   result line of the harness
'
' ERROR POLICY
'   Contains the error of a run that does not pass, or of a refused start:
'   the result line says so
'
' DEPENDENCIES
'   RunSuite; TEST_Harness.LastResult
'
' NOTES
'   Rewrites the output sheets and leaves them cleared
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run the suite and contain its error, so that the result line can be read
        On Error Resume Next
        RunSuite
        On Error GoTo 0

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return the RESULT line of the run
        Test_RunWorkbook = TEST_Harness.LastResult()

End Function

Public Sub CertifyWorkbook( _
    ByVal Folder As String, _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                            CERTIFY WORKBOOK SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Certify the workbook suite's run as evidence of an Excel run: run it,
'   save its log as workbook.log, and write workbook-certification.log
'
' WHY THIS EXISTS
'   Records an Excel run of the suite in logs that can be kept as evidence
'
' INPUTS
'   Folder
'     An existing folder for the logs, without a trailing separator
'
'   Fixtures
'     The full path of tests/fixtures/ in the checkout; not read, and taken so
'     that every certification entry point has the same two arguments
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the certification, runs the suite with its error contained, saves
'   the log as workbook.log, certifies the run and writes
'   workbook-certification.log
'
' ERROR POLICY
'   The error of the run is recorded, not raised, so that the log is always
'   saved. Raises ERR_TEST_FAILED after writing the certification log when the
'   certification did not pass. An error of BeginCertification or SaveLog
'   propagates; run TEST_Harness.ResetTests before trying again
'
' DEPENDENCIES
'   RunSuite; TEST_Harness.BeginCertification, SaveLog, Echo, CertifyRun and
'   EndCertification
'
' NOTES
'   Rewrites the output sheets and leaves them cleared
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RunErrorNumber      As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription As String   'Its description

'------------------------------------------------------------------------------
' BEGIN CERTIFICATION
'------------------------------------------------------------------------------
    'Begin the certification of the suite
        TEST_Harness.BeginCertification SUITE_NAME

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run the suite and keep the error that ended it, so that the log is always saved
        On Error Resume Next
        RunSuite
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
        On Error GoTo 0

'------------------------------------------------------------------------------
' SAVE LOG
'------------------------------------------------------------------------------
    'Save the log of the run as workbook.log and echo where it went and how the run ended
        TEST_Harness.SaveLog Folder & "\workbook.log"
        TEST_Harness.Echo "SAVED=" & Folder & "\workbook.log; mode=NORMAL; error=" & _
            CStr(RunErrorNumber) & IIf(RunErrorNumber = 0, vbNullString, " " & RunErrorDescription)

'------------------------------------------------------------------------------
' CERTIFY RUN
'------------------------------------------------------------------------------
    'Certify the run and write the certification log
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
        TEST_Harness.EndCertification Folder & "\workbook-certification.log"

End Sub

'
'------------------------------------------------------------------------------
'
'                                    RUNNER
'
'------------------------------------------------------------------------------
'

Private Sub RunSuite()

'
'------------------------------------------------------------------------------
'                               RUN SUITE CASES
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. The order
'   matters: the last case clears the output sheets
'
' WHY THIS EXISTS
'   Gives RunWorkbookTests, Test_RunWorkbook and CertifyWorkbook one runner, and
'   ends the run through EndRun whether or not a case raised
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the run, allows writing to the sheets, runs the five cases and ends
'   the run, passing EndRun the error that escaped a case, if any
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it and
'   re-raises it unchanged. BeginRun's refusal propagates directly
'
' DEPENDENCIES
'   TEST_Harness.BeginRun and EndRun; AllowWritingToSheets, CaseSamplePortfolio,
'   CaseFailedRun, CaseVolume, CaseEligibility and CaseOutdatedResults; SUITE_NAME,
'   EXPECTED_CASES and EXPECTED_ASSERTIONS
'
' NOTES
'   None
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Begin the run; a refusal of BeginRun propagates directly, as no handler is active yet
        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS
    'Capture an error that escapes a case, so that EndRun can report it
        On Error GoTo RunFailed

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Allow writing, then run the cases in their fixed order; the last case clears the output sheets
        AllowWritingToSheets
        CaseSamplePortfolio
        CaseFailedRun
        CaseVolume
        CaseEligibility
        CaseOutdatedResults

'------------------------------------------------------------------------------
' END RUN
'------------------------------------------------------------------------------
Finish:
    'Disable the handler and hand the run, with the error it captured if any, to EndRun
        On Error GoTo 0
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the error that escaped a case, then resume at Finish so that EndRun receives it
RunFailed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume Finish

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseSamplePortfolio()

'
'------------------------------------------------------------------------------
'                            CASE SAMPLE PORTFOLIO
'------------------------------------------------------------------------------
' PURPOSE
'   Run the workbook on the shipped sample data and compare the totals with
'   those of the test case; every test of TestCatalogue must be PASS
'
' WHY THIS EXISTS
'   Shows that the shipped workbook, run through its entry point, gives the
'   totals of the test case and a catalogue without a failing test
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case sample-portfolio, runs RunSSA_Silent, asserts the outcome,
'   the requirement, the risk-weighted assets and the rows used, recalculates
'   the workbook and asserts that no test of TestCatalogue is other than PASS
'
' ERROR POLICY
'   An unexpected error is recorded as a failure of the case by
'   TEST_Harness.RecordCaseError; the run continues
'
' DEPENDENCIES
'   TEST_Harness; RunSSA_Silent; NumberOf; the names SSA_Requirement, SSA_RWA and
'   Tests_Result; SAMPLE_REQUIREMENT, SAMPLE_RWA, SAMPLE_POSITIONS and the
'   tolerances
'
'   Reference: test case illustrative-sample-requirement: requirement
'   48,267,441 and risk-weighted assets 603,343,012.5. The sample data has 35
'   rows that a calculation uses
'
' NOTES
'   Rewrites the output sheets
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultText  As String   'RunSSA_Silent result line
    Dim Results     As Variant  'TestCatalogue Result column
    Dim Failures    As Long     'Tests that are not PASS
    Dim Item        As Variant  'One result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an error of the case as a failure of the case, and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "sample-portfolio"

'------------------------------------------------------------------------------
' RUN WORKBOOK
'------------------------------------------------------------------------------
    'Run the calculation on the shipped sample data and assert that it completed
        ResultText = RunSSA_Silent()
        TEST_Harness.AssertEqualText "run-outcome", "RESULT=OK", Left$(ResultText, 9)

'------------------------------------------------------------------------------
' COMPARE TOTALS
'------------------------------------------------------------------------------
    'Compare the requirement and the risk-weighted assets with the test case, within the tolerances
        TEST_Harness.AssertNear "requirement", SAMPLE_REQUIREMENT, _
            NumberOf(ThisWorkbook.Names("SSA_Requirement").RefersToRange.Value), _
            TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "risk-weighted-assets", SAMPLE_RWA, _
            NumberOf(ThisWorkbook.Names("SSA_RWA").RefersToRange.Value), TOL_ABSOLUTE, TOL_RELATIVE
    'Compare the rows the calculation used with the sample rows that have status OK
        TEST_Harness.AssertEqualDouble "positions-used", SAMPLE_POSITIONS, _
            Val(Split(Split(ResultText, "positions_used=")(1), ";")(0))

'------------------------------------------------------------------------------
' CHECK CATALOGUE
'------------------------------------------------------------------------------
    'Recalculate and count the tests of TestCatalogue whose result is not PASS
        Application.Calculate
        Results = ThisWorkbook.Names("Tests_Result").RefersToRange.Value
        For Each Item In Results
            If CStr(Item) <> "PASS" Then Failures = Failures + 1
        Next Item
    'Assert that every test of the catalogue is PASS
        TEST_Harness.AssertEqualDouble "catalogue-tests-not-pass", 0, Failures

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record an error of the case as a failure of the case; the run continues
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseFailedRun()

'
'------------------------------------------------------------------------------
'                               CASE FAILED RUN
'------------------------------------------------------------------------------
' PURPOSE
'   A run that raises an error, while it writes or before, leaves no result
'   of that run or of the one before, and a status line that says FAILED:
'   the status line, the requirement, the first calculation row, the error
'   count of Checks, the fingerprint and what the check returns
'
' WHY THIS EXISTS
'   Shows that a calculation that fails does not leave the results of an
'   earlier run in the workbook
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case failed-run. First the rules of Checks fail while they are
'   written: the first rule identifier is changed, a run is made, the
'   identifier is put back and the outcome is asserted. Then a calculation
'   fails before it writes: the range of the name Ladder_RW is made one row
'   short, a run is made, the name is put back and the outcome is asserted.
'   A last run asserts that the calculation succeeds again
'
' ERROR POLICY
'   An unexpected error is recorded as a failure of the case by
'   TEST_Harness.RecordCaseError; the run continues. The error of a run that
'   is expected to fail is contained by RunError
'
' DEPENDENCIES
'   TEST_Harness; RunSSA_Silent, RunError, CheckOutcome, StatusLine, NamedText
'   and PlainText; the names NM_CHECKS_HEADER, NM_IR_CALC_HEADER,
'   NM_CHECKS_ERRORS, NM_FINGERPRINT and NM_LADDER_RW; ERR_LAYOUT
'
' NOTES
'   Rewrites the output sheets. Changes the first rule identifier of Checks
'   for one run, and the range of the name Ladder_RW for another, and puts
'   both back before it asserts
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target      As Range    'Cell changed to make a run fail
    Dim Saved       As Variant  'Its value before the test
    Dim Ladder      As Object   'Workbook name changed to make a run fail
    Dim SavedRange  As String   'What the name referred to
    Dim ErrNumber   As Long     'Error number of the failed run

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an error of the case as a failure of the case, and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "failed-run"

'------------------------------------------------------------------------------
' FAIL WHILE WRITING
'------------------------------------------------------------------------------
    'Run once without a failure, so that a result exists before the failure
        TEST_Harness.AssertEqualText "run-before-the-failure-while-writing", "RESULT=OK", _
            Left$(RunSSA_Silent(), 9)
    'While writing, after every result table: the first rule of Checks is not found under its header
        Set Target = ThisWorkbook.Names(NM_CHECKS_HEADER).RefersToRange.Offset(1, 0)
        Saved = Target.Value
        Target.Value = "X01"
    'Run the failing calculation and put the rule back before asserting
        ErrNumber = RunError()
        Target.Value = Saved
    'Assert the error raised
        TEST_Harness.AssertRaised "while-writing-error", ERR_LAYOUT, ErrNumber

'------------------------------------------------------------------------------
' ASSERT AFTER WRITING FAILURE
'------------------------------------------------------------------------------
    'Assert that the failed run left a status line that says FAILED and no result
        TEST_Harness.AssertEqualText "while-writing-status-line", "FAILED: ", Left$(StatusLine(), 8)
        TEST_Harness.AssertEqualText "while-writing-requirement", vbNullString, NamedText("SSA_Requirement")
        TEST_Harness.AssertEqualText "while-writing-first-calculation-row", vbNullString, _
            PlainText(ThisWorkbook.Names(NM_IR_CALC_HEADER).RefersToRange.Offset(1, 0).Value)
        TEST_Harness.AssertEqualText "while-writing-error-count-of-checks", vbNullString, _
            NamedText(NM_CHECKS_ERRORS)
        TEST_Harness.AssertEqualText "while-writing-fingerprint", vbNullString, NamedText(NM_FINGERPRINT)
    'Assert what the check returns
        TEST_Harness.AssertEqualText "while-writing-check", "RESULT=FAILED; cleanup=PASS", CheckOutcome()

'------------------------------------------------------------------------------
' FAIL WHILE READING
'------------------------------------------------------------------------------
    'Run once without a failure, so that a result exists before the failure
        TEST_Harness.AssertEqualText "run-before-the-failure-while-reading", "RESULT=OK", _
            Left$(RunSSA_Silent(), 9)
    'Before writing: a column of the maturity ladder is one row short, as the range of its name
        Set Ladder = ThisWorkbook.Names(NM_LADDER_RW)
        SavedRange = Ladder.RefersTo
        Set Target = Ladder.RefersToRange
        Ladder.RefersTo = "='" & Target.Worksheet.Name & "'!" & _
                          Target.Resize(Target.Rows.Count - 1, 1).Address
    'Run the failing calculation and put the name back before asserting
        ErrNumber = RunError()
        Ladder.RefersTo = SavedRange
    'Assert the error raised
        TEST_Harness.AssertRaised "while-reading-error", ERR_LAYOUT, ErrNumber

'------------------------------------------------------------------------------
' ASSERT AFTER READING FAILURE
'------------------------------------------------------------------------------
    'Assert that the failed run left a status line that says FAILED and no result
        TEST_Harness.AssertEqualText "while-reading-status-line", "FAILED: ", Left$(StatusLine(), 8)
        TEST_Harness.AssertEqualText "while-reading-requirement", vbNullString, NamedText("SSA_Requirement")
        TEST_Harness.AssertEqualText "while-reading-first-calculation-row", vbNullString, _
            PlainText(ThisWorkbook.Names(NM_IR_CALC_HEADER).RefersToRange.Offset(1, 0).Value)
        TEST_Harness.AssertEqualText "while-reading-error-count-of-checks", vbNullString, _
            NamedText(NM_CHECKS_ERRORS)
        TEST_Harness.AssertEqualText "while-reading-fingerprint", vbNullString, NamedText(NM_FINGERPRINT)
    'Assert what the check returns
        TEST_Harness.AssertEqualText "while-reading-check", "RESULT=FAILED; cleanup=PASS", CheckOutcome()

'------------------------------------------------------------------------------
' RUN AFTER FAILURES
'------------------------------------------------------------------------------
    'Assert that a run after the two failures succeeds again
        TEST_Harness.AssertEqualText "run-after-the-failures", "RESULT=OK", Left$(RunSSA_Silent(), 9)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record an error of the case as a failure of the case; the run continues
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseVolume()

'
'------------------------------------------------------------------------------
'                                 CASE VOLUME
'------------------------------------------------------------------------------
' PURPOSE
'   A position sheet that holds more rows than the workbook supports stops
'   the run: a PositionID on the row after the last supported one makes the
'   table 501 rows, and a run on exactly 500 rows is not stopped. The stop
'   leaves no result, counts one error on Checks and lists the sheet and the
'   rows found
'
' WHY THIS EXISTS
'   Shows that the limit of MAX_POSITION_ROWS rows of a position sheet is
'   enforced at the first row over it and not before
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case volume. Writes a PositionID on row 500 of IRPositions and
'   asserts that a run is not stopped, clears it, writes one on row 501 and
'   asserts that a run is stopped, clears it, then asserts the status line,
'   that no result is left, the error count of Checks and the sheet and the
'   reason of the first finding
'
' ERROR POLICY
'   An unexpected error is recorded as a failure of the case by
'   TEST_Harness.RecordCaseError; the run continues
'
' DEPENDENCIES
'   TEST_Harness; RunSSA_Silent, StatusLine, NamedText and PlainText; the names
'   NM_IR_POS_HEADER, NM_CHECKS_ERRORS and NM_FINDINGS_HEADER;
'   MAX_POSITION_ROWS, SH_IR_POS and ST_INVALID
'
' NOTES
'   Rewrites the output sheets. Writes one PositionID below the sample rows
'   of IRPositions, twice, and clears the cell before it asserts
'
' UPDATED
'   2026-11-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header      As Range    'First header cell of IRPositions
    Dim Findings    As Range    'Header of the Checks findings table

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an error of the case as a failure of the case, and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "volume"

'------------------------------------------------------------------------------
' RUN ON 500 ROWS
'------------------------------------------------------------------------------
    'Write a PositionID on the last supported row and assert that the run is not stopped
        Set Header = ThisWorkbook.Names(NM_IR_POS_HEADER).RefersToRange
        Header.Offset(MAX_POSITION_ROWS, 0).Value = "V500"
        TEST_Harness.AssertEqualText "run-on-500-rows-is-not-stopped", "RESULT=INCOMPLETE", _
            Left$(RunSSA_Silent(), 17)

'------------------------------------------------------------------------------
' RUN ON 501 ROWS
'------------------------------------------------------------------------------
    'Clear that cell, write a PositionID on the row after it and assert that the run is stopped
        Header.Offset(MAX_POSITION_ROWS, 0).ClearContents
        Header.Offset(MAX_POSITION_ROWS + 1, 0).Value = "V501"
        TEST_Harness.AssertEqualText "run-on-501-rows-is-stopped", "RESULT=STOPPED;", _
            Left$(RunSSA_Silent(), 15)
    'Clear the cell before asserting what the stop left
        Header.Offset(MAX_POSITION_ROWS + 1, 0).ClearContents

'------------------------------------------------------------------------------
' ASSERT STOP
'------------------------------------------------------------------------------
    'Assert the status line, that the stop left no result and that it counts one error on Checks
        TEST_Harness.AssertEqualText "status-line", "STOPPED: 1 unusable input(s)", Left$(StatusLine(), 28)
        TEST_Harness.AssertEqualText "no-result", vbNullString, NamedText("SSA_Requirement")
        TEST_Harness.AssertEqualText "error-count-of-checks", "1", NamedText(NM_CHECKS_ERRORS)
    'Assert that the findings list the sheet and the rows found
        Set Findings = ThisWorkbook.Names(NM_FINDINGS_HEADER).RefersToRange
        TEST_Harness.AssertEqualText "finding-sheet", SH_IR_POS, PlainText(Findings.Offset(1, 1).Value)
        TEST_Harness.AssertEqualText "finding-reason", _
            ST_INVALID & "501 rows found; the workbook supports at most 500", _
            PlainText(Findings.Offset(1, 6).Value)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record an error of the case as a failure of the case; the run continues
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseEligibility()

'
'------------------------------------------------------------------------------
'                              CASE ELIGIBILITY
'------------------------------------------------------------------------------
' PURPOSE
'   The eligibility of Articles 325a and 94 of the CRR on the months the
'   template ships: under the Basel regime nothing is assessed; under the CRR
'   the calculated columns and the result block of Eligibility are written,
'   and Checks counts the approach not met; a month end that is not the last
'   day of its month is counted as an invalid row
'
' WHY THIS EXISTS
'   Shows that a run writes the calculated columns and the result block of
'   Eligibility, which the template protects, and the counts of rules C15 and
'   C16 on Checks
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case eligibility and runs the calculation under the Basel
'   regime of the template. Keeps the regime, the total own funds and the
'   last month end, sets the CRR regime and total own funds, runs and asserts
'   the cells of the first and last month, the result block and the counts of
'   Checks; makes the last month end the day before and asserts again. Then
'   puts the three cells back and asserts it
'
' ERROR POLICY
'   An unexpected error is recorded as a failure of the case by
'   TEST_Harness.RecordCaseError, and the cells are put back on that path too;
'   the run continues
'
' DEPENDENCIES
'   TEST_Harness; RunSSA_Silent, NumberOf and PlainText; the names
'   NM_EL_HEADER, NM_EL_RESULTS, NM_CHECKS_HEADER, PRM_REGIME and
'   PRM_OWN_FUNDS; EL_INPUT_WIDTH, ELIGIBILITY_ROWS and ST_INVALID
'
'   Reference: calculation contract CALC-09. The sample months, from
'   tools/prototype/ssa_reference.py: the size of Article 325a in the last
'   month, 309 million long and 206 million short, is 515 million, above EUR
'   500 million; that of Article 94 is 43 million, below EUR 50 million in each
'   of the last 12 months
'
' NOTES
'   Rewrites the output sheets. Changes the regime, the total own funds and
'   the last month end of Eligibility, and puts them back
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header      As Range    'First header cell of Eligibility
    Dim Block       As Range    'Result block of Eligibility
    Dim Rules       As Range    'Header of the Checks rules
    Dim Regime      As Range    'Regime cell
    Dim OwnFunds    As Range    'TotalOwnFunds cell
    Dim LastMonth   As Range    'Month end of the last row of Eligibility
    Dim SavedRegime As Variant  'Regime before the case
    Dim SavedFunds  As Variant  'TotalOwnFunds before the case
    Dim SavedMonth  As Variant  'Last month end before the case
    Dim ResultText  As String   'RunSSA_Silent result line, not compared
    Dim RestoreDue  As Boolean  'True once the cells are kept
    Dim Restored    As Boolean  'True when the three cells hold their values again

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an error of the case as a failure of the case and go on with the cleanup
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "eligibility"
    'Locate the cells the case reads and changes
        Set Header = ThisWorkbook.Names(NM_EL_HEADER).RefersToRange
        Set Block = ThisWorkbook.Names(NM_EL_RESULTS).RefersToRange
        Set Rules = ThisWorkbook.Names(NM_CHECKS_HEADER).RefersToRange
        Set Regime = ThisWorkbook.Names(PRM_REGIME).RefersToRange
        Set OwnFunds = ThisWorkbook.Names(PRM_OWN_FUNDS).RefersToRange
        Set LastMonth = Header.Offset(ELIGIBILITY_ROWS, 0)

'------------------------------------------------------------------------------
' RUN UNDER BASEL
'------------------------------------------------------------------------------
    'Run under the Basel regime of the template: no approach is assessed and no row written
        ResultText = RunSSA_Silent()
        TEST_Harness.AssertEqualText "basel-status", "NOT ASSESSED: Basel regime", PlainText(Block.Cells(1, 1).Value)
        TEST_Harness.AssertEqualText "basel-writes-no-row", vbNullString, _
            PlainText(Header.Offset(1, EL_INPUT_WIDTH).Value)

'------------------------------------------------------------------------------
' KEEP CELLS
'------------------------------------------------------------------------------
    'Keep the regime, the total own funds and the last month end
        SavedRegime = Regime.Value
        SavedFunds = OwnFunds.Value
        SavedMonth = LastMonth.Value
    'From here the cells are put back, whichever path ends the case
        RestoreDue = True

'------------------------------------------------------------------------------
' RUN UNDER THE CRR
'------------------------------------------------------------------------------
    'Set the CRR regime, with total own funds, and run
        Regime.Value = "CRR"
        OwnFunds.Value = 10000000000#
        ResultText = RunSSA_Silent()
    'Assert the first and the last month of Article 325a
        TEST_Harness.AssertEqualText "first-month-status", "OK", PlainText(Header.Offset(1, EL_INPUT_WIDTH).Value)
        TEST_Harness.AssertEqualDouble "last-month-size-of-325a", 515000000#, _
            NumberOf(Header.Offset(ELIGIBILITY_ROWS, EL_INPUT_WIDTH + 1).Value)
        TEST_Harness.AssertEqualText "last-month-of-325a-not-met", "NOT MET", _
            PlainText(Header.Offset(ELIGIBILITY_ROWS, EL_INPUT_WIDTH + 3).Value)
    'Assert the result block: Article 325a not met, Article 94 met and free to restart
        TEST_Harness.AssertEqualText "status-of-325a", "NOT MET", PlainText(Block.Cells(1, 1).Value)
        TEST_Harness.AssertEqualText "status-of-94", "MET", PlainText(Block.Cells(2, 1).Value)
        TEST_Harness.AssertEqualText "restart-of-94", "YES", PlainText(Block.Cells(2, 7).Value)
    'Assert the counts of rules C15 and C16, on the 15th and 16th rows below the header, in its Count column
        TEST_Harness.AssertEqualText "no-row-counted-in-c15", "0", PlainText(Rules.Offset(15, 3).Value)
        TEST_Harness.AssertEqualText "approach-counted-in-c16", "1", PlainText(Rules.Offset(16, 3).Value)

'------------------------------------------------------------------------------
' RUN WITH AN INVALID MONTH
'------------------------------------------------------------------------------
    'Make the last month end the day before the end of its month, and run
        LastMonth.Value = SavedMonth - 1
        ResultText = RunSSA_Silent()
    'Assert the status of the row, the result block and the count of rule C15
        TEST_Harness.AssertEqualText "invalid-month-status", _
            ST_INVALID & "MonthEnd must be the last day of a month", _
            PlainText(Header.Offset(ELIGIBILITY_ROWS, EL_INPUT_WIDTH).Value)
        TEST_Harness.AssertEqualText "invalid-month-not-assessed", "NOT ASSESSED: a row is invalid", _
            PlainText(Block.Cells(1, 1).Value)
        TEST_Harness.AssertEqualText "invalid-month-shows-no-size", vbNullString, PlainText(Block.Cells(1, 3).Value)
        TEST_Harness.AssertEqualText "invalid-month-counted-in-c15", "1", PlainText(Rules.Offset(15, 3).Value)

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
CleanUp:
    'Stop catching errors, so that an error of the cleanup cannot come back here
        On Error GoTo 0
    'Put the three cells back, on the success and the error path; an empty cell is cleared
        If RestoreDue Then
            Regime.Value = SavedRegime
            If IsEmpty(SavedFunds) Then
                OwnFunds.ClearContents
            Else
                OwnFunds.Value = SavedFunds
            End If
            LastMonth.Value = SavedMonth
            'Compare the texts, whether each cell is empty, and the month end as a serial
            Restored = (PlainText(Regime.Value) = PlainText(SavedRegime)) And _
                       (IsEmpty(OwnFunds.Value) = IsEmpty(SavedFunds)) And _
                       (PlainText(OwnFunds.Value) = PlainText(SavedFunds)) And _
                       (CDbl(LastMonth.Value) = CDbl(SavedMonth))
        Else
            Restored = True
        End If
    'Assert once that the cells are put back, whichever path came here
        TEST_Harness.AssertTrue "inputs-restored", Restored, _
            "the regime, the total own funds and the last month end must hold their values again"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record an error of the case as a failure of the case, then clean up
CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub

Private Sub CaseOutdatedResults()

'
'------------------------------------------------------------------------------
'                            CASE OUTDATED RESULTS
'------------------------------------------------------------------------------
' PURPOSE
'   After a run, a changed parameter or position makes the check mark the
'   results OUTDATED; the value put back makes it take the mark away
'
' WHY THIS EXISTS
'   Shows that results that no longer match the inputs are marked, once, and
'   that the mark goes away when the input is put back
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case outdated-results and runs the calculation. Changes the IR
'   scaling factor, asserts the check and the status line, checks again, puts
'   the factor back and asserts again. Changes the value of the first interest
'   rate position, runs the status check through CheckStatus_OnEvent, puts the
'   value back and asserts again. Clears the outputs and asserts the check
'   without results
'
' ERROR POLICY
'   An unexpected error is recorded as a failure of the case by
'   TEST_Harness.RecordCaseError; the run continues
'
' DEPENDENCIES
'   TEST_Harness; RunSSA_Silent, CheckOutcome, StatusLine, CheckStatus_OnEvent
'   and ClearOutputs_Silent; the names ScalingIR and NM_IR_POS_HEADER; IR_VALUE
'
' NOTES
'   Rewrites the output sheets and ends with them cleared. Changes the IR
'   scaling factor and the value of the first interest rate position for
'   one check each, and puts both back
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target  As Range    'Input cell changed
    Dim Saved   As Variant  'Its value before the test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an error of the case as a failure of the case, and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "outdated-results"

'------------------------------------------------------------------------------
' RUN BEFORE CHANGES
'------------------------------------------------------------------------------
    'Run the calculation and assert that the check and the status line say the results are current
        TEST_Harness.AssertEqualText "run-before-the-changes", "RESULT=OK", Left$(RunSSA_Silent(), 9)
        TEST_Harness.AssertEqualText "check-after-the-run", "RESULT=CURRENT; cleanup=PASS", _
            CheckOutcome()
        TEST_Harness.AssertEqualText "status-after-the-run", "OK: ", Left$(StatusLine(), 4)

'------------------------------------------------------------------------------
' CHANGE PARAMETER
'------------------------------------------------------------------------------
    'A parameter: change the IR scaling factor
        Set Target = ThisWorkbook.Names("ScalingIR").RefersToRange
        Saved = Target.Value
        Target.Value = Saved + 0.1
    'Assert that the check marks the results OUTDATED, and marks them once
        TEST_Harness.AssertEqualText "check-after-a-parameter-changed", _
            "RESULT=OUTDATED; cleanup=PASS", CheckOutcome()
        TEST_Harness.AssertEqualText "status-after-a-parameter-changed", "OUTDATED: OK: ", _
            Left$(StatusLine(), 14)
        TEST_Harness.AssertEqualText "check-again-still-changed", "RESULT=OUTDATED; cleanup=PASS", _
            CheckOutcome()
        TEST_Harness.AssertEqualText "status-is-marked-once", "OUTDATED: OK: ", Left$(StatusLine(), 14)
    'Put the parameter back and assert that the check takes the mark away
        Target.Value = Saved
        TEST_Harness.AssertEqualText "check-after-the-parameter-is-put-back", _
            "RESULT=CURRENT; cleanup=PASS", CheckOutcome()
        TEST_Harness.AssertEqualText "status-after-the-parameter-is-put-back", "OK: ", _
            Left$(StatusLine(), 4)

'------------------------------------------------------------------------------
' CHANGE POSITION
'------------------------------------------------------------------------------
    'A position, through the event entry point: change the value of the first interest rate position
        Set Target = ThisWorkbook.Names(NM_IR_POS_HEADER).RefersToRange.Offset(1, IR_VALUE - 1)
        Saved = Target.Value
        Target.Value = Saved + 1
        CheckStatus_OnEvent
    'Assert that the status line is marked OUTDATED
        TEST_Harness.AssertEqualText "status-after-a-position-changed", "OUTDATED: OK: ", _
            Left$(StatusLine(), 14)
    'Put the position back and assert that the mark is taken away
        Target.Value = Saved
        CheckStatus_OnEvent
        TEST_Harness.AssertEqualText "status-after-the-position-is-put-back", "OK: ", _
            Left$(StatusLine(), 4)

'------------------------------------------------------------------------------
' CLEAR OUTPUTS
'------------------------------------------------------------------------------
    'No result to check; the output sheets stay cleared
        ClearOutputs_Silent
        TEST_Harness.AssertEqualText "check-without-results", _
            "RESULT=NOT_CALCULATED; cleanup=PASS", CheckOutcome()

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record an error of the case as a failure of the case; the run continues
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub AllowWritingToSheets()

'
'------------------------------------------------------------------------------
'                           ALLOW WRITING TO SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Let the suite write the sheets of its workbook, which the template
'   protects: apply UserInterfaceOnly to each of them, as AllowWriting does
'   for the sheets that RunSSA writes
'
' WHY THIS EXISTS
'   The cases write input cells and run the calculation, and the sheets of the
'   template are protected
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls AllowWriting for each worksheet of ThisWorkbook
'
' ERROR POLICY
'   None; an error of AllowWriting propagates to RunSuite
'
' DEPENDENCIES
'   AllowWriting
'
' NOTES
'   Changes the protection mode of every sheet, not who may change it
'
' UPDATED
'   2026-11-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet   As Worksheet    'One sheet

'------------------------------------------------------------------------------
' ALLOW WRITING
'------------------------------------------------------------------------------
    'Apply UserInterfaceOnly to each sheet of the workbook through AllowWriting
        For Each Sheet In ThisWorkbook.Worksheets
            AllowWriting Sheet
        Next Sheet

End Sub

Private Function RunError() As Long

'
'------------------------------------------------------------------------------
'                            RUN EXPECTING AN ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Run RunSSA_Silent where it is expected to raise an error
'
' WHY THIS EXISTS
'   A case changes an input to make the run fail, and has to put the input back
'   afterwards, so the error of the run cannot be left to end the case
'
' INPUTS
'   None
'
' RETURNS
'   The error number; 0 when the run raised none
'
' BEHAVIOR
'   Runs RunSSA_Silent with an error handler and returns Err.Number when it
'   raised
'
' ERROR POLICY
'   Contains the error of the run, so that the caller can put back what it
'   changed
'
' DEPENDENCIES
'   RunSSA_Silent
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultText  As String   'Result line, not used

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain the error that the run is expected to raise
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN CALCULATION
'------------------------------------------------------------------------------
    'Run the calculation and return 0 when it raised no error
        ResultText = RunSSA_Silent()
        RunError = 0

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Return the number of the error that the run raised
Failed:
        RunError = Err.Number

End Function

Private Function CheckOutcome() As String

'
'------------------------------------------------------------------------------
'                               RUN STATUS CHECK
'------------------------------------------------------------------------------
' PURPOSE
'   Run CheckStatus_Silent between a change of an input and its undoing
'
' WHY THIS EXISTS
'   A case changes an input before the check and has to put it back after it,
'   so the error of the check cannot be left to end the case
'
' INPUTS
'   None
'
' RETURNS
'   The result line, or "ERROR", the number and the description
'
' BEHAVIOR
'   Returns the result line of CheckStatus_Silent; when it raises, returns the
'   text ERROR, the error number and the error description, separated by
'   spaces
'
' ERROR POLICY
'   Contains the error of the check, so that the caller can put back what it
'   changed
'
' DEPENDENCIES
'   CheckStatus_Silent
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain the error of the check
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN CHECK
'------------------------------------------------------------------------------
    'Return the result line of the status check
        CheckOutcome = CheckStatus_Silent()

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Return ERROR, the number and the description of the error of the check
Failed:
        CheckOutcome = "ERROR " & Err.Number & " " & Err.Description

End Function

Private Function StatusLine() As String

'
'------------------------------------------------------------------------------
'                               READ STATUS LINE
'------------------------------------------------------------------------------
' PURPOSE
'   Return the status line of Results
'
' WHY THIS EXISTS
'   The cases read the status line in many assertions
'
' INPUTS
'   None
'
' RETURNS
'   The text of the status line; empty for an empty cell
'
' BEHAVIOR
'   Returns NamedText of the name NM_RESULTS_STATUS
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   NamedText; NM_RESULTS_STATUS
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ STATUS LINE
'------------------------------------------------------------------------------
    'Return the text of the cell named NM_RESULTS_STATUS
        StatusLine = NamedText(NM_RESULTS_STATUS)

End Function

Private Function NamedText( _
    ByVal CellName As String) _
    As String
'
'------------------------------------------------------------------------------
'                             READ NAMED CELL TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Return the text of the cell a workbook name refers to; empty for an
'   empty cell
'
' WHY THIS EXISTS
'   The cases compare the text of named cells with expected text
'
' INPUTS
'   CellName
'     A workbook name that refers to one cell
'
' RETURNS
'   The text of the cell, as PlainText gives it
'
' BEHAVIOR
'   Returns PlainText of the value of the range that the name refers to
'
' ERROR POLICY
'   None; a name that does not exist or does not refer to a range raises the
'   error of the host, which the calling case records
'
' DEPENDENCIES
'   CORE_Util.PlainText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ CELL
'------------------------------------------------------------------------------
    'Return the text of the cell that the name refers to, empty for an empty cell
        NamedText = PlainText(ThisWorkbook.Names(CellName).RefersToRange.Value)

End Function

Private Function NumberOf( _
    ByVal Value As Variant) _
    As Double
'
'------------------------------------------------------------------------------
'                           CONVERT VALUE TO NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   The number a cell holds, for an assertion that compares numbers
'
' WHY THIS EXISTS
'   An empty cell, a text or an error value must never be compared as 0
'
' INPUTS
'   Value
'     The value of a cell
'
' RETURNS
'   The value as a Double
'
' BEHAVIOR
'   Raises error 13 when IsNumberValue rejects the value; otherwise converts it
'   with CDbl
'
' ERROR POLICY
'   Raises error 13 when Value is not a number, so that an empty cell, a text
'   or an error value is never compared as 0; the case records it
'
' DEPENDENCIES
'   CORE_Util.IsNumberValue
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Reject a value that is not a number, so that it is never compared as 0
        If Not IsNumberValue(Value) Then
            Err.Raise 13, "TEST_Workbook.NumberOf", "The result is not a number."
        End If

'------------------------------------------------------------------------------
' RETURN NUMBER
'------------------------------------------------------------------------------
    'Return the value as a Double
        NumberOf = CDbl(Value)

End Function
