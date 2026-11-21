Attribute VB_Name = "TEST_Examples"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Examples
'==============================================================================
' PURPOSE
'   Neutral example suite of TEST_Harness: one case per kind of assertion,
'   one case that records the results of the illustrative test cases, and
'   the reference implementation of the cleanup pattern of
'   docs/VBA_HOUSE_STYLE.md (Operation and cleanup outcomes) with its
'   failure-path cases. Nothing here is a calculation of the workbook or a
'   regulatory value
'
' WHY THIS EXISTS
'   An Excel run of this suite in its three modes is the evidence that the
'   harness itself works on the host, and the suite is the worked example of
'   how a project suite is written. It also holds the reference
'   implementation of the cleanup pattern, which the failure-path cases
'   exercise: capture each setting, register its restoration, restore on both
'   paths, keep the primary error and report every cleanup failure
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides four entry points:
'     - RunExampleTests runs the suite in mode NORMAL
'     - RunExampleTestsWithInjectedFailure runs it in mode INJECTED_FAILURE
'       and RunExampleTestsWithSkippedCase in mode SKIPPED_CASE, to show what
'       a failed and an incomplete run report
'     - CertifyExamples certifies all three modes and saves their logs, the
'       certification log and the results of the normal run as the evidence of
'       an Excel run
'   Each takes the folder of the test cases, tests/fixtures/ of the checkout.
'
'   RunSuite runs fourteen cases in a fixed order between BeginRun and
'   EndRun:
'     - six cases, one per kind of assertion: exact double, absolute
'       tolerance, relative tolerance, exact text, error value and raised
'       error
'     - one case that records the results of the illustrative test cases
'     - seven cases of the cleanup pattern: success, primary error kept,
'       partial capture, change failure, primary error only, each restoration
'       failure and cleanup-only failure
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START. CertifyExamples records those errors
'   instead of raising them, saves every log, and raises only when the
'   certification did not pass
'
' DEPENDENCIES
'   TEST_Harness; the VBA and Excel object models
'
' NOTES
'   Public surface: RunExampleTests is the documented test entry point.
'   Option Private Module keeps the entry points out of the workbook's
'   external surface, and Option Explicit requires explicit declarations
'
'   State: owns the mode and the fixtures folder of the running suite. Reads
'   the illustrative test cases of tests/fixtures/. RunExampleOperation
'   changes Application.Calculation and Application.ScreenUpdating and
'   restores them before it returns, on the success and the error path
'
'   Usage: in a development workbook holding the modules of tests/modules/,
'   type TEST_Examples.RunExampleTests "<checkout>\tests\fixtures" in the
'   Immediate window. To record an Excel run, type
'   TEST_Examples.CertifyExamples "<folder>", "<checkout>\tests\fixtures"
'   instead (INSTALLATION.md, Recording an Excel run)
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-04
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "examples"  'Suite identifier in the RUN records
    Private Const EXPECTED_CASES        As Long = 14            'Cases of a complete run
    Private Const EXPECTED_ASSERTIONS   As Long = 34            'Assertions of a complete run

    'Test-only errors of the example procedures, literals equal to vbObjectError + n (see TEST_Harness)
    Private Const ERR_EXAMPLE_INVALID   As Long = &H80040FF0    'vbObjectError + 4080: Input outside the domain
    Private Const ERR_EXAMPLE_OPERATION As Long = &H80040FF1    'vbObjectError + 4081: Primary error of the operation
    Private Const ERR_EXAMPLE_RESTORE   As Long = &H80040FF2    'vbObjectError + 4082: Simulated restore failure
    Private Const ERR_CLEANUP_FAILED    As Long = &H80040FF3    'vbObjectError + 4083: Only the cleanup failed
    Private Const ERR_EXAMPLE_HOST      As Long = &H80040FF4    'vbObjectError + 4084: Simulated capture/change failure
    Private Const ERR_EXAMPLE_FIXTURE   As Long = &H80040FF5    'vbObjectError + 4085: Test case file not as expected

    'The quantity of the illustrative test cases (tests/expected/)
    Private Const ILLUSTRATIVE_QUANTITY As String = "illustrative_total_notional_amount"    'Quantity of each total

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mMode       As String   'NORMAL, INJECTED_FAILURE or SKIPPED_CASE
    Private mFixtures   As String   'Folder of the test cases, tests/fixtures/


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunExampleTests( _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                              RUN EXAMPLE TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the example suite in mode NORMAL
'
' WHY THIS EXISTS
'   It is the documented test entry point of the example suite
'
' INPUTS
'   Fixtures
'     Full path of tests/fixtures/ in the checkout, without a trailing
'     separator
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs every case through RunSuite. A passing run ends with
'   RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE;
'   cases=14; assertions=34; failures=0; cleanup=PASS; recovery=NONE.
'
' ERROR POLICY
'   Raises what RunSuite raises: ERR_TEST_FAILED for a run that does not
'   pass, or the unexpected error that escaped the runner
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run every case in the normal mode
        RunSuite "NORMAL", Fixtures

End Sub

Public Sub RunExampleTestsWithInjectedFailure( _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                          RUN WITH INJECTED FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Shows the failure path: the exact-double case expects a wrong value
'
' WHY THIS EXISTS
'   Shows what a failed run reports, so that a failure is never silent
'
' INPUTS
'   Fixtures
'     As RunExampleTests
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs every case through RunSuite in mode INJECTED_FAILURE. Expected
'   end: RESULT=FAIL; ...; operation=FAIL; completeness=COMPLETE; cases=14;
'   assertions=34; failures=1, then the error ERR_TEST_FAILED.
'
' ERROR POLICY
'   Raises ERR_TEST_FAILED, as designed
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run every case with the exact-double expectation made wrong
        RunSuite "INJECTED_FAILURE", Fixtures

End Sub

Public Sub RunExampleTestsWithSkippedCase( _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                            RUN WITH SKIPPED CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Shows that an incomplete run is never a pass: the last case is skipped
'
' WHY THIS EXISTS
'   Shows what an incomplete run reports, so that a missing case is never
'   silent
'
' INPUTS
'   Fixtures
'     As RunExampleTests
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the cases through RunSuite in mode SKIPPED_CASE. Expected end:
'   RESULT=FAIL; ...; operation=PASS; completeness=INCOMPLETE; cases=13;
'   assertions=32; failures=1, the failure being harness.completeness, then
'   ERR_TEST_FAILED.
'
' ERROR POLICY
'   Raises ERR_TEST_FAILED, as designed
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run every case except the last one
        RunSuite "SKIPPED_CASE", Fixtures

End Sub

Public Sub CertifyExamples( _
    ByVal Folder As String, _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                               CERTIFY EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   The certification entry point: refuse state left by an earlier run, run
'   the suite in its three modes, save each complete log in Folder
'   (examples.log, injected.log and skipped.log) and the results of the
'   normal run (results.json), then save certification.log
'
' WHY THIS EXISTS
'   tools/check_excel_run.py checks the folder against the commit, so the
'   certification is the evidence of an Excel run
'
' INPUTS
'   Folder
'     Existing folder outside the checkout or in the ignored test-results/
'     directory, without a trailing backslash
'
'   Fixtures
'     Full path of tests/fixtures/ in the same checkout, without a trailing
'     separator
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   The normal run must pass and be complete. The injected-failure and
'   skipped-case runs must fail and reach this caller as
'   TEST_Harness.ERR_TEST_FAILED with a description: they are the
'   self-test that a failed run is never silent.
'
' ERROR POLICY
'   The error that ends each run is recorded, not raised, so that every log
'   is written. EndCertification raises ERR_TEST_FAILED after writing
'   certification.log when the certification did not pass. A refused start
'   raises, before any run, TEST_Harness.ERR_TEST_DIRTY_START when a run is
'   active or recovery is required, or TEST_Harness.ERR_TEST_USAGE while an
'   earlier certification is still active. An error of SaveLog or
'   SaveResults propagates; run TEST_Harness.ResetTests before trying again
'
' DEPENDENCIES
'   TEST_Harness (BeginCertification, SaveResults, EndCertification), CertifyMode
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' BEGIN CERTIFICATION
'------------------------------------------------------------------------------
    'Refuse state left by an earlier run and open the certification of the suite
        TEST_Harness.BeginCertification SUITE_NAME

'------------------------------------------------------------------------------
' CERTIFY RUNS
'------------------------------------------------------------------------------
    'Certify the normal run, which must pass, and save its results
        CertifyMode "NORMAL", True, Folder & "\examples.log", Fixtures
        TEST_Harness.SaveResults Folder & "\results.json"
    'Certify the two deliberate failure modes, which must fail
        CertifyMode "INJECTED_FAILURE", False, Folder & "\injected.log", Fixtures
        CertifyMode "SKIPPED_CASE", False, Folder & "\skipped.log", Fixtures

'------------------------------------------------------------------------------
' END CERTIFICATION
'------------------------------------------------------------------------------
    'Save the certification log and raise when the certification did not pass
        TEST_Harness.EndCertification Folder & "\certification.log"

End Sub

'
'------------------------------------------------------------------------------
'
'                                    RUNNER
'
'------------------------------------------------------------------------------
'

Private Sub RunSuite( _
    ByVal Mode As String, _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                                  RUN SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every case in a fixed order between BeginRun and EndRun
'
' WHY THIS EXISTS
'   The three public entry points and the certification differ only in the
'   mode, and share this runner
'
' INPUTS
'   Mode
'     NORMAL, INJECTED_FAILURE or SKIPPED_CASE
'
'   Fixtures
'     Folder of the test cases, as RunExampleTests
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the mode and the fixtures folder, begins the run with the expected
'   counts, calls the cases in order (the last one only when the mode is not
'   SKIPPED_CASE), clears the stored state and ends the run
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it
'   and re-raises it unchanged. BeginRun's refusal propagates directly
'
' DEPENDENCIES
'   TEST_Harness (BeginRun, EndRun) and the Case procedures of this module
'
' NOTES
'   Sets mMode and mFixtures, which it clears before EndRun
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' START RUN
'------------------------------------------------------------------------------
    'Keep the mode and the fixtures folder for the cases
        mMode = Mode
        mFixtures = Fixtures
    'Begin the run with the expected counts; a refusal propagates directly
        TEST_Harness.BeginRun SUITE_NAME, Mode, EXPECTED_CASES, EXPECTED_ASSERTIONS

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Catch an error that escapes a case handler
        On Error GoTo RunFailed

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Run the cases in their fixed order
        CaseExactDouble
        CaseAbsoluteTolerance
        CaseRelativeTolerance
        CaseExactText
        CaseErrorValue
        CaseRaisedError
        CaseIllustrativeResults
        CaseCleanupRestoresState
        CaseCleanupKeepsPrimaryError
        CaseCleanupPartialCapture
        CaseCleanupChangeFailure
        CaseCleanupPrimaryOnly
        CaseCleanupEachRestorationFailure
    'Leave the last case out in the skipped-case mode
        If mMode <> "SKIPPED_CASE" Then
            CaseCleanupOnlyFailure
        End If

'------------------------------------------------------------------------------
' END RUN
'------------------------------------------------------------------------------
Finish:
    'Restore normal error handling
        On Error GoTo 0
    'Clear the stored state
        mMode = vbNullString
        mFixtures = vbNullString
    'End the run, which reports the outcome and raises it unless the run passed
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
RunFailed:
    'Save the runner error before any other call
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'End the run with the saved error
        Resume Finish

End Sub

Private Sub CertifyMode( _
    ByVal Mode As String, _
    ByVal MustPass As Boolean, _
    ByVal Path As String, _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                                 CERTIFY MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the suite in Mode, saves its log to Path, then certifies how the
'   run reached this caller
'
' WHY THIS EXISTS
'   A failing mode ends in an error that CertifyExamples must not let
'   escape before every log is saved
'
' INPUTS
'   Mode
'     NORMAL, INJECTED_FAILURE or SKIPPED_CASE
'
'   MustPass
'     True for the normal run, False for a deliberate failure mode
'
'   Path
'     File that receives the log of the run
'
'   Fixtures
'     Folder of the test cases, as RunExampleTests
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the suite with errors suspended and keeps the number and the
'   description of the error that ended it, saves the log, echoes the path,
'   the mode and the error, and certifies the run with CertifyRun
'
' ERROR POLICY
'   The run's own error is captured, printed and certified after the log is
'   saved; an error of SaveLog propagates
'
' DEPENDENCIES
'   RunSuite and TEST_Harness (SaveLog, Echo, CertifyRun)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RunErrorNumber      As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription As String   'Its description

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run the suite, whose error is captured instead of raised
        On Error Resume Next
        RunSuite Mode, Fixtures
    'Keep the error that ended the run
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' SAVE AND CERTIFY
'------------------------------------------------------------------------------
    'Save the log of the run
        TEST_Harness.SaveLog Path
    'Echo where the log went and how the run ended
        TEST_Harness.Echo "SAVED=" & Path & "; mode=" & Mode & "; error=" & CStr(RunErrorNumber) & _
            IIf(RunErrorNumber = 0, vbNullString, " " & RunErrorDescription)
    'Certify how the run reached this caller
        TEST_Harness.CertifyRun Mode, MustPass, RunErrorNumber, RunErrorDescription

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseExactDouble()

'
'------------------------------------------------------------------------------
'                              CASE EXACT DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Exact equality on a result that is exact in binary. 1 assertion.
'
' WHY THIS EXISTS
'   Shows the assertion of exact equality for a Double, where the result is
'   exact in binary and needs no tolerance
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Expects 0.75, or the wrong value 0.76 when mMode is INJECTED_FAILURE,
'   and asserts that it equals 0.5 + 0.25
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertEqualDouble, RecordCaseError); mMode
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Expected    As Double   'Expected sum; wrong on purpose when injected

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "exact-double"

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Expect 0.75, and a wrong value when the failure is injected
        Expected = 0.75
        If mMode = "INJECTED_FAILURE" Then
            Expected = 0.76
        End If
    'Assert exact equality with a sum that is exact in binary
        TEST_Harness.AssertEqualDouble "half-plus-quarter", Expected, 0.5 + 0.25

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseAbsoluteTolerance()

'
'------------------------------------------------------------------------------
'                           CASE ABSOLUTE TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   The absolute bound: 0.1 + 0.2 is not exactly 0.3 in binary, and near
'   zero only the absolute bound counts. 2 assertions.
'
' WHY THIS EXISTS
'   Shows that the sum 0.1 + 0.2 differs from 0.3 in binary, and that
'   near zero only the absolute bound counts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Asserts 0.1 + 0.2 near 0.3 with an absolute tolerance of 1E-15 and no
'   relative tolerance, and 1E-20 near 0 with an absolute tolerance of 1E-12
'   and a relative tolerance of 1E-9
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertNear, RecordCaseError)
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "tolerance-absolute"

'------------------------------------------------------------------------------
' ASSERT ABSOLUTE BOUND
'------------------------------------------------------------------------------
    'Compare 0.1 + 0.2 with 0.3 within the absolute bound only
        TEST_Harness.AssertNear "tenth-plus-fifth", 0.3, 0.1 + 0.2, 0.000000000000001, 0#
    'Compare a value near zero, where a relative bound of 1E-9 would be 0
        TEST_Harness.AssertNear "near-zero", 0#, 1E-20, 0.000000000001, 0.000000001

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseRelativeTolerance()

'
'------------------------------------------------------------------------------
'                           CASE RELATIVE TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   The relative bound governs large values, the bound is inclusive, and a
'   value outside both bounds is rejected. 3 assertions.
'
' WHY THIS EXISTS
'   Shows the relative bound on a large value, that a difference equal to
'   the bound passes, and that a value outside both bounds fails
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Asserts 1000000000.1 near 1000000000 with a relative tolerance of 1E-9,
'   asserts that |2.5 - 2| = 0.5 is within an absolute tolerance of 0.5, and
'   asserts that |1.1 - 1| = 0.1 is not within 0.01 and 0.01 * 1
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertNear, AssertTrue, WithinTolerance,
'   RecordCaseError)
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "tolerance-relative"

'------------------------------------------------------------------------------
' ASSERT RELATIVE BOUND
'------------------------------------------------------------------------------
    'Accept a large value whose difference is within the relative bound
        TEST_Harness.AssertNear "large-value", 1000000000#, 1000000000.1, 0.000001, 0.000000001
    'Show that a difference equal to the bound is within it
        TEST_Harness.AssertTrue "bound-is-inclusive", _
            TEST_Harness.WithinTolerance(2#, 2.5, 0.5, 0#), "|2.5 - 2| = 0.5 must be within 0.5"
    'Show that a difference outside both bounds is rejected
        TEST_Harness.AssertTrue "outside-both-bounds", _
            Not TEST_Harness.WithinTolerance(1#, 1.1, 0.01, 0.01), _
            "|1.1 - 1| = 0.1 must exceed max(0.01, 0.01 * 1)"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseExactText()

'
'------------------------------------------------------------------------------
'                               CASE EXACT TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Text is compared exactly, case included. 2 assertions.
'
' WHY THIS EXISTS
'   Shows the assertion of exact text equality, which is case sensitive
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Asserts that "CASE" & "-" & CStr(1) equals "CASE-1", and that "a" and
'   "A" differ in a binary comparison
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertEqualText, AssertTrue, RecordCaseError)
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "text-exact"

'------------------------------------------------------------------------------
' ASSERT TEXT
'------------------------------------------------------------------------------
    'Assert that a built text equals the expected text
        TEST_Harness.AssertEqualText "concatenation", "CASE-1", "CASE" & "-" & CStr(1)
    'Show that the comparison is case sensitive
        TEST_Harness.AssertTrue "case-matters", _
            StrComp("a", "A", vbBinaryCompare) <> 0, "a and A must differ"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseErrorValue()

'
'------------------------------------------------------------------------------
'                               CASE ERROR VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   An expected worksheet error value, returned rather than raised.
'   1 assertion.
'
' WHY THIS EXISTS
'   Shows the assertion of a worksheet error value that a procedure
'   returns, as a worksheet function does, instead of raising it
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Asserts that ExampleRatio(1, 0) returns the error value #DIV/0!
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertErrorValue, RecordCaseError);
'   ExampleRatio
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "error-value"

'------------------------------------------------------------------------------
' ASSERT ERROR VALUE
'------------------------------------------------------------------------------
    'Expect the worksheet error #DIV/0! for a zero denominator
        TEST_Harness.AssertErrorValue "zero-denominator", xlErrDiv0, ExampleRatio(1#, 0#)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseRaisedError()

'
'------------------------------------------------------------------------------
'                              CASE RAISED ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   An expected raised error, captured on the line after the call, and the
'   same call with a valid input raising nothing. 2 assertions.
'
' WHY THIS EXISTS
'   Shows how a test captures an error that a procedure raises, and that
'   the same call with a valid input raises none
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls ExampleSquareRoot(-1) with errors suspended and asserts that it
'   raised ERR_EXAMPLE_INVALID, then calls ExampleSquareRoot(4) and asserts
'   that it raised nothing
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertRaised, RecordCaseError);
'   ExampleSquareRoot
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result      As Double   'Value of the checked square root
    Dim RaisedError As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "raised-error"

'------------------------------------------------------------------------------
' CHECK NEGATIVE INPUT
'------------------------------------------------------------------------------
    'Call with a negative value, which must raise
        On Error Resume Next
        Result = ExampleSquareRoot(-1#)
    'Capture the error number on the line after the call
        RaisedError = Err.Number
    'Restore the case handler
        On Error GoTo CaseFailed
    'Assert the error of the domain
        TEST_Harness.AssertRaised "negative-input", ERR_EXAMPLE_INVALID, RaisedError

'------------------------------------------------------------------------------
' CHECK VALID INPUT
'------------------------------------------------------------------------------
    'Call with a valid value, which must raise nothing
        On Error Resume Next
        Result = ExampleSquareRoot(4#)
    'Capture the error number on the line after the call
        RaisedError = Err.Number
    'Restore the case handler
        On Error GoTo CaseFailed
    'Assert that no error was raised
        TEST_Harness.AssertRaised "valid-input", 0, RaisedError

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseIllustrativeResults()

'
'------------------------------------------------------------------------------
'                          CASE ILLUSTRATIVE RESULTS
'------------------------------------------------------------------------------
' PURPOSE
'   The results of an Excel run: read the illustrative test cases from the
'   fixtures folder, add up the notional amounts of each, and record each
'   total with TEST_Harness.RecordResult, for the accuracy contract ACC-01
'   (tests/accuracy/contracts.json). tools/check_accuracy.py compares the
'   totals with their expected values; this case checks only that each file
'   held amounts. 3 assertions.
'
' WHY THIS EXISTS
'   Records the results of an Excel run for the accuracy contract, which
'   tools/check_accuracy.py compares with the expected values of the test
'   cases
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds up the notional amounts of the JSON example and of the CRLF and LF
'   variants of the file-import example, asserts that each file held
'   amounts, and records each total under the quantity
'   ILLUSTRATIVE_QUANTITY
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertTrue, RecordResult, RecordCaseError);
'   FixturePath, SumJsonNumbers, SumCsvColumn
'
' NOTES
'   The total of a file is a sum of amounts read as text, not a calculation of the workbook
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Total   As Double   'Sum of the notional amounts of one file
    Dim Records As Long     'Amounts read from that file

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "illustrative-results"

'------------------------------------------------------------------------------
' RECORD THE JSON EXAMPLE
'------------------------------------------------------------------------------
    'Add up the notional amounts of the JSON example
        Total = SumJsonNumbers(FixturePath("illustrative-format-example.json"), _
            "notional_amount", Records)
    'Require that the file held amounts
        TEST_Harness.AssertTrue "format-example-amounts", Records > 0, _
            "illustrative-format-example.json holds no notional_amount"
    'Record the total for the accuracy contract
        TEST_Harness.RecordResult "illustrative-format-example", vbNullString, _
            ILLUSTRATIVE_QUANTITY, vbNullString, vbNullString, Total

'------------------------------------------------------------------------------
' RECORD THE CRLF VARIANT
'------------------------------------------------------------------------------
    'Add up the notional amounts of the file with CR LF line endings
        Total = SumCsvColumn(FixturePath( _
            "illustrative-file-import-example/crlf-line-endings/positions.csv"), "notional", Records)
    'Require that the file held amounts
        TEST_Harness.AssertTrue "crlf-variant-amounts", Records > 0, _
            "the crlf-line-endings positions.csv holds no notional amount"
    'Record the total for the accuracy contract
        TEST_Harness.RecordResult "illustrative-file-import-example", "crlf-line-endings", _
            ILLUSTRATIVE_QUANTITY, vbNullString, vbNullString, Total

'------------------------------------------------------------------------------
' RECORD THE LF VARIANT
'------------------------------------------------------------------------------
    'Add up the notional amounts of the file with LF line endings
        Total = SumCsvColumn(FixturePath( _
            "illustrative-file-import-example/lf-line-endings/positions.csv"), "notional", Records)
    'Require that the file held amounts
        TEST_Harness.AssertTrue "lf-variant-amounts", Records > 0, _
            "the lf-line-endings positions.csv holds no notional amount"
    'Record the total for the accuracy contract
        TEST_Harness.RecordResult "illustrative-file-import-example", "lf-line-endings", _
            ILLUSTRATIVE_QUANTITY, vbNullString, vbNullString, Total

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCleanupRestoresState()

'
'------------------------------------------------------------------------------
'                         CASE CLEANUP RESTORES STATE
'------------------------------------------------------------------------------
' PURPOSE
'   A successful operation restores what it changed and reports no cleanup
'   failure. 2 assertions.
'
' WHY THIS EXISTS
'   Shows the success path of the cleanup pattern: nothing is reported and
'   the host setting is back
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation with no failure injected and asserts that it
'   reported no cleanup failure and that ScreenUpdating has its value from
'   before the operation
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertEqualText, AssertTrue, RecordCaseError);
'   RunExampleOperation
'
' NOTES
'   Changes no setting itself; it reads Application.ScreenUpdating
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details     As String   'Cleanup failures reported by the operation
    Dim SavedScreen As Boolean  'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "cleanup-restores-state"

'------------------------------------------------------------------------------
' RUN OPERATION
'------------------------------------------------------------------------------
    'Capture ScreenUpdating before the operation
        SavedScreen = Application.ScreenUpdating
    'Run the operation with no failure injected
        RunExampleOperation vbNullString, False, False, Details

'------------------------------------------------------------------------------
' ASSERT CLEANUP
'------------------------------------------------------------------------------
    'Expect no cleanup failure
        TEST_Harness.AssertEqualText "no-cleanup-failure", vbNullString, Details
    'Expect ScreenUpdating restored
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCleanupKeepsPrimaryError()

'
'------------------------------------------------------------------------------
'                       CASE CLEANUP KEEPS PRIMARY ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   When the operation and its cleanup both fail, the caller receives the
'   primary error, and the cleanup failure is still reported. 3 assertions.
'
' WHY THIS EXISTS
'   Shows that a cleanup failure never replaces the primary error and is
'   still reported
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation with a failure at "work" and both
'   restorations failing, and asserts that the error raised is
'   ERR_EXAMPLE_OPERATION and that the details name both restorations
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertRaised, AssertTrue, RecordCaseError);
'   RunExampleOperation
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details     As String   'Cleanup failures reported by the operation
    Dim RaisedError As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "cleanup-keeps-primary-error"

'------------------------------------------------------------------------------
' RUN OPERATION
'------------------------------------------------------------------------------
    'Run with a primary failure and both restorations failing
        On Error Resume Next
        RunExampleOperation "work", True, True, Details
    'Capture the error number on the line after the call
        RaisedError = Err.Number
    'Restore the case handler
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' ASSERT OUTCOMES
'------------------------------------------------------------------------------
    'Expect the primary error, not the cleanup error
        TEST_Harness.AssertRaised "primary-error", ERR_EXAMPLE_OPERATION, RaisedError
    'Expect the Calculation restoration failure to be reported
        TEST_Harness.AssertTrue "calculation-failure-reported", _
            InStr(Details, "Calculation not restored") > 0, "details='" & Details & "'"
    'Expect the ScreenUpdating restoration failure to be reported
        TEST_Harness.AssertTrue "screen-failure-reported", _
            InStr(Details, "ScreenUpdating not restored") > 0, "details='" & Details & "'"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCleanupPartialCapture()

'
'------------------------------------------------------------------------------
'                         CASE CLEANUP PARTIAL CAPTURE
'------------------------------------------------------------------------------
' PURPOSE
'   When capture fails part-way, nothing is restored: no setting was
'   changed, and an uncaptured value is never restored. Both restore seams
'   are set, so any restoration attempt would be reported. 2 assertions.
'
' WHY THIS EXISTS
'   Shows that a setting whose capture did not happen is never restored
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation with a failure at "capture-screen" and both
'   restore seams set, and asserts that the error raised is
'   ERR_EXAMPLE_HOST and that no cleanup detail was reported
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertRaised, AssertEqualText, RecordCaseError);
'   RunExampleOperation
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details     As String   'Cleanup failures reported by the operation
    Dim RaisedError As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "cleanup-partial-capture"

'------------------------------------------------------------------------------
' RUN OPERATION
'------------------------------------------------------------------------------
    'Run with a failure capturing ScreenUpdating and both restore seams set
        On Error Resume Next
        RunExampleOperation "capture-screen", True, True, Details
    'Capture the error number on the line after the call
        RaisedError = Err.Number
    'Restore the case handler
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' ASSERT OUTCOMES
'------------------------------------------------------------------------------
    'Expect the error of the capture
        TEST_Harness.AssertRaised "capture-error", ERR_EXAMPLE_HOST, RaisedError
    'Expect that no restoration was attempted
        TEST_Harness.AssertEqualText "nothing-restored", vbNullString, Details

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCleanupChangeFailure()

'
'------------------------------------------------------------------------------
'                         CASE CLEANUP CHANGE FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   When a change fails after its restoration was registered, that setting
'   is still restored, as is the one changed before it, and the change's
'   error stays the primary error. 3 assertions.
'
' WHY THIS EXISTS
'   Shows that a restoration registered before its change is performed even
'   when the change fails part-way
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation with a failure at "change-calculation" and the
'   Calculation restore seam set, and asserts the error raised, that the
'   Calculation restoration was attempted and that ScreenUpdating has its
'   value from before the operation
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertRaised, AssertTrue, RecordCaseError);
'   RunExampleOperation
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details     As String   'Cleanup failures reported by the operation
    Dim RaisedError As Long     'Err.Number right after the call
    Dim SavedScreen As Boolean  'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "cleanup-change-failure"

'------------------------------------------------------------------------------
' RUN OPERATION
'------------------------------------------------------------------------------
    'Capture ScreenUpdating before the operation
        SavedScreen = Application.ScreenUpdating
    'Run with a failure changing Calculation and its restore seam set
        On Error Resume Next
        RunExampleOperation "change-calculation", True, False, Details
    'Capture the error number on the line after the call
        RaisedError = Err.Number
    'Restore the case handler
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' ASSERT OUTCOMES
'------------------------------------------------------------------------------
    'Expect the error of the change
        TEST_Harness.AssertRaised "change-error", ERR_EXAMPLE_HOST, RaisedError
    'Expect the restoration of Calculation to have been attempted
        TEST_Harness.AssertTrue "calculation-restore-attempted", _
            InStr(Details, "Calculation not restored") > 0, "details='" & Details & "'"
    'Expect ScreenUpdating restored
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCleanupPrimaryOnly()

'
'------------------------------------------------------------------------------
'                          CASE CLEANUP PRIMARY ONLY
'------------------------------------------------------------------------------
' PURPOSE
'   When only the operation fails and every restoration succeeds, the
'   caller receives the primary error, no cleanup failure is reported and
'   both settings are restored: a failed operation stays failed even if its
'   cleanup succeeded. 4 assertions.
'
' WHY THIS EXISTS
'   Shows that a failed operation stays failed when its cleanup succeeded,
'   and that both settings are back
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation with a failure at "work" and no restoration
'   failing, and asserts the error raised, that no cleanup failure was
'   reported and that Calculation and ScreenUpdating have their values from
'   before the operation
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertRaised, AssertEqualText, AssertTrue,
'   RecordCaseError); RunExampleOperation
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details             As String           'Cleanup failures reported by the operation
    Dim RaisedError         As Long             'Err.Number right after the call
    Dim SavedCalculation    As XlCalculation    'Calculation before the operation
    Dim SavedScreen         As Boolean          'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "cleanup-primary-only"

'------------------------------------------------------------------------------
' RUN OPERATION
'------------------------------------------------------------------------------
    'Capture Calculation and ScreenUpdating before the operation
        SavedCalculation = Application.Calculation
        SavedScreen = Application.ScreenUpdating
    'Run with a primary failure and every restoration succeeding
        On Error Resume Next
        RunExampleOperation "work", False, False, Details
    'Capture the error number on the line after the call
        RaisedError = Err.Number
    'Restore the case handler
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' ASSERT OUTCOMES
'------------------------------------------------------------------------------
    'Expect the primary error and no cleanup failure
        TEST_Harness.AssertRaised "primary-error", ERR_EXAMPLE_OPERATION, RaisedError
        TEST_Harness.AssertEqualText "no-cleanup-failure", vbNullString, Details
    'Expect both settings restored
        TEST_Harness.AssertTrue "calculation-restored", _
            Application.Calculation = SavedCalculation, "Calculation differs after the operation"
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCleanupEachRestorationFailure()

'
'------------------------------------------------------------------------------
'                        CASE EACH RESTORATION FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Each restoration failure on its own is reported, names only its
'   setting and does not stop the other restoration. 4 assertions.
'
' WHY THIS EXISTS
'   Shows that every restoration is attempted and that a failure names
'   only its own setting
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation twice with no failure of the operation: once
'   with only the Calculation restoration failing and once with only the
'   ScreenUpdating restoration failing, and asserts for each the error
'   raised and that the details name only the failing setting
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertRaised, AssertTrue, RecordCaseError);
'   RunExampleOperation
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details     As String   'Cleanup failures reported by the operation
    Dim RaisedError As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "cleanup-each-restoration-failure"

'------------------------------------------------------------------------------
' FAIL THE CALCULATION RESTORATION
'------------------------------------------------------------------------------
    'Run with only the Calculation restoration failing
        On Error Resume Next
        RunExampleOperation vbNullString, True, False, Details
    'Capture the error number on the line after the call
        RaisedError = Err.Number
    'Restore the case handler
        On Error GoTo CaseFailed
    'Expect the cleanup error and only the Calculation failure in the details
        TEST_Harness.AssertRaised "calculation-cleanup-error", ERR_CLEANUP_FAILED, RaisedError
        TEST_Harness.AssertTrue "only-calculation-reported", _
            InStr(Details, "Calculation not restored") > 0 _
            And InStr(Details, "ScreenUpdating") = 0, "details='" & Details & "'"

'------------------------------------------------------------------------------
' FAIL THE SCREEN RESTORATION
'------------------------------------------------------------------------------
    'Run with only the ScreenUpdating restoration failing
        On Error Resume Next
        RunExampleOperation vbNullString, False, True, Details
    'Capture the error number on the line after the call
        RaisedError = Err.Number
    'Restore the case handler
        On Error GoTo CaseFailed
    'Expect the cleanup error and only the ScreenUpdating failure in the details
        TEST_Harness.AssertRaised "screen-cleanup-error", ERR_CLEANUP_FAILED, RaisedError
        TEST_Harness.AssertTrue "only-screen-reported", _
            InStr(Details, "ScreenUpdating not restored") > 0 _
            And InStr(Details, "Calculation") = 0, "details='" & Details & "'"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCleanupOnlyFailure()

'
'------------------------------------------------------------------------------
'                          CASE CLEANUP ONLY FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   When only the cleanup fails, the operation raises ERR_CLEANUP_FAILED,
'   never success. 2 assertions.
'
' WHY THIS EXISTS
'   Shows that a successful operation whose cleanup failed is not a
'   success
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation with no failure of the operation and the
'   Calculation restoration failing, and asserts that it raised
'   ERR_CLEANUP_FAILED and returned a cleanup detail
'
' ERROR POLICY
'   An error that escapes the case is recorded with RecordCaseError of
'   TEST_Harness, so that the remaining cases still run
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertRaised, AssertTrue, RecordCaseError);
'   RunExampleOperation
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details     As String   'Cleanup failures reported by the operation
    Dim RaisedError As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Send any error of the case to its handler
        On Error GoTo CaseFailed
    'Open the case with its identifier
        TEST_Harness.BeginCase "cleanup-only-failure"

'------------------------------------------------------------------------------
' RUN OPERATION
'------------------------------------------------------------------------------
    'Run with no failure of the operation and both restorations failing
        On Error Resume Next
        RunExampleOperation vbNullString, True, True, Details
    'Capture the error number on the line after the call
        RaisedError = Err.Number
    'Restore the case handler
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' ASSERT OUTCOMES
'------------------------------------------------------------------------------
    'Expect the cleanup error, never success
        TEST_Harness.AssertRaised "cleanup-error", ERR_CLEANUP_FAILED, RaisedError
    'Expect a cleanup detail
        TEST_Harness.AssertTrue "failure-reported", Len(Details) > 0, "no cleanup detail returned"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error as a failure of the case, so that the other cases still run
        TEST_Harness.RecordCaseError

End Sub

'
'------------------------------------------------------------------------------
'
'                        EXAMPLE PROCEDURES UNDER TEST
'
'------------------------------------------------------------------------------
'

Private Function ExampleRatio( _
    ByVal Numerator As Double, _
    ByVal Denominator As Double) _
    As Variant
'
'------------------------------------------------------------------------------
'                               CALCULATE RATIO
'------------------------------------------------------------------------------
' PURPOSE
'   Numerator / Denominator, or the worksheet error #DIV/0! for a zero
'   denominator, as a worksheet function returns errors (house style)
'
' WHY THIS EXISTS
'   Gives CaseErrorValue a procedure that returns an error value instead of
'   raising it
'
' INPUTS
'   Numerator
'     Dividend, as a Double
'
'   Denominator
'     Divisor, as a Double; zero gives the error value
'
' RETURNS
'   The quotient as a Double, or CVErr(xlErrDiv0) when Denominator is zero
'
' BEHAVIOR
'   Tests the denominator against zero, and returns the error value or the
'   quotient
'
' ERROR POLICY
'   Never raises: a zero denominator is returned as an error value
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CALCULATE RATIO
'------------------------------------------------------------------------------
    'Return the worksheet error for a zero denominator, and the quotient otherwise
        If Denominator = 0# Then
            ExampleRatio = CVErr(xlErrDiv0)
        Else
            ExampleRatio = Numerator / Denominator
        End If

End Function

Private Function ExampleSquareRoot( _
    ByVal Value As Double) _
    As Double
'
'------------------------------------------------------------------------------
'                            CALCULATE SQUARE ROOT
'------------------------------------------------------------------------------
' PURPOSE
'   Square root of Value >= 0; a negative Value raises ERR_EXAMPLE_INVALID,
'   as a core procedure raises (house style)
'
' WHY THIS EXISTS
'   Gives CaseRaisedError a procedure that raises an error for input outside
'   its domain
'
' INPUTS
'   Value
'     Number whose square root is taken; must not be negative
'
' RETURNS
'   The square root of Value, as a Double
'
' BEHAVIOR
'   Raises for a negative Value, and returns Sqr(Value) otherwise
'
' ERROR POLICY
'   Raises ERR_EXAMPLE_INVALID for a negative Value, with the value in the
'   message
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Reject a negative value
        If Value < 0# Then
            Err.Raise ERR_EXAMPLE_INVALID, "TEST_Examples.ExampleSquareRoot", _
                "Value must be >= 0; got " & CStr(Value) & "."
        End If

'------------------------------------------------------------------------------
' CALCULATE ROOT
'------------------------------------------------------------------------------
    'Return the square root
        ExampleSquareRoot = Sqr(Value)

End Function

Private Function FixturePath( _
    ByVal Relative As String) _
    As String
'
'------------------------------------------------------------------------------
'                             COMPOSE FIXTURE PATH
'------------------------------------------------------------------------------
' PURPOSE
'   The full path of a file of tests/fixtures/, given relative to it with
'   "/" separators. The separator is the one the fixtures folder uses: "/"
'   when it has only forward slashes, as in LibreOffice on Linux, otherwise
'   "\"
'
' WHY THIS EXISTS
'   The same relative path must work on a Windows host and on the Linux host
'   of LibreOffice
'
' INPUTS
'   Relative
'     Path of a file below tests/fixtures/, with "/" separators
'
' RETURNS
'   The full path, with the separator of the fixtures folder
'
' BEHAVIOR
'   Requires the fixtures folder, chooses the separator from the folder and
'   joins the folder and the relative path, with its separators replaced
'
' ERROR POLICY
'   Raises ERR_EXAMPLE_FIXTURE when no fixtures folder was passed to the
'   running suite
'
' DEPENDENCIES
'   mFixtures
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Separator   As String   'Path separator of the fixtures folder

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Require the fixtures folder of the running suite
        If Len(mFixtures) = 0 Then
            Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.FixturePath", _
                "No fixtures folder: pass the full path of tests/fixtures."
        End If

'------------------------------------------------------------------------------
' COMPOSE PATH
'------------------------------------------------------------------------------
    'Use a backslash unless the folder has only forward slashes
        Separator = "\"
        If InStr(mFixtures, "/") > 0 And InStr(mFixtures, "\") = 0 Then
            Separator = "/"
        End If
    'Join the folder and the relative path, with the separator of the folder
        FixturePath = mFixtures & Separator & Replace(Relative, "/", Separator)

End Function

Private Function ReadFixture( _
    ByVal Path As String) _
    As String
'
'------------------------------------------------------------------------------
'                                 READ FIXTURE
'------------------------------------------------------------------------------
' PURPOSE
'   The bytes of an ASCII test case file, as text with its line endings kept
'
' WHY THIS EXISTS
'   The fixtures are read as bytes so that CR LF and LF endings stay as they
'   are, and a byte above 127 is refused
'
' INPUTS
'   Path
'     Full path of the file
'
' RETURNS
'   The content of the file
'
' BEHAVIOR
'   Opens the file for binary read, reads it whole, closes it and requires
'   every character to be ASCII
'
' ERROR POLICY
'   A missing or unreadable file raises the file error; a byte above 127
'   raises ERR_EXAMPLE_FIXTURE. The file is closed first
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FileNumber          As Integer  'Handle of the open file
    Dim Text                As String   'Content of the file
    Dim Index               As Long     'Position in Text
    Dim ErrorNumber         As Long     'File error number
    Dim ErrorSource         As String   'File error source
    Dim ErrorDescription    As String   'File error description

'------------------------------------------------------------------------------
' READ FILE
'------------------------------------------------------------------------------
    'Open with Access Read, which raises on a missing file instead of creating it
        FileNumber = FreeFile
        Open Path For Binary Access Read As #FileNumber
    'Send an error of the read to the handler, which closes the file
        On Error GoTo ReadFailed
    'Read the whole file and close it
        Text = Space$(LOF(FileNumber))
        Get #FileNumber, , Text
        Close #FileNumber
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' REQUIRE ASCII
'------------------------------------------------------------------------------
    'Reject any character above 127
        For Index = 1 To Len(Text)
            If AscW(Mid$(Text, Index, 1)) > 127 Or AscW(Mid$(Text, Index, 1)) < 0 Then
                Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.ReadFixture", _
                    Path & " is not ASCII at byte " & CStr(Index) & "."
            End If
        Next Index

'------------------------------------------------------------------------------
' RETURN TEXT
'------------------------------------------------------------------------------
    'Return the text with its line endings kept
        ReadFixture = Text

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ReadFailed:
    'Capture the file error before any other call
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Close the file
        Close #FileNumber
    'Raise the file error again
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Function

Private Function SumJsonNumbers( _
    ByVal Path As String, _
    ByVal Key As String, _
    ByRef Records As Long) _
    As Double
'
'------------------------------------------------------------------------------
'                             ADD UP JSON NUMBERS
'------------------------------------------------------------------------------
' PURPOSE
'   The sum of every number that follows "Key": in a JSON fixture, and how
'   many there are. Enough for the illustrative fixture, whose amounts are
'   plain numbers; not a JSON parser
'
' WHY THIS EXISTS
'   CaseIllustrativeResults needs the total of the notional amounts of the
'   illustrative JSON test case
'
' INPUTS
'   Path
'     Full path of the JSON fixture
'
'   Key
'     Name of the field, without quotes
'
'   Records
'     Output: the number of values added up
'
' RETURNS
'   The sum of the numbers that follow the key
'
' BEHAVIOR
'   Reads the fixture, looks for the key and its colon, and for each match
'   adds Val of the text after it and counts it. Val reads the number with a
'   period as the decimal separator, whatever the regional settings, and
'   skips the blanks
'
' ERROR POLICY
'   Raises what ReadFixture raises
'
' DEPENDENCIES
'   ReadFixture
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Text        As String   'Content of the fixture
    Dim Marker      As String   'The key and its colon
    Dim Position    As Long     'Position of the next marker
    Dim Total       As Double   'Sum so far

'------------------------------------------------------------------------------
' READ FIXTURE
'------------------------------------------------------------------------------
    'Read the fixture and build the marker to look for
        Text = ReadFixture(Path)
        Marker = """" & Key & """:"
    'Start the count from zero
        Records = 0

'------------------------------------------------------------------------------
' ADD UP VALUES
'------------------------------------------------------------------------------
    'Add each number after a marker; Val reads a period as the decimal separator and skips blanks
        Position = InStr(1, Text, Marker, vbBinaryCompare)
        Do While Position > 0
            Total = Total + Val(Mid$(Text, Position + Len(Marker)))
            Records = Records + 1
            Position = InStr(Position + Len(Marker), Text, Marker, vbBinaryCompare)
        Loop
    'Return the sum
        SumJsonNumbers = Total

End Function

Private Function SumCsvColumn( _
    ByVal Path As String, _
    ByVal Column As String, _
    ByRef Records As Long) _
    As Double
'
'------------------------------------------------------------------------------
'                              ADD UP CSV COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   The sum of one column of a comma-separated file with a header row, and
'   how many rows there are. Lines end in CR LF or LF; empty lines are
'   ignored. Enough for the illustrative positions file, whose fields hold
'   no comma or quote; not a general CSV reader
'
' WHY THIS EXISTS
'   CaseIllustrativeResults needs the total of the notional amounts of the
'   illustrative file-import test case, which has CR LF and LF variants
'
' INPUTS
'   Path
'     Full path of the CSV fixture
'
'   Column
'     Name of the column in the header row
'
'   Records
'     Output: the number of rows added up
'
' RETURNS
'   The sum of the column over the rows
'
' BEHAVIOR
'   Reads the fixture, drops the CR characters, splits it into lines, finds
'   the column in the header and adds up Val of that field in each line that
'   is not empty
'
' ERROR POLICY
'   Raises ERR_EXAMPLE_FIXTURE when the header lacks Column or a row lacks
'   the field; raises what ReadFixture raises
'
' DEPENDENCIES
'   ReadFixture
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lines   As Variant  'The file's lines
    Dim Header  As Variant  'The header's fields
    Dim Fields  As Variant  'One row's fields
    Dim Target  As Long     'Index of Column, -1 until found
    Dim Index   As Long     'Line or field index
    Dim Total   As Double   'Sum so far

'------------------------------------------------------------------------------
' FIND COLUMN
'------------------------------------------------------------------------------
    'Split the file into lines whatever the line endings, and the first line into fields
        Lines = Split(Replace(ReadFixture(Path), vbCr, vbNullString), vbLf)
        Header = Split(Lines(0), ",")
    'Find the index of the column in the header
        Target = -1
        For Index = 0 To UBound(Header)
            If Header(Index) = Column Then
                Target = Index
            End If
        Next Index
    'Reject a header without the column
        If Target < 0 Then
            Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.SumCsvColumn", _
                Path & " has no column " & Column & "."
        End If

'------------------------------------------------------------------------------
' ADD UP VALUES
'------------------------------------------------------------------------------
    'Start the count from zero
        Records = 0
    'Add the field of each line that is not empty, and reject a line that lacks it
        For Index = 1 To UBound(Lines)
            If Len(Lines(Index)) > 0 Then
                Fields = Split(Lines(Index), ",")
                If UBound(Fields) < Target Then
                    Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.SumCsvColumn", _
                        Path & " line " & CStr(Index + 1) & " has no " & Column & "."
                End If
                Total = Total + Val(Fields(Target))
                Records = Records + 1
            End If
        Next Index
    'Return the sum
        SumCsvColumn = Total

End Function

Private Sub RunExampleOperation( _
    ByVal FailAt As String, _
    ByVal FailRestoreCalculation As Boolean, _
    ByVal FailRestoreScreen As Boolean, _
    ByRef CleanupDetails As String)
'
'------------------------------------------------------------------------------
'                            RUN EXAMPLE OPERATION
'------------------------------------------------------------------------------
' PURPOSE
'   Reference implementation of the cleanup pattern of
'   docs/VBA_HOUSE_STYLE.md (Operation and cleanup outcomes): capture each
'   setting, register its restoration before changing it, restore only what
'   was registered on both paths, preserve the primary error and report
'   every cleanup failure
'
' WHY THIS EXISTS
'   Gives the failure-path cases an operation that changes host state and
'   fails where they ask, so that each outcome of the pattern can be tested
'   without leaving Excel changed
'
' INPUTS
'   FailAt
'     Test seam. Empty for none, or where to raise ERR_EXAMPLE_HOST or
'     ERR_EXAMPLE_OPERATION: "capture-screen" (after Calculation was
'     captured), "change-calculation" (after its restoration was
'     registered) or "work"
'
'   FailRestoreCalculation
'     Test seam. The restore helper restores Calculation and then reports a
'     simulated failure
'
'   FailRestoreScreen
'     Test seam. The restore helper restores ScreenUpdating and then reports
'     a simulated failure
'
'   CleanupDetails
'     Output: the cleanup failures joined by "; ", empty on success. Set
'     also when an error is raised
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Captures Calculation and ScreenUpdating, registers each restoration just
'   before the change, sets ScreenUpdating to False and Calculation to
'   manual, fails where FailAt says, then restores each registered setting
'   to the value captured. A restoration failure is added to CleanupDetails,
'   and every restoration is attempted
'
' ERROR POLICY
'   Re-raises the primary error after cleanup; when only the cleanup failed,
'   raises ERR_CLEANUP_FAILED with the details
'
' DEPENDENCIES
'   RestoreCalculation, RestoreScreenUpdating and AppendDetail; the
'   Application object (Calculation and ScreenUpdating)
'
' NOTES
'   A restoration is registered only for a captured setting, just before its
'   change, so a change that fails part-way is still restored and an
'   unchanged setting is never touched
'
'   State: changes Application.ScreenUpdating to False and
'   Application.Calculation to manual, and restores each one it registered to
'   the value it captured
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedCalculation        As XlCalculation    'Calculation mode before the operation
    Dim SavedScreen             As Boolean          'ScreenUpdating before the operation
    Dim RestoreCalculationDue   As Boolean          'Calculation captured and about to change
    Dim RestoreScreenDue        As Boolean          'ScreenUpdating captured and about to change
    Dim ErrNumber               As Long             'Primary error number; 0 for none
    Dim ErrSource               As String           'Primary error source
    Dim ErrDescription          As String           'Primary error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Clear the details of an earlier call
        CleanupDetails = vbNullString
    'Save a failure of the operation as the primary error
        On Error GoTo Failed

'------------------------------------------------------------------------------
' CAPTURE AND CHANGE SETTINGS
'------------------------------------------------------------------------------
    'Capture the calculation mode
        SavedCalculation = Application.Calculation
    'Simulate a failure capturing ScreenUpdating
        If FailAt = "capture-screen" Then
            Err.Raise ERR_EXAMPLE_HOST, "TEST_Examples.RunExampleOperation", _
                "Simulated failure capturing ScreenUpdating."
        End If
    'Capture ScreenUpdating, register its restoration and then change it
        SavedScreen = Application.ScreenUpdating
        RestoreScreenDue = True
        Application.ScreenUpdating = False
    'Register the restoration of Calculation before changing it
        RestoreCalculationDue = True
    'Simulate a failure changing Calculation
        If FailAt = "change-calculation" Then
            Err.Raise ERR_EXAMPLE_HOST, "TEST_Examples.RunExampleOperation", _
                "Simulated failure changing Calculation."
        End If
    'Switch to manual calculation
        Application.Calculation = xlCalculationManual

'------------------------------------------------------------------------------
' DO THE WORK
'------------------------------------------------------------------------------
    'Simulate a failure of the operation itself
        If FailAt = "work" Then
            Err.Raise ERR_EXAMPLE_OPERATION, "TEST_Examples.RunExampleOperation", _
                "Simulated failure of the operation."
        End If

'------------------------------------------------------------------------------
' RESTORE SETTINGS
'------------------------------------------------------------------------------
CleanUp:
    'Restore normal error handling before the cleanup and the raise
        On Error GoTo 0
    'Restore Calculation when it was registered, and keep the failure
        If RestoreCalculationDue Then
            CleanupDetails = AppendDetail(CleanupDetails, _
                RestoreCalculation(SavedCalculation, FailRestoreCalculation))
        End If
    'Restore ScreenUpdating when it was registered, and keep the failure
        If RestoreScreenDue Then
            CleanupDetails = AppendDetail(CleanupDetails, _
                RestoreScreenUpdating(SavedScreen, FailRestoreScreen))
        End If

'------------------------------------------------------------------------------
' REPORT OUTCOME
'------------------------------------------------------------------------------
    'Raise the primary error again, after the cleanup
        If ErrNumber <> 0 Then
            Err.Raise ErrNumber, ErrSource, ErrDescription
        End If
    'Raise a cleanup failure when only the cleanup failed
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "TEST_Examples.RunExampleOperation", CleanupDetails
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Save the primary error before any other call
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Go through the cleanup, which raises the saved error again
        Resume CleanUp

End Sub

Private Function RestoreCalculation( _
    ByVal SavedCalculation As XlCalculation, _
    ByVal SimulateFailure As Boolean) _
    As String
'
'------------------------------------------------------------------------------
'                             RESTORE CALCULATION
'------------------------------------------------------------------------------
' PURPOSE
'   Restore Application.Calculation; contain any error
'
' WHY THIS EXISTS
'   Each setting is restored through its own helper and handler, so that a
'   failure of one restoration never stops the other
'
' INPUTS
'   SavedCalculation
'     Value to restore, captured before the change
'
'   SimulateFailure
'     Test seam. After the restoration, raises ERR_EXAMPLE_RESTORE
'     ("simulated") to exercise the failure path
'
' RETURNS
'   Empty on success, otherwise what failed
'
' BEHAVIOR
'   Sets Application.Calculation to the saved value, then raises the simulated
'   failure when asked. The error handler turns any error into the text
'   returned
'
' ERROR POLICY
'   Never raises: an error is returned as "Calculation not restored: " and the
'   description of the error
'
' DEPENDENCIES
'   The Application object (Calculation)
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn any error of the restoration into the returned text
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE SETTING
'------------------------------------------------------------------------------
    'Restore the saved value
        Application.Calculation = SavedCalculation
    'Report a simulated failure after the restoration
        If SimulateFailure Then
            Err.Raise ERR_EXAMPLE_RESTORE, "TEST_Examples.RestoreCalculation", "simulated"
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return what failed
        RestoreCalculation = "Calculation not restored: " & Err.Description

End Function

Private Function RestoreScreenUpdating( _
    ByVal SavedScreen As Boolean, _
    ByVal SimulateFailure As Boolean) _
    As String
'
'------------------------------------------------------------------------------
'                           RESTORE SCREEN UPDATING
'------------------------------------------------------------------------------
' PURPOSE
'   Restore Application.ScreenUpdating; contain any error
'
' WHY THIS EXISTS
'   Each setting is restored through its own helper and handler, so that a
'   failure of one restoration never stops the other
'
' INPUTS
'   SavedScreen
'     Value to restore, captured before the change
'
'   SimulateFailure
'     Test seam. After the restoration, raises ERR_EXAMPLE_RESTORE
'     ("simulated") to exercise the failure path
'
' RETURNS
'   Empty on success, otherwise what failed
'
' BEHAVIOR
'   Sets Application.ScreenUpdating to the saved value, then raises the simulated
'   failure when asked. The error handler turns any error into the text
'   returned
'
' ERROR POLICY
'   Never raises: an error is returned as "ScreenUpdating not restored: " and the
'   description of the error
'
' DEPENDENCIES
'   The Application object (ScreenUpdating)
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn any error of the restoration into the returned text
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE SETTING
'------------------------------------------------------------------------------
    'Restore the saved value
        Application.ScreenUpdating = SavedScreen
    'Report a simulated failure after the restoration
        If SimulateFailure Then
            Err.Raise ERR_EXAMPLE_RESTORE, "TEST_Examples.RestoreScreenUpdating", "simulated"
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return what failed
        RestoreScreenUpdating = "ScreenUpdating not restored: " & Err.Description

End Function

Private Function AppendDetail( _
    ByVal First As String, _
    ByVal Second As String) _
    As String
'
'------------------------------------------------------------------------------
'                                APPEND DETAIL
'------------------------------------------------------------------------------
' PURPOSE
'   Join two failure texts with "; ", skipping empty ones
'
' WHY THIS EXISTS
'   RunExampleOperation collects the failures of its restorations in one
'   text
'
' INPUTS
'   First
'     Failure text collected so far; may be empty
'
'   Second
'     Failure text to add; may be empty
'
' RETURNS
'   The two texts joined with "; ", or the one that is not empty, or an
'   empty text when both are
'
' BEHAVIOR
'   Returns Second when First is empty, First when Second is empty, and
'   First, "; " and Second otherwise
'
' ERROR POLICY
'   Never raises
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' JOIN TEXTS
'------------------------------------------------------------------------------
    'Skip an empty text, and join two with a separator
        If Len(First) = 0 Then
            AppendDetail = Second
        ElseIf Len(Second) = 0 Then
            AppendDetail = First
        Else
            AppendDetail = First & "; " & Second
        End If

End Function
