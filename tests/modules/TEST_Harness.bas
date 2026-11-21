Attribute VB_Name = "TEST_Harness"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Harness
'------------------------------------------------------------------------------
' PURPOSE
'   Deterministic regression harness for the workbook's VBA test suites: one
'   run per suite, stable case identifiers, exact and tolerance assertions,
'   expected error values and raised errors, expected case and assertion
'   counts, a complete key=value log, and the certification of a suite's runs
'   as the evidence of an Excel run
'
' WHY THIS EXISTS
'   Gives every suite the same lifecycle, assertions, log and certification, so
'   that a run in Excel leaves records that tools/check_excel_run.py can check
'   against one commit without Excel. It holds no calculation of the workbook
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Groups the procedures in five families:
'     - suite lifecycle: BeginRun, BeginCase, RecordCaseError, EndRun and
'       ResetTests (recovery)
'     - assertions: AssertEqualDouble, AssertNear, AssertEqualText, AssertTrue,
'       AssertErrorValue and AssertRaised; WithinTolerance is the comparison
'       rule
'     - certification: BeginCertification, CertifyRun and EndCertification
'     - log: Say, Echo, SaveLog and LastResult
'     - results: RecordResult and SaveResults
'
'   The private helpers are LogAssertion, RecordFailure, RequireActiveRun,
'   RequireCertification, CertSay, WriteRecords, WriteText, StampedCommit,
'   JsonText, JsonTextOrNull, ScopeJson, ActiveWindowCaption, VerifyHostState,
'   IsCaseId, IsKnownCase, NumberText, HostProperty, EnvironmentSummary and
'   ResetRunState
'
' ERROR POLICY
'   Assertion failures are recorded and the run continues. EndRun re-raises an
'   unexpected runner error unchanged after cleanup and reporting, or raises
'   ERR_TEST_FAILED for any other failed outcome. A run that starts while
'   another is active, or while recovery is required, is refused with
'   ERR_TEST_DIRTY_START. Misuse of the lifecycle raises ERR_TEST_USAGE.
'   EndCertification writes the certification log first, then raises
'   ERR_TEST_FAILED when the certification did not pass
'
' DEPENDENCIES
'   The VBA and Excel object models, and the build_commit cell of ThisWorkbook,
'   read by SaveResults; no reference, workbook fixture or external framework
'
' NOTES
'   Public surface: the suite lifecycle, the assertions, the log, the results,
'   the recovery and the certification procedures listed under BEHAVIOR. Option
'   Private Module keeps all of them out of the workbook's external surface,
'   and Option Explicit requires explicit declarations
'
'   State ownership: the module owns the run counters, the case identifiers,
'   the failure records, the retained log, the recorded results, the
'   certification records and the active-run, recovery-required and
'   active-certification flags. It reads four Application properties, the
'   number of open workbooks and the caption of the active window to prove a
'   run left them unchanged, and never changes them. It reads the build stamp's
'   commit from ThisWorkbook for SaveResults and never writes a workbook or
'   worksheet
'
'   Log format: one record per line, KEY=value, with the fields separated by a
'   semicolon and a space. The keys and their order are documented in
'   docs/methodology/TEST_HARNESS.md (Log format, Certification log).
'   SaveResults writes the JSON results file of docs/methodology/TEST_CASES.md
'   (Accuracy contracts)
'
'   Compatibility: Excel VBA on Windows. The log records the host, its version,
'   the Office bitness and the VBA generation
'
'   Usage: called by the suites in tests/modules/, for example
'   TEST_Examples.RunExampleTests and TEST_Examples.CertifyExamples; SaveLog
'   and ResetTests are also typed in the Immediate window. It has no suite of
'   its own
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-04
'==============================================================================

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'Test-only error numbers, kept apart from the production range in use
    'Written as literals equal to vbObjectError + n: LibreOffice, the optional
    'smoke host, cannot evaluate vbObjectError in a constant expression
    'ERR_TEST_DIRTY_START is raised while a run is active or recovery is required
    Public Const ERR_TEST_DIRTY_START   As Long = &H80040FE0    'vbObjectError + 4064: Run refused: unsafe start
    Public Const ERR_TEST_FAILED        As Long = &H80040FE1    'vbObjectError + 4065: Run finished without PASS
    Public Const ERR_TEST_USAGE         As Long = &H80040FE2    'vbObjectError + 4066: Lifecycle used out of order

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'The defined name that carries the build stamp's commit
    '(tools/devbuild/workbook.json, stamp.commit)
    Private Const STAMP_COMMIT_NAME As String = "build_commit"  'Defined name of the build stamp's commit

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'One run at a time. Counters and the log stay readable after EndRun,
    'until the next BeginRun or ResetTests
    Private mRunActive          As Boolean          'True from BeginRun to EndRun
    Private mRecoveryRequired   As Boolean          'True until ResetTests after an unsafe end
    Private mRecoveryReason     As String           'Why recovery is required
    Private mSuite              As String           'Name of the running or last suite
    Private mExpectedCases      As Long             'Cases a complete run starts
    Private mExpectedAssertions As Long             'Assertions a complete run evaluates
    Private mCaseCount          As Long             'Cases started in this run
    Private mAssertionCount     As Long             'Assertions evaluated in this run
    Private mFailureCount       As Long             'Failures of any kind in this run
    Private mErrorCount         As Long             'Unexpected errors in this run
    Private mCurrentCase        As String           'Identifier of the running case
    Private mCaseIds            As Collection       'Case identifiers seen in this run
    Private mPrimaryFailure     As String           'First failure recorded in this run
    Private mLastResult         As String           'RESULT record of the last finished run
    Private mLog                As Collection       'Log since the last BeginRun or ResetTests
    Private mResults            As Collection       'Results recorded since the last BeginRun or ResetTests
    Private mSnapshotValid      As Boolean          'True once all four properties were read
    Private mInitialCalculation As XlCalculation    'Calculation at BeginRun
    Private mInitialAlerts      As Boolean          'DisplayAlerts at BeginRun
    Private mInitialEvents      As Boolean          'EnableEvents at BeginRun
    Private mInitialScreen      As Boolean          'ScreenUpdating at BeginRun
    Private mInitialWorkbooks   As Long             'Open workbooks at BeginRun
    Private mInitialWindow      As String           'Active window caption at BeginRun
    Private mCertActive         As Boolean          'True from BeginCertification to EndCertification
    Private mCertSuite          As String           'Suite being certified
    Private mCertLog            As Collection       'Certification records
    Private mCertModes          As Long             'Runs certified
    Private mCertFailures       As Long             'Runs that did not end as required
    Private mCertMandatory      As Long             'Runs required to pass
    Private mCertComplete       As Long             'Required runs that were complete


'
'------------------------------------------------------------------------------
'
'                               SUITE LIFECYCLE
'
'------------------------------------------------------------------------------
'

Public Sub BeginRun( _
    ByVal SuiteName As String, _
    ByVal Mode As String, _
    ByVal ExpectedCases As Long, _
    ByVal ExpectedAssertions As Long)
'
'------------------------------------------------------------------------------
'                                 START A RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Starts one suite run: refuses an unsafe start, resets the run state,
'   records the host state to compare at the end and writes the log header
'
' WHY THIS EXISTS
'   Every suite run needs the same preconditions and the same baseline. A start
'   is refused while another run is active or recovery is required, so that the
'   state of an interrupted run is not overwritten before it has been
'   acknowledged with ResetTests. The host state read here is the one that
'   VerifyHostState compares at EndRun
'
' INPUTS
'   SuiteName
'     Stable suite identifier, String, written in the RUN and RESULT records
'
'   Mode
'     NORMAL, or the name of a deliberate failure mode of the suite; String
'
'   ExpectedCases
'     Cases of a complete run; Long, 1 or more
'
'   ExpectedAssertions
'     Assertions of a complete run; Long, 1 or more
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs these steps in order:
'     - marks recovery as required, with the reason, when a run is still active
'     - when recovery is required, logs the RECOVERY and RESULT records
'       (RESULT=FAIL, operation=NOT_RUN) and raises ERR_TEST_DIRTY_START
'     - raises ERR_TEST_USAGE when an expected count is below 1
'     - resets the run state, the log and the results and claims the run for
'       the suite
'     - reads Calculation, DisplayAlerts, EnableEvents, ScreenUpdating, the
'       number of open workbooks and the active window caption
'     - logs the RUN and ENVIRONMENT records
'
' ERROR POLICY
'   Raises ERR_TEST_DIRTY_START, after logging why, when a run is active or
'   recovery is required; raises ERR_TEST_USAGE for counts below 1. A host
'   property that cannot be read leaves the snapshot invalid, which fails
'   cleanup at EndRun; it does not stop the run
'
' DEPENDENCIES
'   ResetRunState, Say, ActiveWindowCaption and EnvironmentSummary;
'   Application.Calculation, DisplayAlerts, EnableEvents, ScreenUpdating and
'   Workbooks
'
' NOTES
'   An interrupted run or an unverified cleanup must be acknowledged with
'   ResetTests before a new run may overwrite its state. A project reset (End
'   or Reset in the VBE) clears this state, so such a stop is not seen
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REFUSE AN UNSAFE START
'------------------------------------------------------------------------------
    'Treat a run still active as interrupted; a project reset clears its state, so that stop is not seen
        If mRunActive Then
            mRecoveryRequired = True
            mRecoveryReason = "run of " & mSuite & " still active (interrupted?)"
        End If
    'Refuse the start while recovery is required, after logging why
        If mRecoveryRequired Then
            mLastResult = "RESULT=FAIL; suite=" & SuiteName & "; operation=NOT_RUN" & _
                "; completeness=INCOMPLETE; cases=0; assertions=0; failures=0" & _
                "; cleanup=NOT_RUN; recovery=REQUIRED"
            'RESULT is always the last record of a run.
            Say "RECOVERY=REQUIRED; reason=" & mRecoveryReason & _
                "; action=check Excel state, then run TEST_Harness.ResetTests"
            Say mLastResult
            Err.Raise ERR_TEST_DIRTY_START, "TEST_Harness.BeginRun", _
                "Recovery required: " & mRecoveryReason & ". Run TEST_Harness.ResetTests."
        End If
    'Reject expected counts below 1
        If ExpectedCases < 1 Or ExpectedAssertions < 1 Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.BeginRun", _
                "Expected counts must be at least 1; got cases=" & CStr(ExpectedCases) & _
                ", assertions=" & CStr(ExpectedAssertions) & "."
        End If

'------------------------------------------------------------------------------
' RESET AND CLAIM THE RUN
'------------------------------------------------------------------------------
    'Clear the counters, the log and the results of the previous run
        ResetRunState
        Set mLog = New Collection
        Set mResults = New Collection
    'Claim the run for this suite with its expected counts
        mSuite = SuiteName
        mExpectedCases = ExpectedCases
        mExpectedAssertions = ExpectedAssertions
        mRunActive = True

'------------------------------------------------------------------------------
' SNAPSHOT HOST STATE
'------------------------------------------------------------------------------
    'Divert a failed read to the handler below
        On Error GoTo SnapshotFailed
    'Read the four Application properties, the open workbooks and the active window
        mInitialCalculation = Application.Calculation
        mInitialAlerts = Application.DisplayAlerts
        mInitialEvents = Application.EnableEvents
        mInitialScreen = Application.ScreenUpdating
        mInitialWorkbooks = Application.Workbooks.Count
        mInitialWindow = ActiveWindowCaption()
    'Mark the snapshot valid only after every read, so that cleanup never compares a value that was not read
        mSnapshotValid = True

'------------------------------------------------------------------------------
' WRITE LOG HEADER
'------------------------------------------------------------------------------
WriteHeader:
    'Restore normal error handling
        On Error GoTo 0
    'Write the RUN record with the expected counts and the start time
        Say "RUN=" & SuiteName & "; mode=" & Mode & _
            "; expected_cases=" & CStr(ExpectedCases) & _
            "; expected_assertions=" & CStr(ExpectedAssertions) & _
            "; started=" & Format$(Now, "yyyy-mm-dd hh:mm:ss")
    'Write the ENVIRONMENT record
        Say "ENVIRONMENT=" & EnvironmentSummary()

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the handler for a failed snapshot
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
SnapshotFailed:
    'Mark the snapshot invalid, so that cleanup fails at EndRun
        mSnapshotValid = False
    'Resume at the log header: the run still starts
        Resume WriteHeader

End Sub

Public Sub BeginCase( _
    ByVal CaseId As String)
'
'------------------------------------------------------------------------------
'                                 START A CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Counts and logs the start of one case
'
' WHY THIS EXISTS
'   A case identifier is the stable key by which logs are compared case by
'   case, so the harness checks its form and its uniqueness within the run when
'   the case starts
'
' INPUTS
'   CaseId
'     Lower-case kebab-case identifier, String, unique within the run and
'     stable across runs, so that logs can be compared case by case
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Requires an active run, counts the case, makes it the current case and logs
'   the CASE record. A malformed or repeated identifier is recorded as a
'   failure; any other identifier is remembered for the run
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A malformed or repeated identifier is
'   recorded as a failure; the case still counts and runs
'
' DEPENDENCIES
'   RequireActiveRun, Say, IsCaseId, IsKnownCase and RecordFailure
'
' NOTES
'   The identifier is added to the collection of case identifiers with itself
'   as the key
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COUNT THE CASE
'------------------------------------------------------------------------------
    'Require an active run, then count the case and make it the current one
        RequireActiveRun "BeginCase"
        mCaseCount = mCaseCount + 1
        mCurrentCase = CaseId
    'Write the CASE record
        Say "CASE=" & CaseId

'------------------------------------------------------------------------------
' CHECK THE IDENTIFIER
'------------------------------------------------------------------------------
    'Record a failure for a malformed or repeated identifier, otherwise remember it
        If Not IsCaseId(CaseId) Then
            RecordFailure "harness.case-id", "malformed case identifier '" & CaseId & "'"
        ElseIf IsKnownCase(CaseId) Then
            RecordFailure "harness.case-id", "repeated case identifier '" & CaseId & "'"
        Else
            mCaseIds.Add CaseId, CaseId
        End If


End Sub

Public Sub RecordCaseError()

'
'------------------------------------------------------------------------------
'                             RECORD A CASE ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Records the error that ended the running case, so that the remaining cases
'   still run
'
' WHY THIS EXISTS
'   A case ends in its own error handler, which calls this procedure first so
'   that an unexpected error becomes a failure of the run instead of stopping
'   it
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Saves Err.Number, Err.Source and Err.Description, requires an active run,
'   counts an unexpected error, logs the ERROR record and records a failure
'   named after the current case
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. It counts an unexpected error and a
'   failure and evaluates no assertion
'
' DEPENDENCIES
'   RequireActiveRun, Say and RecordFailure
'
' NOTES
'   Call it first thing in the case's error handler: it reads the Err object,
'   which anything else can reset
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrorNumber         As Long     'Number of the case error
    Dim ErrorSource         As String   'Source of the case error
    Dim ErrorDescription    As String   'Description of the case error

'------------------------------------------------------------------------------
' CAPTURE THE ERROR
'------------------------------------------------------------------------------
    'Save the error before any other call can reset it
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description

'------------------------------------------------------------------------------
' RECORD THE FAILURE
'------------------------------------------------------------------------------
    'Require an active run, then count the unexpected error
        RequireActiveRun "RecordCaseError"
        mErrorCount = mErrorCount + 1
    'Write the ERROR record
        Say "ERROR=" & mCurrentCase & "; number=" & CStr(ErrorNumber) & _
            "; source=" & ErrorSource & "; description=" & ErrorDescription
    'Record the failure under the current case
        RecordFailure mCurrentCase & ".unexpected-error", _
            "number=" & CStr(ErrorNumber) & "; source=" & ErrorSource

End Sub

Public Sub EndRun( _
    ByVal RunErrorNumber As Long, _
    ByVal RunErrorSource As String, _
    ByVal RunErrorDescription As String)
'
'------------------------------------------------------------------------------
'                                 FINISH A RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Finishes the run: checks completeness, releases the run, verifies the host
'   state, writes the summary and raises the outcome
'
' WHY THIS EXISTS
'   The outcome of a run has four parts that are judged and reported on their
'   own (operation, completeness, cleanup and recovery), and a run that does
'   not pass must never end silently
'
' INPUTS
'   RunErrorNumber
'     Number of the error that escaped the suite's runner, Long, or 0 when none
'     did
'
'   RunErrorSource
'     Source of that error, String, or an empty string when none did
'
'   RunErrorDescription
'     Description of that error, String, or an empty string when none did
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs these steps in order:
'     - requires an active run and records an error that escaped the runner
'     - sets operation to PASS, FAIL (an assertion or a case identifier failed)
'       or ERROR (an unexpected error occurred)
'     - sets completeness to COMPLETE only when both expected counts were met
'       exactly
'     - releases the run and verifies the host state; cleanup is PASS when the
'       host state is unchanged
'     - sets recovery to REQUIRED when the host state changed or could not be
'       checked, or when a start was refused while this run was active
'     - writes the CASES, ASSERTIONS, FAILURES, PRIMARY_FAILURE, CLEANUP,
'       RECOVERY and RESULT records
'     - raises the outcome
'
'   RESULT=PASS needs operation PASS, COMPLETE, cleanup PASS and no recovery
'
' ERROR POLICY
'   Re-raises the runner error unchanged, after reporting, so that cleanup
'   never replaces the primary failure; otherwise raises ERR_TEST_FAILED unless
'   RESULT=PASS. Raises ERR_TEST_USAGE outside a run
'
' DEPENDENCIES
'   RequireActiveRun, Say, RecordFailure and VerifyHostState
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
    Dim CleanupDetail   As String   'Reason for the cleanup outcome
    Dim CleanupPassed   As Boolean  'True when the host state is unchanged
    Dim Completed       As Boolean  'True when both expected counts were met
    Dim Operation       As String   'PASS, FAIL or ERROR
    Dim Verdict         As String   'PASS or FAIL

'------------------------------------------------------------------------------
' COLLECT THE OUTCOMES
'------------------------------------------------------------------------------
    'Require an active run, then end the current case
        RequireActiveRun "EndRun"
        mCurrentCase = "run"
    'Count and record an error that escaped the runner
        If RunErrorNumber <> 0 Then
            mErrorCount = mErrorCount + 1
            Say "ERROR=run; number=" & CStr(RunErrorNumber) & "; source=" & RunErrorSource & _
                "; description=" & RunErrorDescription
            RecordFailure "runner.unexpected-error", _
                "number=" & CStr(RunErrorNumber) & "; source=" & RunErrorSource
        End If
    'Set the operation outcome, which covers what ran; completeness and cleanup are judged separately below
        If mErrorCount > 0 Then
            Operation = "ERROR"
        ElseIf mFailureCount > 0 Then
            Operation = "FAIL"
        Else
            Operation = "PASS"
        End If

'------------------------------------------------------------------------------
' CHECK COMPLETENESS
'------------------------------------------------------------------------------
    'Require exact counts: a skipped or extra case or assertion means the suite did not run as written
        Completed = (mCaseCount = mExpectedCases) And _
                    (mAssertionCount = mExpectedAssertions)
        If Not Completed Then
            RecordFailure "harness.completeness", _
                "expected cases=" & CStr(mExpectedCases) & _
                ", assertions=" & CStr(mExpectedAssertions) & _
                "; observed cases=" & CStr(mCaseCount) & _
                ", assertions=" & CStr(mAssertionCount)
        End If

'------------------------------------------------------------------------------
' VERIFY THE HOST STATE
'------------------------------------------------------------------------------
    'Compare the host state with the one taken at BeginRun, then release the run
        CleanupPassed = VerifyHostState(CleanupDetail)
        mRunActive = False
    'Require recovery when the host state changed or could not be checked
        If Not CleanupPassed Then
            mRecoveryRequired = True
            mRecoveryReason = CleanupDetail
        End If
    'Pass only when the operation passed, the run was complete, cleanup passed and no recovery is required
        If Operation = "PASS" And Completed And CleanupPassed And Not mRecoveryRequired Then
            Verdict = "PASS"
        Else
            Verdict = "FAIL"
        End If

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Write the counts
        Say "CASES=" & CStr(mCaseCount) & "; expected=" & CStr(mExpectedCases)
        Say "ASSERTIONS=" & CStr(mAssertionCount) & "; expected=" & CStr(mExpectedAssertions)
        Say "FAILURES=" & CStr(mFailureCount) & "; errors=" & CStr(mErrorCount)
    'Write the primary failure, when there is one
        If Len(mPrimaryFailure) > 0 Then
            Say "PRIMARY_FAILURE=" & mPrimaryFailure
        End If
    'Write the cleanup outcome
        Say "CLEANUP=" & IIf(CleanupPassed, "PASS", "FAIL") & "; detail=" & CleanupDetail
    'Write the recovery state
        If mRecoveryRequired Then
            Say "RECOVERY=REQUIRED; reason=" & mRecoveryReason & _
                "; action=check Excel state, then run TEST_Harness.ResetTests"
        Else
            Say "RECOVERY=NONE"
        End If
    'Build the RESULT record and keep it for LastResult
        mLastResult = "RESULT=" & Verdict & "; suite=" & mSuite & _
            "; operation=" & Operation & _
            "; completeness=" & IIf(Completed, "COMPLETE", "INCOMPLETE") & _
            "; cases=" & CStr(mCaseCount) & _
            "; assertions=" & CStr(mAssertionCount) & _
            "; failures=" & CStr(mFailureCount) & _
            "; cleanup=" & IIf(CleanupPassed, "PASS", "FAIL") & _
            "; recovery=" & IIf(mRecoveryRequired, "REQUIRED", "NONE")
    'Write the RESULT record
        Say mLastResult

'------------------------------------------------------------------------------
' RAISE THE OUTCOME
'------------------------------------------------------------------------------
    'Re-raise the runner error unchanged, so that cleanup never replaces the primary failure
        If RunErrorNumber <> 0 Then
            Err.Raise RunErrorNumber, RunErrorSource, RunErrorDescription
        End If
    'Raise ERR_TEST_FAILED for any other run that did not pass
        If Verdict <> "PASS" Then
            Err.Raise ERR_TEST_FAILED, "TEST_Harness.EndRun", _
                "Suite " & mSuite & " did not pass; see the log."
        End If

End Sub

Public Sub ResetTests()

'
'------------------------------------------------------------------------------
'                              RESET THE HARNESS
'------------------------------------------------------------------------------
' PURPOSE
'   Acknowledges an interrupted run or an unverified cleanup, so that the next
'   run may start
'
' WHY THIS EXISTS
'   A start is refused while a run is active or recovery is required. This is
'   the one way to clear that state after the person running the tests has
'   checked Excel by hand
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resets the run state, clears both recovery fields and an unfinished
'   certification, starts an empty log and an empty result list, and logs the
'   RESET record
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   ResetRunState and Say
'
' NOTES
'   Run it only after the interrupted run has stopped and the Excel state has
'   been checked by hand; it restores nothing itself
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR THE STATE
'------------------------------------------------------------------------------
    'Reset the run state and clear the recovery flag and its reason
        ResetRunState
        mRecoveryRequired = False
        mRecoveryReason = vbNullString
    'Discard an unfinished certification
        mCertActive = False
        Set mCertLog = Nothing
    'Start an empty log and an empty result list
        Set mLog = New Collection
        Set mResults = New Collection

'------------------------------------------------------------------------------
' LOG THE RESET
'------------------------------------------------------------------------------
    'Write the RESET record
        Say "RESET=DONE; run_active=no; recovery=NONE"

End Sub

'
'------------------------------------------------------------------------------
'
'                                  ASSERTIONS
'
'------------------------------------------------------------------------------
'

Public Sub AssertEqualDouble( _
    ByVal AssertionName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Double)
'
'------------------------------------------------------------------------------
'                             ASSERT EQUAL DOUBLES
'------------------------------------------------------------------------------
' PURPOSE
'   Requires exact Double equality, for results that are exact in binary
'
' WHY THIS EXISTS
'   A result that is exact in binary can be compared without a tolerance, so
'   that any difference is a failure. AssertNear is the assertion for a result
'   that is not exact
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case; String
'
'   Expected
'     The value the suite requires; Double
'
'   Actual
'     The value the suite computed; Double
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Logs the assertion as passed when Actual equals Expected, with both values
'   as text from NumberText
'
' ERROR POLICY
'   Fails the assertion instead of raising; ERR_TEST_USAGE outside a run comes
'   from LogAssertion
'
' DEPENDENCIES
'   LogAssertion and NumberText
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Log the assertion as passed only when the two values are exactly equal
        LogAssertion AssertionName, (Actual = Expected), _
            "expected=" & NumberText(Expected) & "; actual=" & NumberText(Actual)

End Sub

Public Sub AssertNear( _
    ByVal AssertionName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal AbsoluteTolerance As Double, _
    ByVal RelativeTolerance As Double)
'
'------------------------------------------------------------------------------
'                           ASSERT WITHIN TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   Requires |Actual - Expected| <= max(AbsoluteTolerance, RelativeTolerance *
'   |Expected|), the rule of docs/methodology/TEST_CASES.md (Tolerance)
'
' WHY THIS EXISTS
'   A computed value is compared with a tolerance that the test case states, an
'   absolute and a relative bound, so that the rule is the same in every suite
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case; String
'
'   Expected
'     The value the suite requires; Double
'
'   Actual
'     The value the suite computed; Double
'
'   AbsoluteTolerance
'     Absolute bound, Double, 0 or more; always stated. Near zero the absolute
'     bound governs
'
'   RelativeTolerance
'     Relative bound, Double, 0 or more; always stated. Both 0 demands exact
'     equality
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Fails the assertion when a tolerance is negative; otherwise applies
'   WithinTolerance and logs the four values with the outcome
'
' ERROR POLICY
'   A negative tolerance fails the assertion; nothing is raised. ERR_TEST_USAGE
'   outside a run comes from LogAssertion
'
' DEPENDENCIES
'   LogAssertion, WithinTolerance and NumberText
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE WITH TOLERANCE
'------------------------------------------------------------------------------
    'Fail the assertion for a negative tolerance, otherwise apply the tolerance rule
        If AbsoluteTolerance < 0# Or RelativeTolerance < 0# Then
            LogAssertion AssertionName, False, "tolerances must not be negative"
        Else
            LogAssertion AssertionName, _
                WithinTolerance(Expected, Actual, AbsoluteTolerance, RelativeTolerance), _
                "expected=" & NumberText(Expected) & "; actual=" & NumberText(Actual) & _
                "; absolute=" & NumberText(AbsoluteTolerance) & _
                "; relative=" & NumberText(RelativeTolerance)
        End If

End Sub

Public Sub AssertEqualText( _
    ByVal AssertionName As String, _
    ByVal Expected As String, _
    ByVal Actual As String)
'
'------------------------------------------------------------------------------
'                              ASSERT EQUAL TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Requires the same text, compared binary: case and spaces count
'
' WHY THIS EXISTS
'   Text results are compared exactly, so that a difference of case or of
'   spacing is a failure
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case; String
'
'   Expected
'     The text the suite requires; String
'
'   Actual
'     The text the suite computed; String
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares the two texts with StrComp in binary mode and logs the assertion
'   with both texts in single quotes
'
' ERROR POLICY
'   Fails the assertion instead of raising; ERR_TEST_USAGE outside a run comes
'   from LogAssertion
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Log the assertion as passed only when the texts are equal in a binary comparison
        LogAssertion AssertionName, (StrComp(Actual, Expected, vbBinaryCompare) = 0), _
            "expected='" & Expected & "'; actual='" & Actual & "'"

End Sub

Public Sub AssertTrue( _
    ByVal AssertionName As String, _
    ByVal Condition As Boolean, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                              ASSERT A CONDITION
'------------------------------------------------------------------------------
' PURPOSE
'   Requires a condition that the suite evaluated; Detail is logged on failure
'
' WHY THIS EXISTS
'   Some checks have no expected value to compare, so the suite evaluates the
'   condition itself and the harness counts and logs it like any other
'   assertion
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case; String
'
'   Condition
'     The result of the suite's own check; Boolean, True when the assertion
'     holds
'
'   Detail
'     Text logged with the failure; String
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Logs the assertion as passed when Condition is True
'
' ERROR POLICY
'   Fails the assertion instead of raising; ERR_TEST_USAGE outside a run comes
'   from LogAssertion
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Log the condition the suite evaluated
        LogAssertion AssertionName, Condition, Detail

End Sub

Public Sub AssertErrorValue( _
    ByVal AssertionName As String, _
    ByVal ExpectedCode As Long, _
    ByVal Actual As Variant)
'
'------------------------------------------------------------------------------
'                            ASSERT AN ERROR VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Requires a specific worksheet error value, such as CVErr(xlErrValue)
'
' WHY THIS EXISTS
'   A worksheet function reports an error as an error value, and an error value
'   cannot be compared with an operator without raising an error, so the
'   comparison has its own assertion
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case; String
'
'   ExpectedCode
'     An xlErr constant; Long
'
'   Actual
'     The value returned; Variant
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Fails the assertion when Actual is not an error value, logging its type
'   name. Otherwise compares the text of Actual with the text of the expected
'   error value
'
' ERROR POLICY
'   Fails the assertion instead of raising; ERR_TEST_USAGE outside a run comes
'   from LogAssertion
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   Comparing a value with an error value by the equals operator raises error
'   13, Type mismatch, so the type is checked first and the error values are
'   compared as text, such as Error 2007
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Check the type, then compare as text (Error 2007): comparing an error value with = raises error 13
        If Not IsError(Actual) Then
            LogAssertion AssertionName, False, _
                "expected error " & CStr(ExpectedCode) & "; actual type=" & TypeName(Actual)
        Else
            LogAssertion AssertionName, (CStr(Actual) = CStr(CVErr(ExpectedCode))), _
                "expected error " & CStr(ExpectedCode) & "; actual=" & CStr(Actual)
        End If

End Sub

Public Sub AssertRaised( _
    ByVal AssertionName As String, _
    ByVal ExpectedNumber As Long, _
    ByVal ActualNumber As Long)
'
'------------------------------------------------------------------------------
'                            ASSERT A RAISED ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Requires that a call raised a specific error number
'
' WHY THIS EXISTS
'   A suite tests an error by calling the procedure and capturing Err.Number;
'   this assertion compares that number with the expected one
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case; String
'
'   ExpectedNumber
'     The error number the suite requires; Long
'
'   ActualNumber
'     The number the suite captured; Long, 0 when nothing was raised
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Logs the assertion as passed when the two numbers are equal
'
' ERROR POLICY
'   Fails the assertion instead of raising; ERR_TEST_USAGE outside a run comes
'   from LogAssertion
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   The suite captures Err.Number on the line after the call
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Log the assertion as passed only when the captured number is the expected one
        LogAssertion AssertionName, (ActualNumber = ExpectedNumber), _
            "expected error=" & CStr(ExpectedNumber) & "; actual error=" & CStr(ActualNumber)

End Sub

Public Function WithinTolerance( _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal AbsoluteTolerance As Double, _
    ByVal RelativeTolerance As Double) _
    As Boolean
'
'------------------------------------------------------------------------------
'                                TEST TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   The comparison rule of AssertNear, exposed so that suites can test it
'
' WHY THIS EXISTS
'   The rule of the tolerance assertion is a procedure of its own, so that a
'   suite can check the rule directly
'
' INPUTS
'   Expected
'     The value the suite requires; Double
'
'   Actual
'     The value the suite computed; Double
'
'   AbsoluteTolerance
'     Absolute bound; Double
'
'   RelativeTolerance
'     Relative bound; Double
'
' RETURNS
'   True when |Actual - Expected| <= max(AbsoluteTolerance, RelativeTolerance *
'   |Expected|); the bound is inclusive
'
' BEHAVIOR
'   Takes the larger of the absolute bound and the relative bound times
'   |Expected|, and compares the distance between the two values with it
'
' ERROR POLICY
'   Never raises for finite values
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
' DECLARE
'------------------------------------------------------------------------------
    Dim Bound   As Double   'Larger of the two bounds

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Take the larger of the relative bound and the absolute bound
        Bound = RelativeTolerance * Abs(Expected)
        If AbsoluteTolerance > Bound Then
            Bound = AbsoluteTolerance
        End If
    'Compare the distance with the bound, inclusive
        WithinTolerance = (Abs(Actual - Expected) <= Bound)

End Function

'
'------------------------------------------------------------------------------
'
'                                CERTIFICATION
'
'------------------------------------------------------------------------------
'

Public Sub BeginCertification( _
    ByVal SuiteName As String)
'
'------------------------------------------------------------------------------
'                            START A CERTIFICATION
'------------------------------------------------------------------------------
' PURPOSE
'   Starts the certification of one suite: refuses state left by an earlier
'   run, then keeps a certification log apart from the run logs
'
' WHY THIS EXISTS
'   A certification runs a suite in every mode and records how each run reached
'   its caller, so it needs a log of its own and a clean start
'
' INPUTS
'   SuiteName
'     The suite identifier, String, as in its RUN records
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs these steps in order:
'     - raises ERR_TEST_USAGE while another certification is active
'     - marks recovery as required, with the reason, when a run is still active
'     - raises ERR_TEST_DIRTY_START when recovery is required
'     - claims the certification, clears its counters and the RESULT record of
'       any earlier run, starts the certification log and writes the CERTIFY
'       record
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE, changing nothing, while another certification is
'   active. A run still active is taken as interrupted: it marks recovery
'   required, with the reason, so that ResetTests is needed before any run,
'   then raises ERR_TEST_DIRTY_START, as it does when recovery is already
'   required. It changes no other state before raising
'
' DEPENDENCIES
'   CertSay
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REFUSE LEFTOVER STATE
'------------------------------------------------------------------------------
    'Refuse a second certification while one is active
        If mCertActive Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.BeginCertification", _
                "The certification of " & mCertSuite & " is still active. " & _
                "Run TEST_Harness.ResetTests."
        End If
    'Treat a run that is still active as interrupted
        If mRunActive Then
            mRecoveryRequired = True
            mRecoveryReason = "run of " & mSuite & " still active (interrupted?)"
        End If
    'Refuse the start while recovery is required
        If mRecoveryRequired Then
            Err.Raise ERR_TEST_DIRTY_START, "TEST_Harness.BeginCertification", _
                "Recovery required: " & mRecoveryReason & ". Run TEST_Harness.ResetTests."
        End If

'------------------------------------------------------------------------------
' CLAIM THE CERTIFICATION
'------------------------------------------------------------------------------
    'Claim the certification for the suite and clear its counters
        mCertActive = True
        mCertSuite = SuiteName
        mCertModes = 0
        mCertFailures = 0
        mCertMandatory = 0
        mCertComplete = 0
    'Clear the RESULT record of any earlier run
        mLastResult = vbNullString
    'Start the certification log with the CERTIFY record
        Set mCertLog = New Collection
        CertSay "CERTIFY=" & SuiteName & "; started=" & Format$(Now, "yyyy-mm-dd hh:mm:ss")

End Sub

Public Sub CertifyRun( _
    ByVal Mode As String, _
    ByVal MustPass As Boolean, _
    ByVal RaisedNumber As Long, _
    ByVal RaisedDescription As String)
'
'------------------------------------------------------------------------------
'                                CERTIFY A RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Records how one run of the suite reached its caller, and whether that is
'   how the mode must end
'
' WHY THIS EXISTS
'   A deliberate failure mode is a self-test that a failure is never silent, so
'   a certification must see both the runs that pass and the runs that must
'   fail
'
' INPUTS
'   Mode
'     The run's mode; String
'
'   MustPass
'     True for a run that must pass, False for a deliberate failure mode;
'     Boolean
'
'   RaisedNumber
'     The number of the error that reached the caller of the suite's entry
'     point; Long, 0 when none did
'
'   RaisedDescription
'     The description of that error; String, empty when none did
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads the RESULT record of the run, then judges the run and logs the MODE
'   record:
'     - a run that must pass ends with no error and a RESULT=PASS that is
'       COMPLETE
'     - a deliberate failure ends with RESULT=FAIL and reaches its caller as
'       ERR_TEST_FAILED with a description
'     - a run with no RESULT of its own fails
'
'   Counts the mode, the failures, the runs that must pass and those of them
'   that were complete
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a certification. A run that is still active
'   did not end: it requires recovery and is recorded as failed, so that the
'   certification still writes its log
'
' DEPENDENCIES
'   RequireCertification and CertSay
'
' NOTES
'   The RESULT record is cleared once read, so that the next run starts without
'   one
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Completeness    As String   'COMPLETE, INCOMPLETE or NONE
    Dim Passed          As Boolean  'True when the run ended as required
    Dim Verdict         As String   'PASS, FAIL or NONE

'------------------------------------------------------------------------------
' READ THE RUN'S RESULT
'------------------------------------------------------------------------------
    'Require a certification
        RequireCertification "CertifyRun"
    'Treat a run that is still active as one that did not end, and discard its RESULT record
        If mRunActive Then
            mRunActive = False
            mRecoveryRequired = True
            mRecoveryReason = "run of " & mSuite & " did not end"
            mLastResult = vbNullString
        End If
    'Read the verdict and the completeness from the RESULT record, or NONE when the run left none
        If Len(mLastResult) = 0 Then
            Verdict = "NONE"
            Completeness = "NONE"
        Else
            Verdict = IIf(Left$(mLastResult, 12) = "RESULT=PASS;", "PASS", "FAIL")
            Completeness = IIf(InStr(mLastResult, "; completeness=COMPLETE;") > 0, _
                               "COMPLETE", "INCOMPLETE")
        End If
    'Clear the RESULT record once it is read
        mLastResult = vbNullString

'------------------------------------------------------------------------------
' JUDGE AND RECORD
'------------------------------------------------------------------------------
    'Judge the run against what its mode requires
        If MustPass Then
            Passed = (RaisedNumber = 0) And (Verdict = "PASS") And (Completeness = "COMPLETE")
            mCertMandatory = mCertMandatory + 1
            If Completeness = "COMPLETE" Then
                mCertComplete = mCertComplete + 1
            End If
        Else
            Passed = (RaisedNumber = ERR_TEST_FAILED) And (Len(RaisedDescription) > 0) _
                     And (Verdict = "FAIL")
        End If
    'Count the mode, and the failure when the run did not end as required
        mCertModes = mCertModes + 1
        If Not Passed Then
            mCertFailures = mCertFailures + 1
        End If
    'Write the MODE record
        CertSay "MODE=" & Mode & "; expected=" & IIf(MustPass, "PASS", "FAIL") & _
            "; error=" & CStr(RaisedNumber) & "; result=" & Verdict & _
            "; completeness=" & Completeness & "; status=" & IIf(Passed, "PASS", "FAIL")

End Sub

Public Sub EndCertification( _
    ByVal Path As String)
'
'------------------------------------------------------------------------------
'                            FINISH A CERTIFICATION
'------------------------------------------------------------------------------
' PURPOSE
'   Finishes the certification: writes its verdict, saves the certification log
'   and only then raises the outcome
'
' WHY THIS EXISTS
'   The evidence of an Excel run must exist whatever the outcome, so the log is
'   saved before the failure is raised
'
' INPUTS
'   Path
'     Full path of the certification log; String, created or overwritten
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs these steps in order:
'     - requires a certification
'     - sets completeness to COMPLETE when every run that must pass was
'       complete
'     - sets the verdict to PASS when there was at least one such run,
'       completeness is COMPLETE and every run ended as its mode requires
'     - logs the CERTIFICATION record and closes the certification
'     - saves the log to Path
'     - raises ERR_TEST_FAILED unless the certification passed
'
' ERROR POLICY
'   A file error is re-raised after the file is closed. Otherwise raises
'   ERR_TEST_FAILED, after the log is written, unless the certification passed.
'   Raises ERR_TEST_USAGE outside a certification
'
' DEPENDENCIES
'   RequireCertification, CertSay and the file writer for record collections
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
    Dim Completed   As Boolean  'True when every required run was complete
    Dim Verdict     As String   'PASS or FAIL

'------------------------------------------------------------------------------
' JUDGE
'------------------------------------------------------------------------------
    'Require a certification
        RequireCertification "EndCertification"
    'Judge completeness and the verdict from the counters
        Completed = (mCertMandatory > 0) And (mCertComplete = mCertMandatory)
        If Completed And mCertFailures = 0 Then
            Verdict = "PASS"
        Else
            Verdict = "FAIL"
        End If
    'Write the CERTIFICATION record
        CertSay "CERTIFICATION=" & Verdict & "; suite=" & mCertSuite & _
            "; completeness=" & IIf(Completed, "COMPLETE", "INCOMPLETE") & _
            "; modes=" & CStr(mCertModes) & "; failures=" & CStr(mCertFailures)
    'Mark the certification as finished
        mCertActive = False

'------------------------------------------------------------------------------
' WRITE THE EVIDENCE, THEN RAISE
'------------------------------------------------------------------------------
    'Save the certification log before anything is raised
        WriteRecords Path, mCertLog
    'Fail the caller when the certification did not pass
        If Verdict <> "PASS" Then
            Err.Raise ERR_TEST_FAILED, "TEST_Harness.EndCertification", _
                "Certification of " & mCertSuite & " did not pass; see " & Path & "."
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                                     LOG
'
'------------------------------------------------------------------------------
'

Public Sub Say( _
    ByVal Text As String)
'
'------------------------------------------------------------------------------
'                                 LOG A RECORD
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one log record to the Immediate window and keeps it for SaveLog
'
' WHY THIS EXISTS
'   The Immediate window keeps only its last 200 lines, so every record is also
'   kept in a collection that SaveLog writes in full
'
' INPUTS
'   Text
'     The record, written as KEY=value fields; String
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts the log when there is none, adds the record to it and prints it with
'   Echo
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   Echo
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' KEEP THE RECORD
'------------------------------------------------------------------------------
    'Start a log when none exists
        If mLog Is Nothing Then
            Set mLog = New Collection
        End If
    'Keep the record for SaveLog
        mLog.Add Text

'------------------------------------------------------------------------------
' PRINT THE RECORD
'------------------------------------------------------------------------------
    'Print the record to the Immediate window
        Echo Text

End Sub

Public Sub Echo( _
    ByVal Text As String)
'
'------------------------------------------------------------------------------
'                                 PRINT A LINE
'------------------------------------------------------------------------------
' PURPOSE
'   Prints one line to the Immediate window, without keeping it in the log
'
' WHY THIS EXISTS
'   A host without an Immediate window, such as LibreOffice run headless,
'   raises on Debug.Print, and the kept log, not the window, is the record of a
'   run
'
' INPUTS
'   Text
'     The line to print; String
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Prints the line with Debug.Print, with errors suppressed on that statement
'   only
'
' ERROR POLICY
'   A host without an Immediate window, such as LibreOffice run headless,
'   raises on Debug.Print; that failure is ignored, because the kept log, not
'   the window, is the record of a run
'
' DEPENDENCIES
'   Debug.Print
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRINT THE LINE
'------------------------------------------------------------------------------
    'Suppress the error of a host without an Immediate window, on the print statement only
        On Error Resume Next
    'Print the line
        Debug.Print Text
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Sub SaveLog( _
    ByVal Path As String)
'
'------------------------------------------------------------------------------
'                                 SAVE THE LOG
'------------------------------------------------------------------------------
' PURPOSE
'   Writes every record kept since the last BeginRun or ResetTests to a text
'   file
'
' WHY THIS EXISTS
'   The Immediate window keeps only its last 200 lines; the file is the
'   complete log
'
' INPUTS
'   Path
'     Full path of the file; String, created or overwritten; outside the
'     checkout or in the ignored test-results/ directory
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Refuses to save when no record is kept, then writes the records one per
'   line
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE when no record is kept; a file error is re-raised
'   after the file is closed
'
' DEPENDENCIES
'   The file writer for record collections
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE THE LOG
'------------------------------------------------------------------------------
    'Refuse to save when no record is kept
        If mLog Is Nothing Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.SaveLog", "No log record is kept."
        End If

'------------------------------------------------------------------------------
' WRITE THE FILE
'------------------------------------------------------------------------------
    'Write every kept record to the file
        WriteRecords Path, mLog

End Sub

Public Function LastResult() _
    As String
'
'------------------------------------------------------------------------------
'                              RETURN LAST RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   The RESULT record of the last finished or refused run
'
' WHY THIS EXISTS
'   A caller reads the outcome of a run from its RESULT record, which is kept
'   apart from the log
'
' INPUTS
'   None
'
' RETURNS
'   The RESULT record; an empty string before the first run, and after
'   CertifyRun has read it
'
' BEHAVIOR
'   Returns the RESULT record kept by BeginRun and EndRun
'
' ERROR POLICY
'   None
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
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return the kept RESULT record
        LastResult = mLastResult

End Function

'
'------------------------------------------------------------------------------
'
'                                   RESULTS
'
'------------------------------------------------------------------------------
'

Public Sub RecordResult( _
    ByVal CaseId As String, _
    ByVal VariantName As String, _
    ByVal Quantity As String, _
    ByVal Dataset As String, _
    ByVal Record As String, _
    ByVal Value As Double)
'
'------------------------------------------------------------------------------
'                               RECORD A RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Keeps one computed value of the running suite for SaveResults: the result
'   of one expected output of a test case, which tools/check_accuracy.py
'   --results compares with its expected value
'
' WHY THIS EXISTS
'   The accuracy check compares the values a test case computed with the
'   expected ones, so the suite records each value with the identity of its
'   output
'
' INPUTS
'   CaseId
'     The test case's id (tests/fixtures/); String
'
'   VariantName
'     The variant of a file-based case; String, an empty string for a JSON
'     fixture
'
'   Quantity
'     The output's quantity; String
'
'   Dataset
'     The dataset of a record-scoped output; String, an empty string for a
'     result of the whole calculation
'
'   Record
'     The record id of a record-scoped output; String, an empty string for a
'     result of the whole calculation
'
'   Value
'     The result; Double, written with the 15 significant digits of Str$
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Requires an active run and adds one JSON object, with the case, variant,
'   quantity, scope and value, to the recorded results
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A repeated result is kept; the
'   accuracy check rejects it
'
' DEPENDENCIES
'   RequireActiveRun, JsonText, JsonTextOrNull, ScopeJson and NumberText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REQUIRE A RUN
'------------------------------------------------------------------------------
    'Require an active run
        RequireActiveRun "RecordResult"

'------------------------------------------------------------------------------
' KEEP THE RESULT
'------------------------------------------------------------------------------
    'Keep one JSON object for the result
        mResults.Add "{""case"": " & JsonText(CaseId) & _
            ", ""variant"": " & JsonTextOrNull(VariantName) & _
            ", ""quantity"": " & JsonText(Quantity) & _
            ", ""scope"": " & ScopeJson(Dataset, Record) & _
            ", ""value"": " & NumberText(Value) & "}"

End Sub

Public Sub SaveResults( _
    ByVal Path As String)
'
'------------------------------------------------------------------------------
'                               SAVE THE RESULTS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the results recorded since the last BeginRun or ResetTests as the
'   results file of docs/methodology/TEST_CASES.md (Accuracy contracts), with
'   the commit of the workbook's build stamp
'
' WHY THIS EXISTS
'   The accuracy check reads the computed values of an Excel run from a file,
'   and the commit says which build produced them
'
' INPUTS
'   Path
'     Full path of the file; String, created or overwritten; outside the
'     checkout or in the ignored test-results/ directory
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Refuses to save before the first run, composes one JSON document with the
'   schema version, the kind, the commit and the recorded results, one per
'   line, and writes it to Path. The file is ASCII, a subset of UTF-8, with LF
'   line endings
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE before the first run; a file error is re-raised after
'   the file is closed. A workbook without a readable build stamp writes an
'   empty commit, and a build with local changes writes its marked commit; the
'   accuracy check rejects both, because the results then name no exact commit
'
' DEPENDENCIES
'   StampedCommit, JsonText and WriteText
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
    Dim Text    As String   'The whole file
    Dim Entry   As Variant  'One recorded result
    Dim Count   As Long     'Results written so far

'------------------------------------------------------------------------------
' COMPOSE THE DOCUMENT
'------------------------------------------------------------------------------
    'Refuse to save before any run has recorded results
        If mResults Is Nothing Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.SaveResults", "No run has recorded results."
        End If
    'Open the document with the schema version, the kind and the commit of the build stamp
        Text = "{" & vbLf & "  ""schema_version"": 1," & vbLf & _
            "  ""kind"": ""test-results""," & vbLf & _
            "  ""commit"": " & JsonText(StampedCommit()) & "," & vbLf & _
            "  ""results"": ["
    'Add one line for each recorded result
        For Each Entry In mResults
            Count = Count + 1
            Text = Text & IIf(Count = 1, vbNullString, ",") & vbLf & "    " & CStr(Entry)
        Next Entry
    'Close the list and the document
        Text = Text & IIf(Count = 0, vbNullString, vbLf & "  ") & "]" & vbLf & "}" & vbLf

'------------------------------------------------------------------------------
' WRITE THE FILE
'------------------------------------------------------------------------------
    'Write the document exactly as composed
        WriteText Path, Text

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub LogAssertion( _
    ByVal AssertionName As String, _
    ByVal Passed As Boolean, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                               LOG AN ASSERTION
'------------------------------------------------------------------------------
' PURPOSE
'   Counts one assertion, logs its outcome and records it when it failed
'
' WHY THIS EXISTS
'   Every assertion procedure ends in the same count, log record and failure
'   record, so they are written once
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case; String
'
'   Passed
'     True when the assertion holds; Boolean
'
'   Detail
'     Text logged and recorded with a failure; String
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Requires an active run, counts the assertion, records a failure for an
'   assertion made outside a case, then logs the ASSERT record and, when the
'   assertion failed, records the failure
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. An assertion outside a case is counted
'   and recorded as a failure
'
' DEPENDENCIES
'   RequireActiveRun, RecordFailure and Say
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COUNT THE ASSERTION
'------------------------------------------------------------------------------
    'Require an active run, then count the assertion
        RequireActiveRun "an assertion"
        mAssertionCount = mAssertionCount + 1
    'Record a failure for an assertion made outside a case
        If Len(mCurrentCase) = 0 Then
            RecordFailure "harness.no-case", "assertion '" & AssertionName & "' outside a case"
        End If

'------------------------------------------------------------------------------
' LOG THE OUTCOME
'------------------------------------------------------------------------------
    'Log the outcome, and record the failure when the assertion failed
        If Passed Then
            Say "ASSERT=" & mCurrentCase & "/" & AssertionName & "; status=PASS"
        Else
            Say "ASSERT=" & mCurrentCase & "/" & AssertionName & "; status=FAIL; " & Detail
            RecordFailure mCurrentCase & "/" & AssertionName, Detail
        End If

End Sub

Private Sub RecordFailure( _
    ByVal FailureId As String, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                               RECORD A FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Counts and logs one failure; the first one of a run is its primary failure
'
' WHY THIS EXISTS
'   Failures of every kind are counted in one place, and the first one is kept
'   to be reported with the summary
'
' INPUTS
'   FailureId
'     Identifier of the failure; String
'
'   Detail
'     Text of the failure; String
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts the failure, keeps it as the primary failure when it is the first of
'   the run and logs the FAILURE record
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   Say
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COUNT THE FAILURE
'------------------------------------------------------------------------------
    'Count the failure
        mFailureCount = mFailureCount + 1
    'Keep the first failure of the run as its primary failure
        If Len(mPrimaryFailure) = 0 Then
            mPrimaryFailure = FailureId & "; " & Detail
        End If

'------------------------------------------------------------------------------
' LOG THE FAILURE
'------------------------------------------------------------------------------
    'Write the FAILURE record
        Say "FAILURE=" & FailureId & "; " & Detail

End Sub

Private Sub RequireActiveRun( _
    ByVal Caller As String)
'
'------------------------------------------------------------------------------
'                            REQUIRE AN ACTIVE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Refuses a lifecycle call or an assertion made outside BeginRun..EndRun
'
' WHY THIS EXISTS
'   The lifecycle and assertion procedures all need a run in progress, and the
'   check is the same for each
'
' INPUTS
'   Caller
'     Name of the call that needs the run, written in the message; String
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Raises ERR_TEST_USAGE when no run is active
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE, with the source TEST_Harness, when no run is active
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
' CHECK THE RUN
'------------------------------------------------------------------------------
    'Refuse a call made while no run is active
        If Not mRunActive Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness", _
                "No test run is active for " & Caller & "; call BeginRun first."
        End If


End Sub

Private Sub RequireCertification( _
    ByVal Caller As String)
'
'------------------------------------------------------------------------------
'                           REQUIRE A CERTIFICATION
'------------------------------------------------------------------------------
' PURPOSE
'   Refuses a certification call made outside
'   BeginCertification..EndCertification
'
' WHY THIS EXISTS
'   The certification procedures all need a certification in progress, and the
'   check is the same for each
'
' INPUTS
'   Caller
'     Name of the call that needs the certification, written in the message;
'     String
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Raises ERR_TEST_USAGE when no certification is active
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE, with the source TEST_Harness, when no certification
'   is active
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK THE CERTIFICATION
'------------------------------------------------------------------------------
    'Refuse a call made while no certification is active
        If Not mCertActive Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness", _
                "No certification is active for " & Caller & _
                "; call BeginCertification first."
        End If

End Sub

Private Sub CertSay( _
    ByVal Text As String)
'
'------------------------------------------------------------------------------
'                          LOG A CERTIFICATION RECORD
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one certification record to the Immediate window and keeps it for
'   EndCertification
'
' WHY THIS EXISTS
'   Certification records are kept apart from the run logs, so that the
'   certification log holds only them
'
' INPUTS
'   Text
'     The record, written as KEY=value fields; String
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds the record to the certification log and prints it with Echo
'
' ERROR POLICY
'   None; it relies on the certification log that BeginCertification starts
'
' DEPENDENCIES
'   Echo
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' KEEP THE RECORD
'------------------------------------------------------------------------------
    'Keep the record for EndCertification
        mCertLog.Add Text

'------------------------------------------------------------------------------
' PRINT THE RECORD
'------------------------------------------------------------------------------
    'Print the record to the Immediate window
        Echo Text

End Sub

Private Sub WriteRecords( _
    ByVal Path As String, _
    ByVal Records As Collection)
'
'------------------------------------------------------------------------------
'                           WRITE RECORDS TO A FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes records to a text file, one per line
'
' WHY THIS EXISTS
'   SaveLog and EndCertification write their records the same way, and the file
'   must be closed whatever happens while it is written
'
' INPUTS
'   Path
'     Full path of the file; String, created or overwritten
'
'   Records
'     The records, in order; Collection of values written with CStr
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens the file for output, prints each record on its own line and closes
'   the file. When a write fails, saves the error, closes the file and raises
'   the saved error again
'
' ERROR POLICY
'   A file error is re-raised after the file is closed. An error of the open
'   statement itself is not trapped here
'
' DEPENDENCIES
'   VBA file statements: FreeFile, Open, Print # and Close
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
    Dim FileNumber          As Integer  'Handle of the open file
    Dim Entry               As Variant  'One record
    Dim ErrorNumber         As Long     'File error number
    Dim ErrorSource         As String   'File error source
    Dim ErrorDescription    As String   'File error description

'------------------------------------------------------------------------------
' OPEN THE FILE
'------------------------------------------------------------------------------
    'Take a free file number and open the file for output
        FileNumber = FreeFile
        Open Path For Output As #FileNumber
    'Trap file errors from here on
        On Error GoTo WriteFailed

'------------------------------------------------------------------------------
' WRITE THE RECORDS
'------------------------------------------------------------------------------
    'Print each record on its own line
        For Each Entry In Records
            Print #FileNumber, CStr(Entry)
        Next Entry
    'Close the file
        Close #FileNumber

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the handler for a write error
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
WriteFailed:
    'Save the file error before any other call
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Close the file
        Close #FileNumber
    'Raise the saved error again
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub

Private Sub WriteText( _
    ByVal Path As String, _
    ByVal Text As String)
'
'------------------------------------------------------------------------------
'                             WRITE TEXT TO A FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes text to a file exactly as given, without a line ending added
'
' WHY THIS EXISTS
'   SaveResults writes a document whose line endings are part of its content,
'   so none may be added
'
' INPUTS
'   Path
'     Full path of the file; String, created or overwritten
'
'   Text
'     The text to write; String, ASCII text
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens the file for output, prints the text with a trailing semicolon and
'   closes the file. When the write fails, saves the error, closes the file and
'   raises the saved error again
'
' ERROR POLICY
'   A file error is re-raised after the file is closed. An error of the open
'   statement itself is not trapped here
'
' DEPENDENCIES
'   VBA file statements: FreeFile, Open, Print # and Close
'
' NOTES
'   The trailing semicolon keeps Print from adding CR LF
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FileNumber          As Integer  'Handle of the open file
    Dim ErrorNumber         As Long     'File error number
    Dim ErrorSource         As String   'File error source
    Dim ErrorDescription    As String   'File error description

'------------------------------------------------------------------------------
' OPEN THE FILE
'------------------------------------------------------------------------------
    'Take a free file number and open the file for output
        FileNumber = FreeFile
        Open Path For Output As #FileNumber
    'Trap file errors from here on
        On Error GoTo WriteFailed

'------------------------------------------------------------------------------
' WRITE THE TEXT
'------------------------------------------------------------------------------
    'Print the text with a trailing semicolon, which keeps Print from adding CR LF
        Print #FileNumber, Text;
    'Close the file
        Close #FileNumber

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the handler for a write error
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
WriteFailed:
    'Save the file error before any other call
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Close the file
        Close #FileNumber
    'Raise the saved error again
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub

Private Function StampedCommit() _
    As String
'
'------------------------------------------------------------------------------
'                            READ THE BUILD COMMIT
'------------------------------------------------------------------------------
' PURPOSE
'   The commit of the workbook's build stamp: the value of the defined name
'   STAMP_COMMIT_NAME in ThisWorkbook, the workbook that holds this module
'
' WHY THIS EXISTS
'   SaveResults writes the commit into the results file, so that the accuracy
'   check can tell which build produced the results
'
' INPUTS
'   None
'
' RETURNS
'   The stamp as written by tools/devbuild/build_workbook.py, or an empty
'   string when the workbook has no such name or it cannot be read
'
' BEHAVIOR
'   Reads the cell that the defined name refers to and returns its value as
'   text
'
' ERROR POLICY
'   Never raises: any error while reading returns an empty string
'
' DEPENDENCIES
'   ThisWorkbook.Names and the constant STAMP_COMMIT_NAME
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Divert a failed read to the handler below
        On Error GoTo ReadFailed

'------------------------------------------------------------------------------
' READ THE COMMIT
'------------------------------------------------------------------------------
    'Read the value of the defined name in ThisWorkbook as text
        StampedCommit = CStr(ThisWorkbook.Names(STAMP_COMMIT_NAME).RefersToRange.Value)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the handler for a failed read
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ReadFailed:
    'Return an empty string when the name is missing or cannot be read
        StampedCommit = vbNullString

End Function

Private Function JsonText( _
    ByVal Value As String) _
    As String
'
'------------------------------------------------------------------------------
'                               QUOTE JSON TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   A JSON string literal of Value in ASCII: the quote, the backslash, the
'   control characters and every character above code 126 are escaped
'
' WHY THIS EXISTS
'   The results file is ASCII, a subset of UTF-8, so every character outside
'   printable ASCII is written as a \u escape
'
' INPUTS
'   Value
'     The text to quote; String
'
' RETURNS
'   The literal, with its surrounding double quotes
'
' BEHAVIOR
'   Walks the characters of Value and escapes each one: the quote and the
'   backslash with a backslash, and a code below 32 or above 126 as \u and four
'   hexadecimal digits of its UTF-16 code unit
'
' ERROR POLICY
'   None
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
    Dim Index       As Long     'Position in Value
    Dim Character   As String   'Character at Index
    Dim Code        As Long     'Its UTF-16 code unit, 0 to 65535
    Dim Text        As String   'Escaped text so far

'------------------------------------------------------------------------------
' ESCAPE THE VALUE
'------------------------------------------------------------------------------
    'Escape the value character by character
        For Index = 1 To Len(Value)
            Character = Mid$(Value, Index, 1)
            Code = AscW(Character) And &HFFFF&
            'Escape a quote or a backslash with a backslash, a control or high code as a u escape
            If Character = """" Or Character = "\" Then
                Text = Text & "\" & Character
            ElseIf Code < 32 Or Code > 126 Then
                Text = Text & "\u" & Right$("000" & Hex$(Code), 4)
            Else
                Text = Text & Character
            End If
        Next Index

'------------------------------------------------------------------------------
' QUOTE THE RESULT
'------------------------------------------------------------------------------
    'Wrap the escaped text in double quotes
        JsonText = """" & Text & """"

End Function

Private Function JsonTextOrNull( _
    ByVal Value As String) _
    As String
'
'------------------------------------------------------------------------------
'                           QUOTE JSON TEXT OR NULL
'------------------------------------------------------------------------------
' PURPOSE
'   JSON null for an empty string, otherwise the string literal of Value
'
' WHY THIS EXISTS
'   An empty variant name is written as null in the results file, and any other
'   text as a string literal
'
' INPUTS
'   Value
'     The text to quote; String, may be empty
'
' RETURNS
'   The text null for an empty string; otherwise the result of JsonText
'
' BEHAVIOR
'   Tests for an empty string and returns null or the quoted text
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   JsonText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHOOSE THE LITERAL
'------------------------------------------------------------------------------
    'Return null for an empty string, otherwise the string literal
        If Len(Value) = 0 Then
            JsonTextOrNull = "null"
        Else
            JsonTextOrNull = JsonText(Value)
        End If

End Function

Private Function ScopeJson( _
    ByVal Dataset As String, _
    ByVal Record As String) _
    As String
'
'------------------------------------------------------------------------------
'                             FORMAT RESULT SCOPE
'------------------------------------------------------------------------------
' PURPOSE
'   The scope of a result as the expected-results format writes it: null for
'   the whole calculation, otherwise {"dataset": ..., "record": ...}, the only
'   shape tools/check_accuracy.py matches with a record-scoped output
'
' WHY THIS EXISTS
'   A result is either for the whole calculation or for one record of a
'   dataset, and the accuracy check matches a record-scoped output only in this
'   shape
'
' INPUTS
'   Dataset
'     The dataset of a record-scoped output; String, empty for the whole
'     calculation
'
'   Record
'     The record id of a record-scoped output; String, empty for the whole
'     calculation
'
' RETURNS
'   The text null when both are empty; otherwise a JSON object with the dataset
'   and the record
'
' BEHAVIOR
'   Tests whether both arguments are empty and returns null or the object
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   JsonText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHOOSE THE SCOPE
'------------------------------------------------------------------------------
    'Return null for the whole calculation, otherwise an object with the dataset and the record
        If Len(Dataset) = 0 And Len(Record) = 0 Then
            ScopeJson = "null"
        Else
            ScopeJson = "{""dataset"": " & JsonText(Dataset) & ", ""record"": " & _
                JsonText(Record) & "}"
        End If

End Function

Private Function ActiveWindowCaption() _
    As String
'
'------------------------------------------------------------------------------
'                          READ ACTIVE WINDOW CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   The caption of Excel's active window, or an empty string when there is
'   none, so that a run which leaves another window active is seen
'
' WHY THIS EXISTS
'   VerifyHostState compares the caption with the one read at BeginRun
'
' INPUTS
'   None
'
' RETURNS
'   The caption; an empty string when Excel has no active window
'
' BEHAVIOR
'   Returns an empty string when Application.ActiveWindow is Nothing, otherwise
'   its caption
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   Application.ActiveWindow
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ THE CAPTION
'------------------------------------------------------------------------------
    'Return an empty string when Excel has no active window, otherwise its caption
        If Application.ActiveWindow Is Nothing Then
            ActiveWindowCaption = vbNullString
        Else
            ActiveWindowCaption = Application.ActiveWindow.Caption
        End If

End Function

Private Function VerifyHostState( _
    ByRef Detail As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                              VERIFY HOST STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Compares the four Application properties, the number of open workbooks and
'   the caption of the active window with their BeginRun values
'
' WHY THIS EXISTS
'   A run must leave Excel as it found it. The harness only checks this: it
'   never changed these settings, and the procedure that did owns their
'   restoration
'
' INPUTS
'   Detail
'     Output, String: names the changed properties, or says why they could not
'     be compared
'
' RETURNS
'   True when all are unchanged; False when one changed, when the snapshot of
'   BeginRun is not valid or when a read fails
'
' BEHAVIOR
'   Refuses to compare an invalid snapshot, then compares Calculation,
'   DisplayAlerts, EnableEvents, ScreenUpdating, the number of workbooks and
'   the active window caption, and describes the outcome in Detail
'
' ERROR POLICY
'   A failed read is reported in Detail and returns False; nothing is raised
'
' DEPENDENCIES
'   ActiveWindowCaption; the Application properties and the snapshot taken by
'   BeginRun
'
' NOTES
'   The harness never restores the properties
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Changed As String   'Names of the changed properties

'------------------------------------------------------------------------------
' REFUSE AN UNREAD SNAPSHOT
'------------------------------------------------------------------------------
    'Report a snapshot that could not be read at BeginRun as not compared
        If Not mSnapshotValid Then
            Detail = "Excel state not compared: it could not be read at BeginRun"
            VerifyHostState = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' COMPARE THE HOST STATE
'------------------------------------------------------------------------------
    'Divert a failed read to the handler below
        On Error GoTo ReadFailed
    'Compare each property with its BeginRun value and name the ones that changed
        If Application.Calculation <> mInitialCalculation Then Changed = Changed & " Calculation"
        If Application.DisplayAlerts <> mInitialAlerts Then Changed = Changed & " DisplayAlerts"
        If Application.EnableEvents <> mInitialEvents Then Changed = Changed & " EnableEvents"
        If Application.ScreenUpdating <> mInitialScreen Then Changed = Changed & " ScreenUpdating"
        If Application.Workbooks.Count <> mInitialWorkbooks Then Changed = Changed & " Workbooks"
        If ActiveWindowCaption() <> mInitialWindow Then Changed = Changed & " ActiveWindow"
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' REPORT THE OUTCOME
'------------------------------------------------------------------------------
    'Describe the outcome: unchanged, or the names of the changed properties
        If Len(Changed) = 0 Then
            Detail = "Excel state unchanged (Calculation, DisplayAlerts, EnableEvents, " & _
                "ScreenUpdating, Workbooks, ActiveWindow)"
            VerifyHostState = True
        Else
            Detail = "Excel state changed during the run:" & Changed
            VerifyHostState = False
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the handler for a failed read
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ReadFailed:
    'Report that the state was not compared because a read failed
        Detail = "Excel state not compared: read failed with error " & CStr(Err.Number) & _
            " (" & Err.Description & ")"
        VerifyHostState = False

End Function

Private Function IsCaseId( _
    ByVal CaseId As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                             TEST CASE IDENTIFIER
'------------------------------------------------------------------------------
' PURPOSE
'   Tests whether CaseId is lower-case kebab-case: a-z and 0-9 in groups joined
'   by single hyphens, as the test-case files use (TEST_CASES.md, Files)
'
' WHY THIS EXISTS
'   BeginCase records a failure for an identifier that does not have the form
'   of the test-case files
'
' INPUTS
'   CaseId
'     The identifier to test; String
'
' RETURNS
'   True for a well-formed identifier; False for an empty one, one that starts
'   or ends with a hyphen, one with two hyphens in a row and one with any
'   character other than a-z, 0-9 and the hyphen
'
' BEHAVIOR
'   Tests the hyphens first, then each character with the pattern [a-z0-9-]
'
' ERROR POLICY
'   None
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
' DECLARE
'------------------------------------------------------------------------------
    Dim Character   As String   'Character under test
    Dim i           As Long     'Position in CaseId

'------------------------------------------------------------------------------
' CHECK THE HYPHENS
'------------------------------------------------------------------------------
    'Reject an empty identifier and one that starts or ends with a hyphen
        If Len(CaseId) = 0 Or Left$(CaseId, 1) = "-" Or Right$(CaseId, 1) = "-" Then
            IsCaseId = False
            Exit Function
        End If
    'Reject two hyphens in a row
        If InStr(CaseId, "--") > 0 Then
            IsCaseId = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' CHECK THE CHARACTERS
'------------------------------------------------------------------------------
    'Reject any character other than a to z, 0 to 9 and the hyphen
        For i = 1 To Len(CaseId)
            Character = Mid$(CaseId, i, 1)
            If Not (Character Like "[a-z0-9-]") Then
                IsCaseId = False
                Exit Function
            End If
        Next i

'------------------------------------------------------------------------------
' ACCEPT THE IDENTIFIER
'------------------------------------------------------------------------------
    'Accept an identifier that passed every test
        IsCaseId = True

End Function

Private Function IsKnownCase( _
    ByVal CaseId As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                               TEST KNOWN CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Tests whether CaseId already started in this run
'
' WHY THIS EXISTS
'   BeginCase records a failure for a repeated identifier
'
' INPUTS
'   CaseId
'     The identifier to look up; String
'
' RETURNS
'   True when the identifier is already in the run's collection of case
'   identifiers
'
' BEHAVIOR
'   Looks the identifier up in the collection with errors suppressed on that
'   statement only, and reports the case as known when the lookup raised
'   nothing
'
' ERROR POLICY
'   A Collection raises error 5 for an unknown key; that error is the answer,
'   and nothing is raised
'
' DEPENDENCIES
'   The collection of case identifiers kept by ResetRunState
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
    Dim Found   As Variant  'Item stored under CaseId, when present

'------------------------------------------------------------------------------
' LOOK UP THE CASE
'------------------------------------------------------------------------------
    'Look the identifier up on this statement only: an unknown key raises error 5, and that is the answer
        On Error Resume Next
        Found = mCaseIds.Item(CaseId)
    'Report the case as known when the lookup raised nothing
        IsKnownCase = (Err.Number = 0)
    'Restore normal error handling
        On Error GoTo 0

End Function

Private Function NumberText( _
    ByVal Value As Double) _
    As String
'
'------------------------------------------------------------------------------
'                               FORMAT A NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   A Double as text that does not depend on the regional settings: Str$ always
'   uses a period as the decimal separator. Str$ drops the zero before the
'   point (" .75"), which is put back ("0.75", "-0.75")
'
' WHY THIS EXISTS
'   Log records and the results file must read the same in every regional
'   setting
'
' INPUTS
'   Value
'     The number to write; Double
'
' RETURNS
'   The number as text, with the 15 significant digits of Str$, a period as the
'   decimal separator and no leading space
'
' BEHAVIOR
'   Converts the number with Str$, trims its sign space and puts back the zero
'   that Str$ drops before the point
'
' ERROR POLICY
'   None
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
' DECLARE
'------------------------------------------------------------------------------
    Dim Text    As String   'Str$ result without its sign space

'------------------------------------------------------------------------------
' CONVERT THE NUMBER
'------------------------------------------------------------------------------
    'Convert with Str$, which uses a period whatever the regional settings, and drop the sign space
        Text = Trim$(Str$(Value))

'------------------------------------------------------------------------------
' RESTORE THE LEADING ZERO
'------------------------------------------------------------------------------
    'Put back the zero that Str$ drops before the point, for a positive and a negative number
        If Left$(Text, 1) = "." Then
            Text = "0" & Text
        ElseIf Left$(Text, 2) = "-." Then
            Text = "-0" & Mid$(Text, 2)
        End If

'------------------------------------------------------------------------------
' RETURN THE TEXT
'------------------------------------------------------------------------------
    'Return the text
        NumberText = Text

End Function

Private Function HostProperty( _
    ByVal PropertyName As String) _
    As String
'
'------------------------------------------------------------------------------
'                              READ HOST PROPERTY
'------------------------------------------------------------------------------
' PURPOSE
'   One Application property as text, or an empty string when the host cannot
'   read it
'
' WHY THIS EXISTS
'   EnvironmentSummary describes the host in the log, and a host that lacks a
'   property must still start the run
'
' INPUTS
'   PropertyName
'     Name, String, of the property to read: Name, Version, Build or
'     OperatingSystem
'
' RETURNS
'   The property as text; an empty string when the host cannot read it or the
'   name is not one of the four
'
' BEHAVIOR
'   Selects the property by name and returns it as text
'
' ERROR POLICY
'   A failed read returns an empty string; nothing is raised
'
' DEPENDENCIES
'   Application.Name, Version, Build and OperatingSystem
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Divert a failed read to the handler below
        On Error GoTo ReadFailed

'------------------------------------------------------------------------------
' READ THE PROPERTY
'------------------------------------------------------------------------------
    'Return the property that the name selects
        Select Case PropertyName
            Case "Name"
                HostProperty = Application.Name
            Case "Version"
                HostProperty = Application.Version
            Case "Build"
                HostProperty = CStr(Application.Build)
            Case "OperatingSystem"
                HostProperty = Application.OperatingSystem
        End Select

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the handler for a failed read
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ReadFailed:
    'Return an empty string when the host cannot read the property
        HostProperty = vbNullString

End Function

Private Function EnvironmentSummary() _
    As String
'
'------------------------------------------------------------------------------
'                             DESCRIBE ENVIRONMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Describes the Excel host and the compiled VBA environment
'
' WHY THIS EXISTS
'   The log records the host, its version, the Office bitness and the VBA
'   generation, so that the evidence of a run says where it ran
'
' INPUTS
'   None
'
' RETURNS
'   host, version, build, os, office (bitness) and runtime (VBA generation) as
'   one KEY=value text; the last two come from conditional compilation, not
'   from Windows. A host property that cannot be read is left empty, so that
'   the run still starts and tools/check_excel_run.py rejects the evidence
'
' BEHAVIOR
'   Takes the bitness and the VBA generation from conditional compilation, then
'   joins them with the host properties read by HostProperty
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   HostProperty
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
    Dim Bitness As String   'Office bitness
    Dim Runtime As String   'VBA generation

'------------------------------------------------------------------------------
' DESCRIBE THE ENVIRONMENT
'------------------------------------------------------------------------------
    'Take the Office bitness from conditional compilation
#If Win64 Then
        Bitness = "64-bit"
#Else
        Bitness = "32-bit"
#End If
    'Take the VBA generation from conditional compilation
#If VBA7 Then
        Runtime = "VBA7"
#Else
        Runtime = "VBA6"
#End If
    'Join the host properties, the Office bitness and the VBA generation
        EnvironmentSummary = "host=" & HostProperty("Name") & "; version=" & HostProperty("Version") & _
            "; build=" & HostProperty("Build") & "; os=" & HostProperty("OperatingSystem") & _
            "; office=" & Bitness & "; runtime=" & Runtime

End Function

Private Sub ResetRunState()

'
'------------------------------------------------------------------------------
'                               RESET RUN STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the counters, case identifiers, failures and flags of a run to
'   their idle values; the recovery flag and the log are left to callers
'
' WHY THIS EXISTS
'   BeginRun and ResetTests both start from the same idle state of a run
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the active flag, the suite name, the expected and observed counts,
'   the current case, the primary failure and the snapshot flag, and starts an
'   empty collection of case identifiers
'
' ERROR POLICY
'   None
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
' CLEAR THE RUN
'------------------------------------------------------------------------------
    'Clear the active flag, the suite name and the expected counts
        mRunActive = False
        mSuite = vbNullString
        mExpectedCases = 0
        mExpectedAssertions = 0
    'Clear the observed counts
        mCaseCount = 0
        mAssertionCount = 0
        mFailureCount = 0
        mErrorCount = 0
    'Clear the current case and start an empty collection of case identifiers
        mCurrentCase = vbNullString
        Set mCaseIds = New Collection
    'Clear the primary failure and the snapshot flag
        mPrimaryFailure = vbNullString
        mSnapshotValid = False

End Sub
