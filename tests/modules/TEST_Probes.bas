Attribute VB_Name = "TEST_Probes"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Probes
'------------------------------------------------------------------------------
' PURPOSE
'   The probes suite of the workbook's tests: what the prototype relies on
'   Excel to do, as cases and assertions of TEST_Harness. It reads an error
'   value from an input cell, writes a text into a cell whose number format is
'   not text, takes a range as a table, checks the status line on the other
'   output sheets, checks the save state after the status check and checks that
'   the template protects its sheets as the workbook relies on
'
' WHY THIS EXISTS
'   The prototype relies on behaviours of the host that no calculation shows.
'   Each case is the probe of a behaviour of the Host behaviour table of
'   docs/methodology/VALIDATION_PLAN.md. A pass measures the behaviour on the
'   host that ran it; it validates no rule of the calculation
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Runs the seven cases in a fixed order between BeginRun and EndRun of
'   TEST_Harness, after it has allowed writing to every sheet:
'     - CaseErrorValuesInInputs: an error value typed in an input cell reads as
'       the text #ERROR where an output shows that text, and no output cell
'       holds an error value (HB-08)
'     - CaseReadingAnErrorValue: a cell given an error value holds one, and the
'       text #N/A written into a cell formatted as text stays a text (HB-08)
'     - CaseTextWrittenIntoACell: a text that reads as a number becomes a
'       number in a cell that is not formatted as text, and keeps its digits in
'       a cell formatted as text (HB-09)
'     - CaseAreasAsTables: the Value and the Formula of a range of more than
'       one cell are tables with bounds from 1, and those of one cell are not
'     - CaseSheetProtection: every sheet is protected, only the input cells are
'       unlocked, and a write to a locked cell is refused until AllowWriting
'       has been applied to the sheet (HB-15)
'     - CaseStatusOnTheOtherSheets: E1 of each calculation sheet and of Checks
'       shows the status line of Results (HB-10)
'     - CaseSaveStateAfterTheCheck: CheckStatus_OnEvent leaves a workbook with
'       no unsaved change as it was when the results are current, and leaves an
'       unsaved change when it writes OUTDATED: before the status line (HB-11)
'
'   Entry points:
'     - RunProbeTests is the documented test entry point
'     - Test_RunProbes runs the same suite and returns its result line where a
'       caller cannot read an error, such as tools/devbuild/lo_macro.py
'     - CertifyProbes certifies the run and saves its log and its certification
'       log as evidence of an Excel run
'
'   RunSuite, the cases and the helpers are private
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED, or
'   the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START. An unexpected error inside a case is
'   recorded as a failure of that case, and the run continues. Test_RunProbes
'   contains those errors; CertifyProbes records them, saves the log, and
'   raises only when the certification did not pass
'
' DEPENDENCIES
'   TEST_Harness; M_Main, CORE_Config and CORE_Util; the workbook's names and
'   sheets, with the sample data the template ships
'
'   Reference: the expected texts and numbers of error-values-in-inputs are
'   those of error_keys and error_values of tools/prototype/verify.py, which
'   checks the formula workbook of the same rows; host behaviours HB-08, HB-09,
'   HB-10, HB-11 and HB-15 of the Host behaviour table of
'   docs/methodology/VALIDATION_PLAN.md
'
' NOTES
'   Option Explicit requires explicit declarations; Option Private Module keeps
'   the tests inside this project, out of the workbook's external surface.
'   Never part of the production import set
'
'   State ownership: owns the rows it appends under the four position tables,
'   which it removes at the end of each case that appends them, and the rows
'   each table had before. Changes ScalingIR and ThisWorkbook.Saved for one
'   check and puts them back. Rewrites the output sheets through RunSSA_Silent,
'   and the last case leaves them cleared
'
'   Usage: in a development workbook, type TEST_Probes.RunProbeTests in the
'   Immediate window. To record an Excel run, type the call below
'   (INSTALLATION.md, Recording an Excel run):
'
'       TEST_Probes.CertifyProbes "<folder>", "<checkout>\tests\fixtures"
'
'   Compatibility: Excel VBA. LibreOffice's VBA mode runs it, and differs from
'   Excel on the probes of an error value in a cell: tools/devbuild/lo_smoke.py
'   lists the assertions it is known to fail there. That is never Excel
'   evidence
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-19
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME          As String = "probes"  'Name of the suite in the logs and the certification
    Private Const EXPECTED_CASES      As Long = 7           'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS As Long = 49          'Assertions a complete run evaluates

    'The error values a cell holds: the codes of CVErr
    Private Const ERR_NA              As Long = 2042        '#N/A
    Private Const ERR_VALUE           As Long = 2015        '#VALUE!
    Private Const ERR_REF             As Long = 2023        '#REF!

    'The error Excel raises for a write to a locked cell of a protected sheet
    Private Const ERR_PROTECTED_WRITE As Long = 1004        'Run-time error of a write to a locked cell

    'The text that stands for an error value in an output
    Private Const ERROR_TEXT          As String = "#ERROR"  'Text of an error value in an output

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mOriginal(1 To 4)   As Long     'Rows under the header of a table before the case
    Private mFirst(1 To 4)      As Long     'Row, under the header, of the first row appended
    Private mAppended(1 To 4)   As Long     'Rows appended to the table


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunProbeTests()

'
'------------------------------------------------------------------------------
'                               RUN PROBES SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Run the probes suite. A passing run ends with RESULT=PASS; suite=probes;
'   operation=PASS; completeness=COMPLETE; cases=7; assertions=49; failures=0;
'   cleanup=PASS; recovery=NONE
'
' WHY THIS EXISTS
'   It is the documented test entry point of the suite: a run that does not
'   pass raises, so that a person typing it in the Immediate window sees the
'   failure
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls RunSuite, which runs every case between BeginRun and EndRun of
'   TEST_Harness
'
' ERROR POLICY
'   As TEST_Harness.EndRun: raises when the run does not pass. The procedure
'   has no error handler
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   State ownership: rewrites the output sheets and leaves them cleared
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run every case and raise when the run does not pass
        RunSuite

End Sub

Public Function Test_RunProbes() As String

'
'------------------------------------------------------------------------------
'                             RETURN PROBES RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Run the probes suite for a caller that reads a result and not an error
'
' WHY THIS EXISTS
'   A caller such as tools/devbuild/lo_macro.py cannot read an error: it reads
'   the result line of the run
'
' INPUTS
'   None
'
' RETURNS
'   The RESULT line of the run, as TEST_Harness.LastResult gives it
'
' BEHAVIOR
'   Runs RunSuite with errors suppressed, restores normal error handling and
'   returns the result line that the harness holds
'
' ERROR POLICY
'   Contains the error of a run that does not pass, or of a refused start: the
'   result line says so
'
' DEPENDENCIES
'   RunSuite, TEST_Harness.LastResult
'
' NOTES
'   State ownership: rewrites the output sheets and leaves them cleared
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Contain the error of a run that does not pass, or of a refused start
        On Error Resume Next
        RunSuite
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return the result line that the harness holds
        Test_RunProbes = TEST_Harness.LastResult()

End Function

Public Sub CertifyProbes( _
    ByVal Folder As String, _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                              CERTIFY PROBES RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Certify the probes suite's run as evidence of an Excel run: run it, save
'   its log as probes.log, and write probes-certification.log
'
' WHY THIS EXISTS
'   An Excel run is recorded as evidence by its log and its certification log:
'   this entry point runs the suite and writes both
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
'   In order:
'     - begins the certification of the suite
'     - runs RunSuite and keeps the number and the description of its error
'     - saves the log as probes.log in Folder and writes a SAVED record
'     - certifies the run as NORMAL and writes probes-certification.log in
'       Folder
'
' ERROR POLICY
'   The error of the run is recorded, not raised, so that the log is always
'   saved. Raises ERR_TEST_FAILED after writing the certification log when the
'   certification did not pass. An error of BeginCertification or SaveLog
'   propagates; run TEST_Harness.ResetTests before trying again
'
' DEPENDENCIES
'   RunSuite; TEST_Harness (BeginCertification, SaveLog, Echo, CertifyRun and
'   EndCertification); SUITE_NAME
'
' NOTES
'   State ownership: rewrites the output sheets and leaves them cleared
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RunErrorNumber      As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription As String   'Its description

'------------------------------------------------------------------------------
' BEGIN CERTIFICATION
'------------------------------------------------------------------------------
    'Open the certification of the suite
        TEST_Harness.BeginCertification SUITE_NAME

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Keep the error of the run instead of raising it
        On Error Resume Next
        RunSuite
    'Capture the error number and description
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' SAVE LOG
'------------------------------------------------------------------------------
    'Save the log of the run
        TEST_Harness.SaveLog Folder & "\probes.log"
    'Record where the log was saved and how the run ended
        TEST_Harness.Echo "SAVED=" & Folder & "\probes.log; mode=NORMAL; error=" & _
            CStr(RunErrorNumber) & IIf(RunErrorNumber = 0, vbNullString, " " & RunErrorDescription)

'------------------------------------------------------------------------------
' CERTIFY RUN
'------------------------------------------------------------------------------
    'Certify the run, then write the certification log
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
        TEST_Harness.EndCertification Folder & "\probes-certification.log"

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
'                                RUN EVERY CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. The order
'   matters: the last case clears the output sheets
'
' WHY THIS EXISTS
'   The three entry points run the same suite, so the runner is written once
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - begins the run with the suite name and the expected counts
'     - allows writing to every sheet
'     - runs the seven cases
'     - ends the run with the error that escaped a case, if any
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it and
'   re-raises it unchanged. BeginRun's refusal propagates directly
'
' DEPENDENCIES
'   TEST_Harness (BeginRun and EndRun); AllowWritingToSheets; the seven Case
'   procedures; SUITE_NAME, EXPECTED_CASES and EXPECTED_ASSERTIONS
'
' NOTES
'   None
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' BEGIN RUN
'------------------------------------------------------------------------------
    'Open the run, which a dirty start refuses
        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Record an error that escapes a case, to pass it to the end of the run
        On Error GoTo RunFailed
    'Allow the suite to write the sheets the template protects
        AllowWritingToSheets
    'Run the cases in their fixed order
        CaseErrorValuesInInputs
        CaseReadingAnErrorValue
        CaseTextWrittenIntoACell
        CaseAreasAsTables
        CaseSheetProtection
        CaseStatusOnTheOtherSheets
        CaseSaveStateAfterTheCheck

'------------------------------------------------------------------------------
' END RUN
'------------------------------------------------------------------------------
Finish:
    'Restore normal error handling
        On Error GoTo 0
    'End the run with the error that escaped a case, if any
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
    'Keep the error of the case runner
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Resume at the end of the run
        Resume Finish

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseErrorValuesInInputs()

'
'------------------------------------------------------------------------------
'                         PROBE ERROR VALUES IN INPUTS
'------------------------------------------------------------------------------
' PURPOSE
'   An error value typed in an input cell, as an Issue ID, a currency, a
'   PositionID or a market, a commodity and an amount code, reads as the text
'   #ERROR where an output shows that text, and no output cell holds an error
'   value. HB-08
'
' WHY THIS EXISTS
'   The inputs are cells that a user types in, and a typed error value must not
'   reach an output as an error value. The case measures that on the host
'
' INPUTS
'   None
'
' RETURNS
'   Nothing; the run's own result is not compared
'
' BEHAVIOR
'   In order:
'     - appends 3, 3, 2 and 1 rows under the interest rate, equity, commodity
'       and foreign exchange position tables
'     - writes rows in which one cell holds an error value given by CVErr
'     - runs the calculation and asserts the #ERROR issue keys, the net of the
'       issue, the currency text and status, the PositionID text and that no
'       output cell holds an error value
'     - removes the appended rows and asserts that each table has the rows it
'       had
'
' ERROR POLICY
'   An error of the case is recorded with RecordCaseError of TEST_Harness and
'   the case resumes at its cleanup, which runs on the success and the error
'   path
'
'   RemoveAppended contains its own error and returns it as a text, so that the
'   removal is asserted on both paths
'
' DEPENDENCIES
'   Appending and writing: AppendRows, PutRow, RunSSA_Silent and
'   RemoveAppended; reading: RowCell, PlainText, NumberOf and ErrorCells;
'   TEST_Harness (BeginCase, the assertions and RecordCaseError)
'
'   Reference: rows of error_keys of tools/prototype/verify.py. Two Issue IDs
'   that both read #ERROR are one issue: its key is #ERROR and the net of its
'   long 3,000,000 and short 1,000,000 is 2,000,000
'
' NOTES
'   State ownership: appends rows under the position tables and removes them at
'   the end
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultText  As String   'RunSSA_Silent result line, not compared
    Dim RestoreDue  As Boolean  'True once rows are appended
    Dim Detail      As String   'Why the appended rows could not be removed

'------------------------------------------------------------------------------
' START CASE
'------------------------------------------------------------------------------
    'Record an error of the case instead of ending the run, and name the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "error-values-in-inputs"
    'Remember that the cleanup has rows to remove
        RestoreDue = True

'------------------------------------------------------------------------------
' APPEND ROWS
'------------------------------------------------------------------------------
    'Make room under each of the four position tables
        AppendRows 1, 3
        AppendRows 2, 3
        AppendRows 3, 2
        AppendRows 4, 1

'------------------------------------------------------------------------------
' WRITE ROWS
'------------------------------------------------------------------------------
    'Write three interest rate rows: two with an error value in the Issue ID, one in the currency
        PutRow 1, 0, Array("IR-K1", "T", "Bond", "EUR", "Long", 3000000, "Fixed", 0.05, _
            DateSerial(2030, 1, 1), Empty, "OTHER", "BB+ to BB-", CVErr(ERR_NA))
        PutRow 1, 1, Array("IR-K2", "T", "Bond", "EUR", "Short", 1000000, "Fixed", 0.05, _
            DateSerial(2030, 1, 1), Empty, "OTHER", "BB+ to BB-", CVErr(ERR_VALUE))
        PutRow 1, 2, Array("IR-K3", "T", "Bond", CVErr(ERR_NA), "Long", 1000000, "Fixed", 0.05, _
            DateSerial(2030, 1, 1), Empty, "NONE")
    'Write three equity rows, each with an error value in one cell
        PutRow 2, 0, Array(CVErr(ERR_NA), "T", "IT", "Single name", "ENI", "EUR", "Short", 1000000)
        PutRow 2, 1, Array("EQ-K2", "T", "IT", "Single name", CVErr(ERR_REF), "EUR", "Long", 2000000)
        PutRow 2, 2, Array("EQ-K3", "T", CVErr(ERR_NA), "Index", "DAX", "EUR", "Long", 1000000)
    'Write two commodity rows and one foreign exchange row, each with an error value in one cell
        PutRow 3, 0, Array("CO-K1", "T", CVErr(ERR_NA), "Long", 100, "t", 250, "EUR")
        PutRow 3, 1, Array("CO-K2", "T", "Wheat", "Long", 100, "t", 250, CVErr(ERR_NA))
        PutRow 4, 0, Array("FX-K1", "FX", CVErr(ERR_NA), "Long", 1000000)

'------------------------------------------------------------------------------
' RUN CALCULATION
'------------------------------------------------------------------------------
    'Run the calculation; its result line is not compared
        ResultText = RunSSA_Silent()

'------------------------------------------------------------------------------
' ASSERT RESULTS
'------------------------------------------------------------------------------
    'Assert the issue key of each of the two Issue IDs typed as an error
        TEST_Harness.AssertEqualText "issue-key-of-an-issue-id-typed-as-an-error", _
            ERROR_TEXT, PlainText(RowCell(NM_IR_CALC_HEADER, mFirst(1), 16))
        TEST_Harness.AssertEqualText "issue-key-of-another-issue-id-typed-as-an-error", _
            ERROR_TEXT, PlainText(RowCell(NM_IR_CALC_HEADER, mFirst(1) + 1, 16))
    'Assert the net of that issue on both of its rows
        TEST_Harness.AssertEqualDouble "net-of-the-issue-of-error-ids", 2000000, _
            NumberOf(RowCell(NM_IR_CALC_HEADER, mFirst(1), 17))
        TEST_Harness.AssertEqualDouble "net-of-the-issue-on-its-second-row", 2000000, _
            NumberOf(RowCell(NM_IR_CALC_HEADER, mFirst(1) + 1, 17))
    'Assert the text and the status of the currency typed as an error
        TEST_Harness.AssertEqualText "currency-typed-as-an-error", _
            ERROR_TEXT, PlainText(RowCell(NM_IR_CALC_HEADER, mFirst(1) + 2, 2))
        TEST_Harness.AssertEqualText "status-of-a-currency-typed-as-an-error", _
            "INVALID: currency not in FX_Table", _
            PlainText(RowCell(NM_IR_CALC_HEADER, mFirst(1) + 2, 1))
    'Assert the text of the PositionID typed as an error
        TEST_Harness.AssertEqualText "position-id-typed-as-an-error", _
            ERROR_TEXT, PlainText(RowCell(NM_EQ_CALC_HEADER, mFirst(2), 0))
    'Assert that no output cell holds an error value
        TEST_Harness.AssertEqualDouble "output-cells-that-hold-an-error-value", 0, ErrorCells()

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
    'Restore normal error handling, then remove the appended rows on the success and the error path
CleanUp:
        On Error GoTo 0
    'Remove the rows only when the case appended them
        If RestoreDue Then
            Detail = RemoveAppended()
        End If
    'Assert that each table has the rows it had
        TEST_Harness.AssertTrue "appended-rows-removed", Len(Detail) = 0, Detail

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error of the case and continue at the cleanup
CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub

Private Sub CaseReadingAnErrorValue()

'
'------------------------------------------------------------------------------
'                         PROBE READING AN ERROR VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   A cell given an error value holds one: its Value is an error, which
'   PlainText reads as the text #ERROR, while the text #N/A written into a cell
'   that is formatted as text stays a text. HB-08
'
' WHY THIS EXISTS
'   An error value given by CVErr and the text #N/A are told apart through
'   Range.Value and PlainText; this case measures that on a scratch cell
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - appends one foreign exchange row
'     - writes CVErr(#N/A) into its amount cell and asserts that the value is
'       an error and that PlainText reads it as #ERROR
'     - writes the text #N/A into the first cell of the row and asserts that it
'       stays a text
'     - removes the appended row and asserts that the table has the rows it had
'
' ERROR POLICY
'   An error of the case is recorded with RecordCaseError of TEST_Harness and
'   the case resumes at its cleanup, which runs on the success and the error
'   path
'
'   RemoveAppended contains its own error and returns it as a text, so that the
'   removal is asserted on both paths
'
' DEPENDENCIES
'   AppendRows, PositionHeader, PlainText and RemoveAppended; TEST_Harness
'   (BeginCase, the assertions and RecordCaseError)
'
' NOTES
'   State ownership: appends rows under the position tables and removes them at
'   the end
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target      As Range    'Scratch cell: the amount of an appended row
    Dim RestoreDue  As Boolean  'True once rows are appended
    Dim Detail      As String   'Why the appended rows could not be removed

'------------------------------------------------------------------------------
' START CASE
'------------------------------------------------------------------------------
    'Record an error of the case instead of ending the run, and name the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "reading-an-error-value"
    'Remember that the cleanup has rows to remove
        RestoreDue = True

'------------------------------------------------------------------------------
' APPEND ROW
'------------------------------------------------------------------------------
    'Make room under the foreign exchange table
        AppendRows 4, 1

'------------------------------------------------------------------------------
' PROBE ERROR VALUE
'------------------------------------------------------------------------------
    'Give the amount cell an error value
        Set Target = PositionHeader(4).Offset(mFirst(4), 4)
        Target.Value = CVErr(ERR_NA)
    'Assert that the cell holds an error value, which PlainText reads as the error text
        TEST_Harness.AssertTrue "error-value-is-read-as-an-error", IsError(Target.Value), _
            "a cell given CVErr(#N/A) must hold an error value"
        TEST_Harness.AssertEqualText "error-value-reads-as-the-error-text", _
            ERROR_TEXT, PlainText(Target.Value)

'------------------------------------------------------------------------------
' PROBE TEXT
'------------------------------------------------------------------------------
    'Write the text #N/A into a cell formatted as text
        Set Target = PositionHeader(4).Offset(mFirst(4), 0)
        Target.Value = "#N/A"
    'Assert that the cell holds a text
        TEST_Harness.AssertTrue "text-in-a-text-cell-stays-a-text", _
            Not IsError(Target.Value) And VarType(Target.Value) = vbString, _
            "#N/A written into a text-formatted cell must stay a text"

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
    'Restore normal error handling, then remove the appended rows on the success and the error path
CleanUp:
        On Error GoTo 0
    'Remove the rows only when the case appended them
        If RestoreDue Then
            Detail = RemoveAppended()
        End If
    'Assert that each table has the rows it had
        TEST_Harness.AssertTrue "appended-rows-removed", Len(Detail) = 0, Detail

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error of the case and continue at the cleanup
CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub

Private Sub CaseTextWrittenIntoACell()

'
'------------------------------------------------------------------------------
'                        PROBE TEXT WRITTEN INTO A CELL
'------------------------------------------------------------------------------
' PURPOSE
'   A text that reads as a number, written as it is into a cell that is not
'   formatted as text, becomes a number; written with the cell formatted as
'   text, and the cell given its format back, it stays a text with its digits.
'   This is how the cases suite writes a text. HB-09
'
' WHY THIS EXISTS
'   The cases suite writes a text into a cell through the number format, so
'   that the text keeps its digits; this case measures that on the host
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - appends one foreign exchange row and takes its amount cell
'     - asserts that the cell is not formatted as text, writes 00123 as it is
'       and asserts that the value is a number
'     - clears the cell, formats it as text, writes 00123, gives the format
'       back and asserts that the value is a text with its digits
'     - removes the appended row and asserts that the table has the rows it had
'
' ERROR POLICY
'   An error of the case is recorded with RecordCaseError of TEST_Harness and
'   the case resumes at its cleanup, which runs on the success and the error
'   path. The cleanup gives the cell its format back when the error came while
'   it was formatted as text
'
'   RemoveAppended contains its own error and returns it as a text, so that the
'   removal is asserted on both paths
'
' DEPENDENCIES
'   AppendRows, PositionHeader and RemoveAppended; TEST_Harness (BeginCase, the
'   assertions and RecordCaseError)
'
' NOTES
'   State ownership: appends rows under the position tables and removes them at
'   the end; gives the amount cell its number format back
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target          As Range    'Scratch cell: the amount of an appended row
    Dim Saved           As String   'Number format of that cell
    Dim FormatChanged   As Boolean  'True while the cell is formatted as text
    Dim RestoreDue      As Boolean  'True once rows are appended
    Dim Detail          As String   'Why the appended rows could not be removed

'------------------------------------------------------------------------------
' START CASE
'------------------------------------------------------------------------------
    'Record an error of the case instead of ending the run, and name the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "text-written-into-a-cell"
    'Remember that the cleanup has rows to remove
        RestoreDue = True

'------------------------------------------------------------------------------
' APPEND ROW
'------------------------------------------------------------------------------
    'Make room under the foreign exchange table
        AppendRows 4, 1

'------------------------------------------------------------------------------
' READ NUMBER FORMAT
'------------------------------------------------------------------------------
    'Take the amount cell and keep its number format
        Set Target = PositionHeader(4).Offset(mFirst(4), 4)
        Saved = CStr(Target.NumberFormat)
    'Assert that the cell is not formatted as text
        TEST_Harness.AssertTrue "amount-cell-is-not-text-formatted", Saved <> "@", _
            "the scratch cell must not be formatted as text; it is " & Saved

'------------------------------------------------------------------------------
' PROBE DIRECT WRITE
'------------------------------------------------------------------------------
    'Write the text as it is and assert that it became a number
        Target.Value = "00123"
        TEST_Harness.AssertTrue "text-written-directly-becomes-a-number", _
            VarType(Target.Value) = vbDouble, _
            "00123 written as it is must become a number"
    'Empty the cell before the second write
        Target.ClearContents

'------------------------------------------------------------------------------
' PROBE WRITE THROUGH FORMAT
'------------------------------------------------------------------------------
    'Format the cell as text, write the text and give the format back
        FormatChanged = True
        Target.NumberFormat = "@"
        Target.Value = "00123"
        Target.NumberFormat = Saved
        FormatChanged = False
    'Assert that the text keeps its digits and is a text, and that the format is back
        TEST_Harness.AssertEqualText "text-written-through-the-format-keeps-its-digits", _
            "00123", CStr(Target.Value)
        TEST_Harness.AssertTrue "text-written-through-the-format-is-a-text", _
            VarType(Target.Value) = vbString, _
            "the value must be a text, not a number"
        TEST_Harness.AssertEqualText "number-format-is-back", Saved, CStr(Target.NumberFormat)

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
    'Restore normal error handling, give the cell its format back after an error, then remove the appended rows
CleanUp:
        On Error GoTo 0
        If FormatChanged Then Target.NumberFormat = Saved
    'Remove the rows only when the case appended them
        If RestoreDue Then
            Detail = RemoveAppended()
        End If
    'Assert that each table has the rows it had
        TEST_Harness.AssertTrue "appended-rows-removed", Len(Detail) = 0, Detail

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error of the case and continue at the cleanup
CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub

Private Sub CaseAreasAsTables()

'
'------------------------------------------------------------------------------
'                            PROBE AREAS AS TABLES
'------------------------------------------------------------------------------
' PURPOSE
'   The Value and the Formula of a range of more than one cell are tables with
'   bounds from 1, and those of one cell are not: the cases suite keeps and
'   restores the input areas this way
'
' WHY THIS EXISTS
'   The cases suite keeps and restores the input areas through Value and
'   Formula, and relies on the shape of what they return; this case measures
'   that shape on the host
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - appends two foreign exchange rows and takes two rows and two columns of
'       them, and one cell
'     - asserts that the Value and the Formula of the area are arrays, that the
'       Formula has two rows and starts at 1, and that those of the one cell
'       are not arrays
'     - removes the appended rows and asserts that the table has the rows it
'       had
'
' ERROR POLICY
'   An error of the case is recorded with RecordCaseError of TEST_Harness and
'   the case resumes at its cleanup, which runs on the success and the error
'   path
'
'   RemoveAppended contains its own error and returns it as a text, so that the
'   removal is asserted on both paths
'
' DEPENDENCIES
'   AppendRows, PositionHeader and RemoveAppended; TEST_Harness (BeginCase, the
'   assertions and RecordCaseError)
'
' NOTES
'   State ownership: appends rows under the position tables and removes them at
'   the end
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Area        As Range    'Two rows and two columns of an appended table
    Dim One         As Range    'One cell of it
    Dim RestoreDue  As Boolean  'True once rows are appended
    Dim Detail      As String   'Why the appended rows could not be removed

'------------------------------------------------------------------------------
' START CASE
'------------------------------------------------------------------------------
    'Record an error of the case instead of ending the run, and name the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "areas-as-tables"
    'Remember that the cleanup has rows to remove
        RestoreDue = True

'------------------------------------------------------------------------------
' APPEND ROWS
'------------------------------------------------------------------------------
    'Make room under the foreign exchange table
        AppendRows 4, 2

'------------------------------------------------------------------------------
' TAKE AREA
'------------------------------------------------------------------------------
    'Take two rows and two columns of the appended rows, and one cell of them
        Set Area = PositionHeader(4).Offset(mFirst(4), 0).Resize(2, 2)
        Set One = Area.Cells(1, 1)

'------------------------------------------------------------------------------
' ASSERT TABLES
'------------------------------------------------------------------------------
    'Assert that the Value and the Formula of the area are arrays
        TEST_Harness.AssertTrue "value-of-an-area-is-a-table", IsArray(Area.Value), _
            "Value of two rows and two columns must be an array"
        TEST_Harness.AssertTrue "formula-of-an-area-is-a-table", IsArray(Area.Formula), _
            "Formula of two rows and two columns must be an array"
    'Assert the rows and the lower bound of the Formula of the area
        TEST_Harness.AssertEqualDouble "formula-of-an-area-has-two-rows", 2, UBound(Area.Formula, 1)
        TEST_Harness.AssertEqualDouble "formula-of-an-area-starts-at-one", 1, LBound(Area.Formula, 1)
    'Assert that the Value and the Formula of one cell are not arrays
        TEST_Harness.AssertTrue "value-of-one-cell-is-not-a-table", Not IsArray(One.Value), _
            "Value of one cell must not be an array"
        TEST_Harness.AssertTrue "formula-of-one-cell-is-not-a-table", Not IsArray(One.Formula), _
            "Formula of one cell must not be an array"

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
    'Restore normal error handling, then remove the appended rows on the success and the error path
CleanUp:
        On Error GoTo 0
    'Remove the rows only when the case appended them
        If RestoreDue Then
            Detail = RemoveAppended()
        End If
    'Assert that each table has the rows it had
        TEST_Harness.AssertTrue "appended-rows-removed", Len(Detail) = 0, Detail

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error of the case and continue at the cleanup
CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub

Private Sub CaseSheetProtection()

'
'------------------------------------------------------------------------------
'                            PROBE SHEET PROTECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Every sheet of the template is protected, and only the input cells are
'   unlocked. A write by code to a locked cell is refused on a protected sheet,
'   and accepted once AllowWriting has applied UserInterfaceOnly to it, which
'   leaves the sheet protected. HB-15
'
' WHY THIS EXISTS
'   The workbook writes its output sheets by code and users type only in the
'   input cells; this case measures on the host that the protection allows that
'   and nothing else
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - takes the header cell of Results and keeps its text
'     - asserts that no sheet is unprotected, that AsOfDate and the first cell
'       under the header of IRPositions are unlocked and that the header of
'       Results is locked
'     - protects Results as the template does, without UserInterfaceOnly,
'       writes the header and asserts that the write raises the error 1004
'     - applies AllowWriting, writes again and asserts that the write succeeds
'       and that the sheet is still protected
'
' ERROR POLICY
'   An error of the case is recorded with RecordCaseError of TEST_Harness and
'   the case resumes at its cleanup, which runs on the success and the error
'   path; the error of a probed write is captured, not raised
'
' DEPENDENCIES
'   UnprotectedSheets, PositionHeader and AllowWriting; NM_RESULTS_HEADER;
'   TEST_Harness (BeginCase, the assertions and RecordCaseError)
'
' NOTES
'   State ownership: protects Results again for one write and applies
'   UserInterfaceOnly to it again; changes the label of its header for that
'   write and puts it back
'
' UPDATED
'   2026-11-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'Output sheet used for the writes
    Dim Target      As Range        'A locked label on it
    Dim KeptValue   As Variant      'Its text before the case
    Dim Written     As Boolean      'True while the label is not its own
    Dim Raised      As Long         'Err.Number right after a write

'------------------------------------------------------------------------------
' START CASE
'------------------------------------------------------------------------------
    'Record an error of the case instead of ending the run, and name the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "sheet-protection"

'------------------------------------------------------------------------------
' READ TARGET
'------------------------------------------------------------------------------
    'Take the locked header label of Results, its sheet and its text
        Set Target = ThisWorkbook.Names(NM_RESULTS_HEADER).RefersToRange
        Set Sheet = Target.Worksheet
        KeptValue = Target.Value

'------------------------------------------------------------------------------
' CHECK PROTECTION
'------------------------------------------------------------------------------
    'Assert that every sheet is protected
        TEST_Harness.AssertEqualDouble "sheets-that-are-not-protected", 0, UnprotectedSheets()
    'Assert that the input cells are unlocked and the label is locked
        TEST_Harness.AssertTrue "input-cell-of-params-is-unlocked", _
            Not ThisWorkbook.Names("AsOfDate").RefersToRange.Locked, _
            "AsOfDate must be unlocked: users type in it"
        TEST_Harness.AssertTrue "position-cell-is-unlocked", _
            Not PositionHeader(1).Offset(1, 0).Locked, _
            "the first cell of IRPositions under its header must be unlocked"
        TEST_Harness.AssertTrue "label-is-locked", Target.Locked, _
            "the header of Results must be locked"

'------------------------------------------------------------------------------
' PROBE REFUSED WRITE
'------------------------------------------------------------------------------
    'Protect the sheet as the template protects it, without UserInterfaceOnly
        Sheet.Protect Contents:=True
    'Remember that the label may no longer be its own
        Written = True
    'Write the label and capture the error number without ending the case
        On Error Resume Next
        Target.Value = "probe"
        Raised = Err.Number
        On Error GoTo CaseFailed
    'Assert that the write to a locked cell is refused
        TEST_Harness.AssertRaised "write-to-a-locked-cell-is-refused", ERR_PROTECTED_WRITE, Raised

'------------------------------------------------------------------------------
' PROBE ALLOWED WRITE
'------------------------------------------------------------------------------
    'Apply UserInterfaceOnly to the sheet
        AllowWriting Sheet
    'Write the label again and capture the error number
        On Error Resume Next
        Target.Value = "probe"
        Raised = Err.Number
        On Error GoTo CaseFailed
    'Assert that the write succeeds and that the sheet is still protected
        TEST_Harness.AssertRaised "write-after-allow-writing-succeeds", 0, Raised
        TEST_Harness.AssertTrue "sheet-is-still-protected", Sheet.ProtectContents, _
            "AllowWriting must leave the sheet protected"

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
    'Restore normal error handling, then put the label back and protect the sheet again on both paths
CleanUp:
        On Error GoTo 0
        If Written Then Target.Value = KeptValue
        AllowWriting Sheet

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error of the case and continue at the cleanup
CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

End Sub

Private Sub CaseStatusOnTheOtherSheets()

'
'------------------------------------------------------------------------------
'                         PROBE STATUS ON OTHER SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   After a run, E1 of each calculation sheet and of Checks shows the status
'   line of Results. HB-10
'
' WHY THIS EXISTS
'   The status line of a run is shown on the other output sheets as well as on
'   Results; this case measures that on the host after a calculation
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - runs the calculation and asserts the run outcome and that the status
'       line of Results starts with OK:
'     - asserts that E1 of IR_Zones, IR_Ladder, Buckets, IRCalc, EQCalc,
'       COCalc, FXCalc and Checks shows that status line
'
' ERROR POLICY
'   An error of the case is recorded with RecordCaseError of TEST_Harness and
'   the run continues; the case has no cleanup
'
' DEPENDENCIES
'   RunSSA_Silent, StatusLine and PlainText; TEST_Harness (BeginCase, the
'   assertions and RecordCaseError)
'
' NOTES
'   State ownership: rewrites the output sheets; the case that follows rewrites
'   them again
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultText  As String   'RunSSA_Silent result line

'------------------------------------------------------------------------------
' START CASE
'------------------------------------------------------------------------------
    'Record an error of the case instead of ending the run, and name the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "status-on-the-other-sheets"

'------------------------------------------------------------------------------
' RUN CALCULATION
'------------------------------------------------------------------------------
    'Run the calculation and keep its result line
        ResultText = RunSSA_Silent()

'------------------------------------------------------------------------------
' CHECK STATUS LINE
'------------------------------------------------------------------------------
    'Assert the outcome of the run and the start of the status line of Results
        TEST_Harness.AssertEqualText "run-outcome", "RESULT=OK", Left$(ResultText, 9)
        TEST_Harness.AssertEqualText "status-line-of-the-run", "OK: ", Left$(StatusLine(), 4)

'------------------------------------------------------------------------------
' CHECK OTHER SHEETS
'------------------------------------------------------------------------------
    'Assert that E1 of each other output sheet shows that status line
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

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseSaveStateAfterTheCheck()

'
'------------------------------------------------------------------------------
'                         PROBE SAVE STATE AFTER CHECK
'------------------------------------------------------------------------------
' PURPOSE
'   The status check that the workbook runs when it is opened and before it is
'   saved, CheckStatus_OnEvent, leaves a workbook with no unsaved change as it
'   was when the results are current, and leaves it with an unsaved change when
'   it writes OUTDATED: before the status line. HB-11. The events themselves,
'   Workbook_Open and Workbook_BeforeSave, are checked by hand: INSTALLATION.md
'
' WHY THIS EXISTS
'   The status check may write the status line, and a write marks the workbook
'   as changed; this case measures on the host when the check leaves the
'   workbook saved
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - keeps the save state, runs the calculation, marks the workbook saved
'       and runs CheckStatus_OnEvent; asserts that the workbook is still saved
'       and that the status line starts with OK:
'     - changes ScalingIR by 0.1, marks the workbook saved, runs the check and
'       asserts that the status line starts with OUTDATED: OK: and that the
'       workbook has an unsaved change
'     - puts ScalingIR back, runs the check and asserts that the status is
'       current again
'     - clears the output sheets, restores the save state and asserts that the
'       outputs are cleared
'
' ERROR POLICY
'   An error of the case is recorded with RecordCaseError of TEST_Harness and
'   the case resumes at its cleanup, which runs on the success and the error
'   path
'
' DEPENDENCIES
'   RunSSA_Silent, ClearOutputs_Silent, CheckStatus_OnEvent and StatusLine;
'   TEST_Harness (BeginCase, the assertions and RecordCaseError)
'
' NOTES
'   State ownership: changes ScalingIR and ThisWorkbook.Saved, and puts both
'   back. Rewrites the output sheets and ends with them cleared
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WasSaved    As Boolean  'ThisWorkbook.Saved at the start
    Dim Target      As Range    'The IR scaling factor
    Dim KeptValue   As Variant  'Its value before the case
    Dim Changed     As Boolean  'True while the factor is not its own
    Dim Cleared     As String   'ClearOutputs_Silent result line
    Dim ResultText  As String   'RunSSA_Silent result line, not compared

'------------------------------------------------------------------------------
' START CASE
'------------------------------------------------------------------------------
    'Record an error of the case instead of ending the run, and name the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "save-state-after-the-check"

'------------------------------------------------------------------------------
' RUN CALCULATION
'------------------------------------------------------------------------------
    'Keep the save state of the workbook
        WasSaved = ThisWorkbook.Saved
    'Run the calculation, so that the results are current
        ResultText = RunSSA_Silent()

'------------------------------------------------------------------------------
' CHECK CURRENT RESULTS
'------------------------------------------------------------------------------
    'Mark the workbook saved and run the status check
        ThisWorkbook.Saved = True
        CheckStatus_OnEvent
    'Assert that the check left the workbook saved and the status current
        TEST_Harness.AssertTrue "saved-after-a-check-of-current-results", ThisWorkbook.Saved, _
            "a check that writes nothing must leave the workbook saved"
        TEST_Harness.AssertEqualText "status-after-a-check-of-current-results", "OK: ", _
            Left$(StatusLine(), 4)

'------------------------------------------------------------------------------
' CHECK OUTDATED RESULTS
'------------------------------------------------------------------------------
    'Change the IR scaling factor so that the results are outdated
        Set Target = ThisWorkbook.Names("ScalingIR").RefersToRange
        KeptValue = Target.Value
        Changed = True
        Target.Value = KeptValue + 0.1
    'Mark the workbook saved and run the status check
        ThisWorkbook.Saved = True
        CheckStatus_OnEvent
    'Assert that the check marked the status outdated and left an unsaved change
        TEST_Harness.AssertEqualText "status-marked-outdated-by-the-check", "OUTDATED: OK: ", _
            Left$(StatusLine(), 14)
        TEST_Harness.AssertTrue "unsaved-after-the-status-was-written", Not ThisWorkbook.Saved, _
            "a check that writes OUTDATED: must leave an unsaved change"

'------------------------------------------------------------------------------
' PUT INPUT BACK
'------------------------------------------------------------------------------
    'Put the factor back and run the status check
        Target.Value = KeptValue
        Changed = False
        CheckStatus_OnEvent
    'Assert that the status is current again
        TEST_Harness.AssertEqualText "status-current-after-the-input-is-put-back", "OK: ", _
            Left$(StatusLine(), 4)

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
    'Restore normal error handling, put the factor back after an error, clear the outputs and restore the save state
CleanUp:
        On Error GoTo 0
        If Changed Then Target.Value = KeptValue
        Cleared = ClearOutputs_Silent()
        ThisWorkbook.Saved = WasSaved
    'Assert that the outputs are cleared
        TEST_Harness.AssertEqualText "outputs-cleared", "RESULT=CLEARED", Left$(Cleared, 14)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error of the case and continue at the cleanup
CaseFailed:
        TEST_Harness.RecordCaseError
        Resume CleanUp

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
'   protects: apply UserInterfaceOnly to each of them, as AllowWriting does for
'   the sheets that RunSSA writes
'
' WHY THIS EXISTS
'   The cases write cells of sheets that the template protects, and code may
'   write a protected sheet only after UserInterfaceOnly is applied to it
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
'   The procedure has no error handler: an error of AllowWriting propagates to
'   RunSuite, which passes it to EndRun
'
' DEPENDENCIES
'   AllowWriting; ThisWorkbook.Worksheets
'
' NOTES
'   State ownership: changes the protection mode of every sheet, not who may
'   change it
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
    'Apply UserInterfaceOnly to each sheet of the workbook
        For Each Sheet In ThisWorkbook.Worksheets
            AllowWriting Sheet
        Next Sheet

End Sub

Private Function PositionHeader( _
    ByVal Index As Long) _
    As Range
'
'------------------------------------------------------------------------------
'                            RETURN POSITION HEADER
'------------------------------------------------------------------------------
' PURPOSE
'   The header cell of one of the four position sheets
'
' WHY THIS EXISTS
'   The cases address each position table through its header cell, whose
'   workbook name differs for each risk class
'
' INPUTS
'   Index
'     Long; 1 interest rate, 2 equity, 3 commodity, 4 foreign exchange
'
' RETURNS
'   The cell that the workbook name of the table refers to; Nothing for an
'   Index outside 1 to 4
'
' BEHAVIOR
'   Selects the workbook name of the table by Index and returns the cell it
'   refers to
'
' ERROR POLICY
'   The procedure has no error handler: an Index outside 1 to 4 selects no case
'   and the result stays Nothing, and a missing workbook name raises the error
'   of Names to the caller
'
' DEPENDENCIES
'   ThisWorkbook.Names; NM_IR_POS_HEADER, NM_EQ_POS_HEADER, NM_CO_POS_HEADER
'   and NM_FX_POS_HEADER of CORE_Config
'
' NOTES
'   None
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' SELECT HEADER
'------------------------------------------------------------------------------
    'Return the header cell of the table that Index names
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
'------------------------------------------------------------------------------
'                            RETURN LAST ROW OFFSET
'------------------------------------------------------------------------------
' PURPOSE
'   The rows of a position table: from the row under the header to the last row
'   that holds a PositionID, as M_Main reads it
'
' WHY THIS EXISTS
'   A case must know how many rows a table holds before it appends rows under
'   it, and the table has no fixed length
'
' INPUTS
'   Index
'     Long; the position table, as PositionHeader
'
' RETURNS
'   The number of rows under the header of the table, down to the last
'   PositionID
'
' BEHAVIOR
'   Takes the header cell and its sheet, finds the last used row of the
'   PositionID column and subtracts the row of the header
'
' ERROR POLICY
'   The procedure has no error handler: an error of PositionHeader propagates
'
' DEPENDENCIES
'   PositionHeader; Range.End
'
' NOTES
'   None
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header  As Range        'Header cell
    Dim Sheet   As Worksheet    'Sheet of the table

'------------------------------------------------------------------------------
' FIND HEADER
'------------------------------------------------------------------------------
    'Take the header cell of the table and its sheet
        Set Header = PositionHeader(Index)
        Set Sheet = Header.Worksheet

'------------------------------------------------------------------------------
' FIND LAST ROW
'------------------------------------------------------------------------------
    'Count the rows from the header down to the last PositionID
        LastOffset = Sheet.Cells(Sheet.Rows.Count, Header.Column).End(xlUp).Row - Header.Row

End Function

Private Sub AppendRows( _
    ByVal Index As Long, _
    ByVal Count As Long)
'
'------------------------------------------------------------------------------
'                             RECORD APPENDED ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Make room for rows under a position table: remember how many rows it has
'   and where the first appended row is
'
' WHY THIS EXISTS
'   PutRow writes at the first appended row and RemoveAppended clears the
'   appended rows and compares the table with the rows it had; both read what
'   this procedure records
'
' INPUTS
'   Index
'     Long; the position table, as PositionHeader
'
'   Count
'     Long; rows that the case writes
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Records the rows the table has, the row under them as the first appended
'   row, and the number of rows to append
'
' ERROR POLICY
'   The procedure has no error handler: an error of LastOffset propagates
'
' DEPENDENCIES
'   LastOffset
'
' NOTES
'   State ownership: sets mOriginal, mFirst and mAppended of the table
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD ROWS
'------------------------------------------------------------------------------
    'Remember the rows the table has and where the first appended row is
        mOriginal(Index) = LastOffset(Index)
        mFirst(Index) = mOriginal(Index) + 1
    'Remember how many rows the case appends
        mAppended(Index) = Count

End Sub

Private Sub PutRow( _
    ByVal Index As Long, _
    ByVal Row As Long, _
    ByVal Values As Variant)
'
'------------------------------------------------------------------------------
'                              WRITE APPENDED ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Write one appended row of a position table
'
' WHY THIS EXISTS
'   The cases write rows of different lengths, with an error value or an empty
'   cell in some places; one procedure writes them cell by cell
'
' INPUTS
'   Index
'     Long; the position table, as PositionHeader
'
'   Row
'     Long; 0 for the first appended row, 1 for the next
'
'   Values
'     Variant array; the cells of the row from the PositionID. Empty leaves one
'     empty, and an error value from CVErr is written as it is
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Takes the first cell of the row and writes each element of Values, left to
'   right, except an Empty one
'
' ERROR POLICY
'   The procedure has no error handler: an error of a write propagates to the
'   case
'
' DEPENDENCIES
'   PositionHeader; mFirst
'
' NOTES
'   None
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target  As Range    'First cell of the row
    Dim c       As Long     'Cell of the row

'------------------------------------------------------------------------------
' FIND ROW
'------------------------------------------------------------------------------
    'Take the first cell of the row to write
        Set Target = PositionHeader(Index).Offset(mFirst(Index) + Row, 0)

'------------------------------------------------------------------------------
' WRITE CELLS
'------------------------------------------------------------------------------
    'Write each element of Values, and leave a cell empty for an Empty one
        For c = LBound(Values) To UBound(Values)
            If Not IsEmpty(Values(c)) Then
                Target.Offset(0, c - LBound(Values)).Value = Values(c)
            End If
        Next c

End Sub

Private Function RemoveAppended() As String

'
'------------------------------------------------------------------------------
'                             REMOVE APPENDED ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the rows a case appended, and check that each table has the rows it
'   had
'
' WHY THIS EXISTS
'   Every case that appends rows must leave the tables as it found them, on the
'   success and the error path, and must be able to assert it
'
' INPUTS
'   None
'
' RETURNS
'   An empty text when every table is as it was; otherwise what is left
'
' BEHAVIOR
'   For each of the four tables with appended rows, clears the appended rows
'   over the columns that the table reads, compares the rows of the table with
'   the rows it had and resets the count of appended rows
'
' ERROR POLICY
'   Contains its own error and returns it, so that the case asserts the result
'   on the success and the error path
'
' DEPENDENCIES
'   PositionHeader and LastOffset; IR_INPUT_WIDTH, EQ_INPUT_WIDTH,
'   CO_INPUT_WIDTH and FX_INPUT_WIDTH of CORE_Config; mOriginal, mFirst and
'   mAppended
'
' NOTES
'   State ownership: clears the appended rows and sets mAppended of each table
'   to 0
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i       As Long     'Position table
    Dim Left    As Long     'Rows that remain
    Dim Widths  As Variant  'Columns of each table

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return an error as text instead of raising it
        On Error GoTo Failed
    'Take the columns that each table reads
        Widths = Array(IR_INPUT_WIDTH, EQ_INPUT_WIDTH, CO_INPUT_WIDTH, FX_INPUT_WIDTH)

'------------------------------------------------------------------------------
' CLEAR APPENDED ROWS
'------------------------------------------------------------------------------
    'Clear the appended rows of each table and count the tables that differ from the rows they had
        For i = 1 To 4
            If mAppended(i) > 0 Then
                PositionHeader(i).Offset(mFirst(i), 0).Resize(mAppended(i), Widths(i - 1)).ClearContents
                If LastOffset(i) <> mOriginal(i) Then Left = Left + 1
                mAppended(i) = 0
            End If
        Next i

'------------------------------------------------------------------------------
' REPORT RESULT
'------------------------------------------------------------------------------
    'Report the tables that do not have the rows they had
        If Left <> 0 Then RemoveAppended = CStr(Left) & " table(s) do not have the rows they had"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return the error as text
        RemoveAppended = "error " & CStr(Err.Number) & " while removing rows: " & Err.Description

End Function

Private Function RowCell( _
    ByVal HeaderName As String, _
    ByVal Index As Long, _
    ByVal ColumnOffset As Long) _
    As Variant
'
'------------------------------------------------------------------------------
'                           RETURN CALCULATION CELL
'------------------------------------------------------------------------------
' PURPOSE
'   The value of one cell of a calculation sheet
'
' WHY THIS EXISTS
'   The assertions read cells of the calculation sheets by the workbook name of
'   the sheet's header and a position, not by an address
'
' INPUTS
'   HeaderName
'     String; workbook name of the sheet's header cell
'
'   Index
'     Long; row under the header, from 1: the position's row
'
'   ColumnOffset
'     Long; columns to the right of the PositionID
'
' RETURNS
'   The Value of the cell, as Excel gives it: it can be an error value
'
' BEHAVIOR
'   Offsets the header cell that the name refers to by Index rows and
'   ColumnOffset columns, and returns its Value
'
' ERROR POLICY
'   The procedure has no error handler: a missing workbook name raises the
'   error of Names to the caller
'
' DEPENDENCIES
'   ThisWorkbook.Names
'
' NOTES
'   None
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ CELL
'------------------------------------------------------------------------------
    'Return the value of the cell at the offset from the header
        RowCell = ThisWorkbook.Names(HeaderName).RefersToRange.Offset(Index, ColumnOffset).Value

End Function

Private Function ErrorCells() As Long

'
'------------------------------------------------------------------------------
'                              COUNT ERROR CELLS
'------------------------------------------------------------------------------
' PURPOSE
'   Count the cells that hold an error value in the output sheets
'
' WHY THIS EXISTS
'   A typed error value must not reach any output cell; the count lets one
'   assertion cover every output sheet
'
' INPUTS
'   None
'
' RETURNS
'   The number of such cells, in the used range of each sheet
'
' BEHAVIOR
'   Reads the used range of each output sheet (IR_Zones, IR_Ladder, Buckets,
'   IRCalc, EQCalc, COCalc, FXCalc, Results and Checks) and counts its error
'   values, whether the range is one cell or an array
'
' ERROR POLICY
'   The procedure has no error handler: a missing sheet raises the error of
'   Worksheets to the case
'
' DEPENDENCIES
'   ThisWorkbook.Worksheets
'
' NOTES
'   None
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheets  As Variant  'Names of the output sheets
    Dim Item    As Variant  'One name
    Dim Cells   As Variant  'Values of a used range
    Dim r       As Long     'Row
    Dim c       As Long     'Column
    Dim Count   As Long     'Cells that hold an error value

'------------------------------------------------------------------------------
' LIST SHEETS
'------------------------------------------------------------------------------
    'Name the output sheets to count
        Sheets = Array("IR_Zones", "IR_Ladder", "Buckets", "IRCalc", "EQCalc", "COCalc", _
                       "FXCalc", "Results", "Checks")

'------------------------------------------------------------------------------
' COUNT ERRORS
'------------------------------------------------------------------------------
    'Count the error values of the used range of each sheet
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

'------------------------------------------------------------------------------
' RETURN COUNT
'------------------------------------------------------------------------------
    'Return the number of error values
        ErrorCells = Count

End Function

Private Function UnprotectedSheets() As Long

'
'------------------------------------------------------------------------------
'                           COUNT UNPROTECTED SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Count the sheets of the workbook that are not protected
'
' WHY THIS EXISTS
'   The template protects every sheet; the count lets one assertion cover all
'   of them
'
' INPUTS
'   None
'
' RETURNS
'   The number of sheets whose contents are not protected
'
' BEHAVIOR
'   Counts the worksheets of ThisWorkbook whose ProtectContents is False
'
' ERROR POLICY
'   None; the procedure has no error handler
'
' DEPENDENCIES
'   ThisWorkbook.Worksheets
'
' NOTES
'   None
'
' UPDATED
'   2026-11-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet   As Worksheet    'One sheet
    Dim Count   As Long         'Sheets that are not protected

'------------------------------------------------------------------------------
' COUNT SHEETS
'------------------------------------------------------------------------------
    'Count the sheets whose contents are not protected
        For Each Sheet In ThisWorkbook.Worksheets
            If Not Sheet.ProtectContents Then Count = Count + 1
        Next Sheet

'------------------------------------------------------------------------------
' RETURN COUNT
'------------------------------------------------------------------------------
    'Return the number of unprotected sheets
        UnprotectedSheets = Count

End Function

Private Function StatusLine() As String

'
'------------------------------------------------------------------------------
'                              RETURN STATUS LINE
'------------------------------------------------------------------------------
' PURPOSE
'   Return the status line of Results
'
' WHY THIS EXISTS
'   Several assertions compare the status line; this reads it from the cell
'   that NM_RESULTS_STATUS names
'
' INPUTS
'   None
'
' RETURNS
'   The text of the Results status cell; empty for an empty cell
'
' BEHAVIOR
'   Returns NamedText of NM_RESULTS_STATUS
'
' ERROR POLICY
'   The procedure has no error handler: an error of NamedText propagates
'
' DEPENDENCIES
'   NamedText; NM_RESULTS_STATUS of CORE_Config
'
' NOTES
'   None
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ STATUS
'------------------------------------------------------------------------------
    'Return the text of the Results status cell
        StatusLine = NamedText(NM_RESULTS_STATUS)

End Function

Private Function NamedText( _
    ByVal CellName As String) _
    As String
'
'------------------------------------------------------------------------------
'                              RETURN NAMED TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Return the text of the cell a workbook name refers to; empty for an empty
'   cell
'
' WHY THIS EXISTS
'   A cell read through a workbook name can hold a number, an error value or
'   nothing; the callers compare a text
'
' INPUTS
'   CellName
'     String; a workbook name that refers to one cell
'
' RETURNS
'   The text of the cell as PlainText reads its value
'
' BEHAVIOR
'   Reads the Value of the cell that the name refers to and converts it with
'   PlainText
'
' ERROR POLICY
'   The procedure has no error handler: a missing workbook name raises the
'   error of Names to the caller
'
' DEPENDENCIES
'   PlainText of CORE_Util; ThisWorkbook.Names
'
' NOTES
'   None
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ TEXT
'------------------------------------------------------------------------------
    'Return the text of the cell that the name refers to
        NamedText = PlainText(ThisWorkbook.Names(CellName).RefersToRange.Value)

End Function

Private Function NumberOf( _
    ByVal Value As Variant) _
    As Double
'
'------------------------------------------------------------------------------
'                              RETURN CELL NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   The number a cell holds, for an assertion that compares numbers
'
' WHY THIS EXISTS
'   An empty cell, a text or an error value must never be compared as 0, so an
'   assertion on a number reads it through this check
'
' INPUTS
'   Value
'     Variant; the value read from a cell
'
' RETURNS
'   The value as a Double
'
' BEHAVIOR
'   Checks that the value is a number and converts it with CDbl
'
' ERROR POLICY
'   Raises error 13 when Value is not a number, so that an empty cell, a text
'   or an error value is never compared as 0; the case records it
'
' DEPENDENCIES
'   IsNumberValue of CORE_Util
'
' NOTES
'   None
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE VALUE
'------------------------------------------------------------------------------
    'Reject a value that is not a number
        If Not IsNumberValue(Value) Then
            Err.Raise 13, "TEST_Probes.NumberOf", "The result is not a number."
        End If

'------------------------------------------------------------------------------
' CONVERT VALUE
'------------------------------------------------------------------------------
    'Convert the number to a Double
        NumberOf = CDbl(Value)

End Function
