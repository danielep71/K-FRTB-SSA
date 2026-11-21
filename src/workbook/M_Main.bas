Attribute VB_Name = "M_Main"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Host layer of K-FRTB-SSA: the command macros, which read the input sheets
'   and Params into the engine's input vector, call CORE_Engine.Calculate and
'   write every output table back, with Excel application state captured and
'   restored around each operation. A run that fails leaves no result, and a
'   check tells whether an input changed after the last run
'
' WHY THIS EXISTS
'   Keeps the Excel object model out of the calculation: CORE_Engine takes and
'   returns arrays, and this module reads the sheets and workbook names into
'   those arrays and writes the output arrays back to the sheets. It also gives
'   the workbook's event procedures and the test modules one place to call
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   ExecuteOperation runs every operation: it captures Calculation,
'   ScreenUpdating and EnableEvents, sets manual calculation, no screen
'   updating and no events, lets code write the protected output sheets, runs
'   the operation and restores the three settings. The operations are:
'     - OP_RUN (RunSSA, RunSSA_Silent): RunCalculation reads the inputs, calls
'       Calculate, clears the outputs and writes them
'     - OP_CLEAR (ClearOutputs, ClearOutputs_Silent): ClearAllOutputs
'     - OP_CHECK (CheckStatus, CheckStatus_Silent, CheckStatus_OnEvent):
'       CheckResults compares the fingerprint of the inputs with the one the
'       last run stored
'
'   The procedures are grouped in six families: command macros, operation,
'   reading, writing, clearing and helpers
'
' ERROR POLICY
'   The command macros report a failure in a message box (silent variants:
'   re-raise it; CheckStatus_OnEvent: contain it). A run that raises an error
'   clears every output area and writes a FAILED status line before the error
'   is reported. The primary error is preserved through cleanup; a cleanup-only
'   failure raises ERR_CLEANUP_FAILED
'
' DEPENDENCIES
'   CORE_Config, CORE_Engine and CORE_Util. Locates every table through the
'   workbook names listed in CORE_Config (TABLE LOCATION). Excel object model:
'   Application, ThisWorkbook, Worksheets, Names, Range and MsgBox
'
' NOTES
'   Public surface: RunSSA, RunSSA_Silent, ClearOutputs, ClearOutputs_Silent,
'   CheckStatus, CheckStatus_Silent (command macros), CheckStatus_OnEvent (for
'   the workbook's event procedures) and AllowWriting (for the test modules);
'   not part of the supported worksheet-function API
'
'   State ownership: reads the input sheets and Params; clears and writes the
'   output sheets (IRCalc, EQCalc, COCalc, FXCalc, IR_Ladder, IR_Zones,
'   Buckets, Results, Checks) and the calculated columns and the result block
'   of Eligibility. Changes Calculation, ScreenUpdating and EnableEvents for
'   the duration of an operation and restores them. Applies UserInterfaceOnly
'   to the protection of each output sheet and of Eligibility, without a
'   password, so that code may write a sheet that users cannot change; the
'   setting lasts until the workbook is closed. mSilent holds the silent flag
'   of the running operation, and mStatusWritten tells whether the last check
'   changed the status line
'
'   Option Explicit requires explicit declarations
'
'   Compatibility: Excel for Windows, VBA7; no references beyond the defaults
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const OP_RUN                As Long = 1                                 'Operation: calculate
    Private Const OP_CLEAR              As Long = 2                                 'Operation: clear outputs
    Private Const OP_CHECK              As Long = 3                                 'Operation: check the status line

    Private Const CLEAR_PARTS           As Long = 20                                'Parts ClearPart clears
    Private Const CHECKS_COUNT_COLUMN   As Long = 4                                 'Count, from the rule column (F)
    Private Const CHECKS_RESULT_COLUMN  As Long = 5                                 'Result, from the rule column (G)

    Private Const NOT_RUN               As String = "NOT RUN"                       'Result of a rule not evaluated
    Private Const STATUS_EMPTY          As String = "Not calculated: run RunSSA."   'No result is shown
    Private Const STATUS_FAILED         As String = "FAILED: "                      'Start of the line of a failed run
    Private Const STATUS_OUTDATED       As String = "OUTDATED: "                    'Prefix: inputs changed since run

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mSilent         As Boolean  'True while a _Silent macro runs
    Private mStatusWritten  As Boolean  'True when the last check changed the status line


'
'------------------------------------------------------------------------------
'
'                                COMMAND MACROS
'
'------------------------------------------------------------------------------
'

Public Sub RunSSA()

'
'------------------------------------------------------------------------------
'                     RUN SIMPLIFIED STANDARDISED APPROACH
'------------------------------------------------------------------------------
' PURPOSE
'   Calculates the MAR40 simplified standardised approach and writes every
'   output sheet, then shows Results (or Checks when the run stopped) and a
'   summary message
'
' WHY THIS EXISTS
'   Gives the workbook one macro that runs the whole calculation and tells the
'   user the outcome in a message box
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Allows message boxes and sheet activation for the run, runs
'   ExecuteOperation(OP_RUN) and shows its result line with one item per line.
'   When the run raises an error, it shows the description and the source of
'   the error
'
' ERROR POLICY
'   Reports a failure in a message box and does not re-raise it
'
' DEPENDENCIES
'   ExecuteOperation, mSilent, PRODUCT_NAME (CORE_Config); MsgBox
'
' NOTES
'   Usage: Alt+F8 > RunSSA, or a button assigned to it
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Summary As String   'Result line of the run

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the failure message
        On Error GoTo Failed
    'Allow message boxes and sheet activation for this run
        mSilent = False

'------------------------------------------------------------------------------
' RUN CALCULATION
'------------------------------------------------------------------------------
    'Run the calculation and keep its result line
        Summary = ExecuteOperation(OP_RUN)

'------------------------------------------------------------------------------
' SHOW SUMMARY
'------------------------------------------------------------------------------
    'Show the result line, one item per line
        MsgBox Replace(Summary, "; ", vbLf), vbInformation, PRODUCT_NAME

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Report the failure with its description and source
        MsgBox "The run failed." & vbLf & Err.Description & vbLf & "(" & Err.Source & ")", _
               vbCritical, PRODUCT_NAME

End Sub

Public Function RunSSA_Silent() As String

'
'------------------------------------------------------------------------------
'                                RUN UNATTENDED
'------------------------------------------------------------------------------
' PURPOSE
'   Runs RunSSA unattended: no message box, no sheet activation; failures are
'   raised to the caller
'
' WHY THIS EXISTS
'   Lets a caller with no user at the keyboard run the calculation and read its
'   result line without a dialog
'
' INPUTS
'   None
'
' RETURNS
'   "RESULT=OK|INCOMPLETE|STOPPED; errors=n; warnings=n; positions_used=n;
'   requirement=x; rwa=x; cleanup=PASS", numbers with a decimal point
'
' BEHAVIOR
'   Sets mSilent, runs ExecuteOperation(OP_RUN), clears mSilent and returns the
'   result line. When the run raises an error, it saves the error, clears
'   mSilent and raises the error again
'
' ERROR POLICY
'   Re-raises the operation's error after mSilent is cleared
'
' DEPENDENCIES
'   ExecuteOperation, mSilent
'
' NOTES
'   State ownership: sets mSilent for the run and clears it afterwards, also on
'   error
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber       As Long     'Primary error number; 0 for none
    Dim ErrSource       As String   'Primary error source
    Dim ErrDescription  As String   'Primary error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the handler
        On Error GoTo Failed
    'Suppress message boxes and sheet activation for this run
        mSilent = True

'------------------------------------------------------------------------------
' RUN CALCULATION
'------------------------------------------------------------------------------
    'Run the calculation, then clear the silent flag
        RunSSA_Silent = ExecuteOperation(OP_RUN)
        mSilent = False

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Save the primary error before any other call
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Clear the silent flag before the error leaves
        mSilent = False
    'Raise the saved error again
        Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

Public Sub ClearOutputs()

'
'------------------------------------------------------------------------------
'                             CLEAR OUTPUT SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Clears every output written by RunSSA; inputs and Params are untouched
'
' WHY THIS EXISTS
'   Gives the workbook one macro that removes the results of a run
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Allows message boxes for the operation and runs ExecuteOperation(OP_CLEAR).
'   When it raises an error, it shows the description of the error
'
' ERROR POLICY
'   Reports a failure in a message box and does not re-raise it
'
' DEPENDENCIES
'   ExecuteOperation, mSilent, PRODUCT_NAME (CORE_Config); MsgBox
'
' NOTES
'   Usage: Alt+F8 > ClearOutputs
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the failure message
        On Error GoTo Failed
    'Allow message boxes for this operation
        mSilent = False

'------------------------------------------------------------------------------
' CLEAR OUTPUTS
'------------------------------------------------------------------------------
    'Clear every output area
        ExecuteOperation OP_CLEAR

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Report the failure with its description
        MsgBox "Clearing failed." & vbLf & Err.Description, vbCritical, PRODUCT_NAME

End Sub

Public Function ClearOutputs_Silent() As String

'
'------------------------------------------------------------------------------
'                        CLEAR OUTPUT SHEETS UNATTENDED
'------------------------------------------------------------------------------
' PURPOSE
'   Runs ClearOutputs unattended
'
' WHY THIS EXISTS
'   Lets a caller with no user at the keyboard clear the outputs without a
'   dialog
'
' INPUTS
'   None
'
' RETURNS
'   "RESULT=CLEARED; cleanup=PASS"
'
' BEHAVIOR
'   Sets mSilent, runs ExecuteOperation(OP_CLEAR), clears mSilent and returns
'   the result line. When the operation raises an error, it saves the error,
'   clears mSilent and raises the error again
'
' ERROR POLICY
'   Re-raises the operation's error after mSilent is cleared
'
' DEPENDENCIES
'   ExecuteOperation, mSilent
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber       As Long     'Primary error number; 0 for none
    Dim ErrSource       As String   'Primary error source
    Dim ErrDescription  As String   'Primary error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the handler
        On Error GoTo Failed
    'Suppress message boxes for this operation
        mSilent = True

'------------------------------------------------------------------------------
' CLEAR OUTPUTS
'------------------------------------------------------------------------------
    'Clear every output area, then clear the silent flag
        ClearOutputs_Silent = ExecuteOperation(OP_CLEAR)
        mSilent = False

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Save the primary error before any other call
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Clear the silent flag before the error leaves
        mSilent = False
    'Raise the saved error again
        Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

Public Sub CheckStatus()

'
'------------------------------------------------------------------------------
'                             CHECK RESULT STATUS
'------------------------------------------------------------------------------
' PURPOSE
'   Tells whether the results shown are those of the inputs on the sheets,
'   marks them OUTDATED on the status line when an input changed after the run,
'   and shows the outcome in a message
'
' WHY THIS EXISTS
'   Gives the user one macro that finds out whether the results on the sheets
'   are still current
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Allows message boxes for the operation, runs ExecuteOperation(OP_CHECK) and
'   shows its result line followed by the status line of Results. When it
'   raises an error, it shows the description and the source of the error
'
' ERROR POLICY
'   Reports a failure in a message box and does not re-raise it
'
' DEPENDENCIES
'   ExecuteOperation, PlainText (CORE_Util), NamedRange, mSilent,
'   NM_RESULTS_STATUS and PRODUCT_NAME (CORE_Config); MsgBox
'
' NOTES
'   Usage: Alt+F8 > CheckStatus
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Summary As String   'Result line of the check

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the failure message
        On Error GoTo Failed
    'Allow message boxes for this operation
        mSilent = False

'------------------------------------------------------------------------------
' RUN CHECK
'------------------------------------------------------------------------------
    'Run the check and keep its result line
        Summary = ExecuteOperation(OP_CHECK)

'------------------------------------------------------------------------------
' SHOW OUTCOME
'------------------------------------------------------------------------------
    'Show the result line and the status line of Results
        MsgBox Replace(Summary, "; ", vbLf) & vbLf & _
               PlainText(NamedRange(NM_RESULTS_STATUS).Value), vbInformation, PRODUCT_NAME

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Report the failure with its description and source
        MsgBox "The check failed." & vbLf & Err.Description & vbLf & "(" & Err.Source & ")", _
               vbCritical, PRODUCT_NAME

End Sub

Public Function CheckStatus_Silent() As String

'
'------------------------------------------------------------------------------
'                        CHECK RESULT STATUS UNATTENDED
'------------------------------------------------------------------------------
' PURPOSE
'   Runs CheckStatus unattended
'
' WHY THIS EXISTS
'   Lets a caller with no user at the keyboard check the status line without a
'   dialog
'
' INPUTS
'   None
'
' RETURNS
'   "RESULT=CURRENT|OUTDATED|FAILED|NOT_CALCULATED; cleanup=PASS"
'
' BEHAVIOR
'   Sets mSilent, runs ExecuteOperation(OP_CHECK), clears mSilent and returns
'   the result line. When the operation raises an error, it saves the error,
'   clears mSilent and raises the error again
'
' ERROR POLICY
'   Re-raises the operation's error after mSilent is cleared
'
' DEPENDENCIES
'   ExecuteOperation, mSilent
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber       As Long     'Primary error number; 0 for none
    Dim ErrSource       As String   'Primary error source
    Dim ErrDescription  As String   'Primary error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the handler
        On Error GoTo Failed
    'Suppress message boxes for this operation
        mSilent = True

'------------------------------------------------------------------------------
' RUN CHECK
'------------------------------------------------------------------------------
    'Run the check, then clear the silent flag
        CheckStatus_Silent = ExecuteOperation(OP_CHECK)
        mSilent = False

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Save the primary error before any other call
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Clear the silent flag before the error leaves
        mSilent = False
    'Raise the saved error again
        Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

Public Sub CheckStatus_OnEvent()

'
'------------------------------------------------------------------------------
'                           CHECK STATUS FROM EVENT
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the check from an event procedure of the workbook: when it is opened
'   and before it is saved
'
' WHY THIS EXISTS
'   An event procedure has no caller to report to, and a message box would stop
'   an unattended run, so this entry point contains every error
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Remembers whether the workbook had no unsaved change, runs
'   ExecuteOperation(OP_CHECK) silently and, when the status line was not
'   written, sets ThisWorkbook.Saved to True again if it was True before. When
'   the operation raises an error, it clears mSilent and calls MarkUnchecked
'
' ERROR POLICY
'   Contains every error: an event procedure has no caller to report to, and a
'   message box would stop an unattended run. When the check cannot be made,
'   the results are marked OUTDATED; CheckStatus reports why
'
' DEPENDENCIES
'   ExecuteOperation, MarkUnchecked, mSilent, mStatusWritten;
'   ThisWorkbook.Saved
'
' NOTES
'   State ownership: writes the status line only when it must change, and then
'   leaves the workbook marked as changed; otherwise the workbook stays saved
'   if it was
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WasSaved    As Boolean  'Workbook had no unsaved change

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the handler
        On Error GoTo Failed
    'Remember whether the workbook had no unsaved change
        WasSaved = ThisWorkbook.Saved
    'Suppress message boxes and clear the status-written flag
        mSilent = True
        mStatusWritten = False

'------------------------------------------------------------------------------
' RUN CHECK
'------------------------------------------------------------------------------
    'Run the check, then clear the silent flag
        ExecuteOperation OP_CHECK
        mSilent = False

'------------------------------------------------------------------------------
' RESTORE SAVED STATE
'------------------------------------------------------------------------------
    'Leave the workbook saved when the check wrote nothing
        If WasSaved And Not mStatusWritten Then ThisWorkbook.Saved = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Clear the silent flag and mark the results OUTDATED
        mSilent = False
        MarkUnchecked

End Sub

'
'------------------------------------------------------------------------------
'
'                                  OPERATION
'
'------------------------------------------------------------------------------
'

Private Function ExecuteOperation( _
    ByVal Operation As Long) _
    As String
'
'------------------------------------------------------------------------------
'                              RUN ONE OPERATION
'------------------------------------------------------------------------------
' PURPOSE
'   Runs one operation with Excel state captured, changed and restored
'
' WHY THIS EXISTS
'   Every command macro needs the same capture, change and restoration of the
'   application settings, with the primary error kept through cleanup, so it is
'   written once
'
' INPUTS
'   Operation
'     OP_RUN, OP_CLEAR or OP_CHECK (Long)
'
' RETURNS
'   The operation's result line, with "; cleanup=PASS" appended
'
' BEHAVIOR
'   Captures Calculation, ScreenUpdating and EnableEvents, sets manual
'   calculation, no screen updating and no events, and lets code write the
'   output sheets. It then runs RunCalculation, ClearAllOutputs (result
'   "RESULT=CLEARED") or CheckResults according to Operation, restores all
'   three settings on success and on error, and recalculates when calculation
'   was automatic
'
' ERROR POLICY
'   Preserves the primary error through cleanup and re-raises it; raises
'   ERR_CLEANUP_FAILED when only the restoration failed. A run that raised an
'   error is first marked as failed on the sheets
'
' DEPENDENCIES
'   AllowWritingToOutputs, RunCalculation, ClearAllOutputs, CheckResults,
'   RestoreCalculation, RestoreFlag, AppendDetail, MarkFailed;
'   ERR_CLEANUP_FAILED (CORE_Config)
'
' NOTES
'   State ownership: captures Calculation, ScreenUpdating and EnableEvents;
'   sets manual calculation, no screen updating and no events; restores all
'   three on success and on error, then recalculates when calculation was
'   automatic
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedCalculation    As Long     'Calculation mode before the operation
    Dim SavedScreen         As Boolean  'ScreenUpdating before the operation
    Dim SavedEvents         As Boolean  'EnableEvents before the operation
    Dim StateCaptured       As Boolean  'True once all three were read
    Dim CleanupDetails      As String   'Restoration failures, "" for none
    Dim ResultLine          As String   'Result of the operation
    Dim ErrNumber           As Long     'Primary error number; 0 for none
    Dim ErrSource           As String   'Primary error source
    Dim ErrDescription      As String   'Primary error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the handler
        On Error GoTo Failed

'------------------------------------------------------------------------------
' CAPTURE APPLICATION STATE
'------------------------------------------------------------------------------
    'Read the three settings, then record that all were read
        SavedCalculation = Application.Calculation
        SavedScreen = Application.ScreenUpdating
        SavedEvents = Application.EnableEvents
        StateCaptured = True

'------------------------------------------------------------------------------
' CHANGE APPLICATION STATE
'------------------------------------------------------------------------------
    'Switch off screen updating and events, and calculate manually
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    'Let code write the protected output sheets
        AllowWritingToOutputs

'------------------------------------------------------------------------------
' RUN OPERATION
'------------------------------------------------------------------------------
    'Run the operation that was asked for
        Select Case Operation
            Case OP_RUN
                ResultLine = RunCalculation()
            Case OP_CLEAR
                ClearAllOutputs
                ResultLine = "RESULT=CLEARED"
            Case OP_CHECK
                ResultLine = CheckResults()
        End Select

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
CleanUp:
    'Disable the handler before the restoration and the raise
        On Error GoTo 0
    'Restore the captured settings and collect every restoration failure
        If StateCaptured Then
            CleanupDetails = AppendDetail(RestoreCalculation(SavedCalculation), _
                             AppendDetail(RestoreFlag("ScreenUpdating", SavedScreen), _
                                          RestoreFlag("EnableEvents", SavedEvents)))
        End If
    'Raise the primary error again
        If ErrNumber <> 0 Then Err.Raise ErrNumber, ErrSource, ErrDescription
    'Raise a cleanup failure when only the restoration failed
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "M_Main.ExecuteOperation", CleanupDetails
        End If
    'Return the result line
        ExecuteOperation = ResultLine & "; cleanup=PASS"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Save the primary error before any other call
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Mark a failed run on the sheets and add any marking failure to the description
        If Operation = OP_RUN Then ErrDescription = AppendDetail(ErrDescription, MarkFailed(ErrDescription))
    'Continue at the cleanup
        Resume CleanUp

End Function

Private Sub AllowWritingToOutputs()

'
'------------------------------------------------------------------------------
'                           ALLOW WRITING TO OUTPUTS
'------------------------------------------------------------------------------
' PURPOSE
'   Lets the operation write every output sheet, and the calculated columns of
'   Eligibility, which the template protects: applies UserInterfaceOnly to each
'   of those sheets
'
' WHY THIS EXISTS
'   The output sheets and Eligibility are protected, and code can write a
'   protected sheet only after UserInterfaceOnly is applied to it
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Lists the output sheets (Results, IR_Zones, IR_Ladder, Buckets, IRCalc,
'   EQCalc, COCalc, FXCalc and Checks) and Eligibility, by their SH_
'   constants, and calls AllowWriting for each of them
'
' ERROR POLICY
'   Raises the error of AllowWriting for the first sheet it cannot reach
'
' DEPENDENCIES
'   AllowWriting; the SH_ sheet-name constants of CORE_Config;
'   ThisWorkbook.Worksheets
'
' NOTES
'   State ownership: changes the protection mode of the output sheets and of
'   Eligibility, not who may change them: users still cannot
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SheetNames  As Variant  'Names of the sheets code writes
    Dim Item        As Variant  'One name

'------------------------------------------------------------------------------
' LIST OUTPUT SHEETS
'------------------------------------------------------------------------------
    'Name the output sheets and Eligibility, whose calculated columns code writes
        SheetNames = Array(SH_RESULTS, SH_IR_ZONES, SH_IR_LADDER, SH_BUCKETS, SH_IR_CALC, _
                           SH_EQ_CALC, SH_CO_CALC, SH_FX_CALC, SH_CHECKS, SH_ELIGIBILITY)

'------------------------------------------------------------------------------
' ALLOW WRITING
'------------------------------------------------------------------------------
    'Apply the protection mode to each sheet
        For Each Item In SheetNames
            AllowWriting ThisWorkbook.Worksheets(CStr(Item))
        Next Item

End Sub

Public Sub AllowWriting( _
    ByVal Sheet As Worksheet)
'
'------------------------------------------------------------------------------
'                          ALLOW CODE TO WRITE SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Lets code write a sheet that is protected without a password: protects it
'   again with UserInterfaceOnly and the options of the template's protection
'   (docs/methodology/workbook-protection.json): users may select cells and
'   size columns and rows, and nothing else. Users still cannot change a locked
'   cell. Nothing changes for a sheet that is not protected. The setting is not
'   saved with the workbook, so every operation applies it again
'
' WHY THIS EXISTS
'   A protected sheet cannot be written by code until UserInterfaceOnly is set,
'   and the setting is lost when the workbook is closed
'
' INPUTS
'   Sheet
'     The sheet code is about to write (Worksheet)
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Leaves a sheet that is not protected as it is. Otherwise it calls Protect
'   on the sheet, bound late, with UserInterfaceOnly and the options of the
'   template. When Excel refuses, it unprotects the sheet with an empty
'   password
'
' ERROR POLICY
'   Raises ERR_LAYOUT, with the sheet and the reason, when Excel refuses: a
'   sheet that someone protected with a password cannot be protected again, or
'   unprotected, without it. The attempt to unprotect passes an empty password,
'   so that Excel raises the error and never asks for the password: a run that
'   nobody watches, such as the check when the workbook opens or is saved, is
'   not stopped by a dialog. A host that does not know the options of Protect,
'   such as LibreOffice's VBA mode, which is not a supported host, gets the
'   sheet unprotected
'
' DEPENDENCIES
'   Worksheet.Protect, Worksheet.Unprotect and ProtectContents; ERR_LAYOUT
'   (CORE_Config)
'
' NOTES
'   State ownership: changes the protection mode of Sheet, and sets its options
'   to those of the template. It marks the workbook as changed; the caller that
'   must leave it saved restores that
'
' UPDATED
'   2026-11-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Late            As Object   'Sheet, so that the host binds the options when it runs
    Dim ErrNumber       As Long     'Primary error number
    Dim ErrDescription  As String   'Primary error description

'------------------------------------------------------------------------------
' CHECK PROTECTION
'------------------------------------------------------------------------------
    'Leave a sheet that is not protected as it is
        If Not Sheet.ProtectContents Then Exit Sub

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Bind the sheet late, so that the host binds the options when it runs
        Set Late = Sheet
    'Route a refusal to the next attempt
        On Error GoTo Narrow

'------------------------------------------------------------------------------
' PROTECT SHEET
'------------------------------------------------------------------------------
    'Protect it again with UserInterfaceOnly and the options of the template
        Late.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, _
                     UserInterfaceOnly:=True, AllowFormattingColumns:=True, _
                     AllowFormattingRows:=True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' REFUSED PROTECTION
'------------------------------------------------------------------------------
Narrow:
    'Leave the handler before the next call, so that an error of that call
    'reaches Failed and not the caller
        Resume Unprotecting

'------------------------------------------------------------------------------
' UNPROTECT SHEET
'------------------------------------------------------------------------------
Unprotecting:
    'A host that does not know those options, as LibreOffice's VBA mode does not
    'and which clears a protected sheet only after it is unprotected, gets the
    'sheet unprotected. Excel gets here only when it refuses the call above,
    'and refuses to unprotect a sheet that has a password too: the empty
    'password makes it raise an error instead of asking for one
        On Error GoTo Failed
        Sheet.Unprotect ""
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Save the primary error, then disable the handler
        ErrNumber = Err.Number
        ErrDescription = Err.Description
        On Error GoTo 0
    'Raise ERR_LAYOUT with the sheet, the reason and the error number
        Err.Raise ERR_LAYOUT, "M_Main.AllowWriting", "Sheet " & Sheet.Name & _
                  " is protected and cannot be written by code: " & ErrDescription & _
                  " (error " & CStr(ErrNumber) & ")"

End Sub

Private Function RunCalculation() As String

'
'------------------------------------------------------------------------------
'                               RUN CALCULATION
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the inputs, calculates, replaces the outputs and activates the sheet
'   to review
'
' WHY THIS EXISTS
'   Puts the sequence of an OP_RUN operation in one place: read, calculate,
'   clear, write and show
'
' INPUTS
'   None
'
' RETURNS
'   "RESULT=...; errors=n; warnings=n; positions_used=n; requirement=x; rwa=x"
'
' BEHAVIOR
'   Reads the inputs, calls Calculate, clears every output and writes the
'   outputs with the fingerprint of the inputs. It classifies the outcome as
'   STOPPED, INCOMPLETE (errors found) or OK, activates Checks when the run
'   stopped and Results otherwise (unless mSilent), and returns the result line
'
' ERROR POLICY
'   Propagates ERR_LAYOUT (a workbook name is missing) and any engine error
'
' DEPENDENCIES
'   ReadInputs, Calculate (CORE_Engine), ClearAllOutputs, WriteOutputs,
'   Fingerprint (CORE_Util), InvariantNumber, mSilent; the OUT_ indexes,
'   SH_CHECKS and SH_RESULTS (CORE_Config)
'
' NOTES
'   State ownership: clears and writes the output sheets; activates Results or
'   Checks unless silent. Nothing is cleared before the calculation ended, and
'   the status line is written after every table
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Inputs()    As Variant  'Engine input vector
    Dim Outputs()   As Variant  'Engine output vector
    Dim Outcome     As String   'OK, INCOMPLETE or STOPPED

'------------------------------------------------------------------------------
' READ AND CALCULATE
'------------------------------------------------------------------------------
    'Read the sheets into the engine input vector
        ReadInputs Inputs
    'Calculate every output table
        Calculate Inputs, Outputs

'------------------------------------------------------------------------------
' WRITE OUTPUTS
'------------------------------------------------------------------------------
    'Replace the previous outputs, with the fingerprint of these inputs
        ClearAllOutputs
        WriteOutputs Outputs, Fingerprint(Inputs)

'------------------------------------------------------------------------------
' CLASSIFY OUTCOME
'------------------------------------------------------------------------------
    'Tell a stopped run from an incomplete one and a complete one
        If Outputs(OUT_STOPPED) Then
            Outcome = "STOPPED"
        ElseIf Outputs(OUT_ERRORS) > 0 Then
            Outcome = "INCOMPLETE"
        Else
            Outcome = "OK"
        End If

'------------------------------------------------------------------------------
' ACTIVATE SHEET
'------------------------------------------------------------------------------
    'Show Checks for a stopped run and Results otherwise, unless silent
        If Not mSilent Then
            If Outputs(OUT_STOPPED) Then
                ThisWorkbook.Worksheets(SH_CHECKS).Activate
            Else
                ThisWorkbook.Worksheets(SH_RESULTS).Activate
            End If
        End If

'------------------------------------------------------------------------------
' BUILD RESULT LINE
'------------------------------------------------------------------------------
    'Build the result line, with numbers that use a decimal point
        RunCalculation = "RESULT=" & Outcome & "; errors=" & Outputs(OUT_ERRORS) & _
                         "; warnings=" & InvariantNumber(Outputs(OUT_WARNINGS)) & _
                         "; positions_used=" & InvariantNumber(Outputs(OUT_USED)) & _
                         "; requirement=" & InvariantNumber(Outputs(OUT_REQUIREMENT)) & _
                         "; rwa=" & InvariantNumber(Outputs(OUT_RWA))

End Function

Private Function CheckResults() As String

'
'------------------------------------------------------------------------------
'                         COMPARE RESULTS WITH INPUTS
'------------------------------------------------------------------------------
' PURPOSE
'   Compares the fingerprint of the inputs on the sheets with the one the last
'   run stored, and puts the status line in agreement: the line of the run,
'   with STATUS_OUTDATED before it when the two differ
'
' WHY THIS EXISTS
'   A user can change an input after a run, and the results then no longer
'   belong to the inputs shown
'
' INPUTS
'   None
'
' RETURNS
'   "RESULT=CURRENT" (the results are those of the inputs), "RESULT=OUTDATED"
'   (an input changed after the run), "RESULT=FAILED" (the last run failed) or
'   "RESULT=NOT_CALCULATED" (no result is shown)
'
' BEHAVIOR
'   Reads the status line shown and extracts the line of the run. With no line
'   of a run, the outcome is FAILED when the line starts with STATUS_FAILED and
'   NOT_CALCULATED otherwise. With a line of a run, it reads the inputs,
'   compares their fingerprint with NM_FINGERPRINT and sets the outcome to
'   CURRENT or OUTDATED. It writes the status line only when the line to show
'   differs from the one shown
'
' ERROR POLICY
'   Propagates ERR_LAYOUT (a workbook name is missing)
'
' DEPENDENCIES
'   LineOfRun, ReadInputs, NamedRange, PlainText and Fingerprint (CORE_Util),
'   mStatusWritten; NM_RESULTS_STATUS and NM_FINGERPRINT (CORE_Config)
'
' NOTES
'   State ownership: writes the status line only when it must change, and then
'   sets mStatusWritten
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Inputs()    As Variant  'Engine input vector
    Dim Shown       As String   'Status line on the sheet
    Dim RunLine     As String   'Line the run wrote
    Dim Wanted      As String   'Line to show
    Dim Outcome     As String   'CURRENT, OUTDATED, FAILED or NOT_CALCULATED

'------------------------------------------------------------------------------
' READ STATUS LINE
'------------------------------------------------------------------------------
    'Clear the flag that tells whether this check wrote the line
        mStatusWritten = False
    'Read the status line shown and the line of the run in it
        Shown = PlainText(NamedRange(NM_RESULTS_STATUS).Value)
        RunLine = LineOfRun(Shown)

'------------------------------------------------------------------------------
' COMPARE WITH INPUTS
'------------------------------------------------------------------------------
    'Look for a line of a run
        If Len(RunLine) = 0 Then
            'Tell a failed run from a sheet that shows no result
            If Left$(Shown, Len(STATUS_FAILED)) = STATUS_FAILED Then
                Outcome = "FAILED"
            Else
                Outcome = "NOT_CALCULATED"
            End If
        Else
            'Read the inputs to take their fingerprint
            ReadInputs Inputs
            'Compare the fingerprint with the one the last run stored
            If Fingerprint(Inputs) = PlainText(NamedRange(NM_FINGERPRINT).Value) Then
                Wanted = RunLine
                Outcome = "CURRENT"
            Else
                Wanted = STATUS_OUTDATED & RunLine
                Outcome = "OUTDATED"
            End If
            'Write the status line only when it must change
            If Wanted <> Shown Then
                NamedRange(NM_RESULTS_STATUS).Value = Wanted
                mStatusWritten = True
            End If
        End If

'------------------------------------------------------------------------------
' RETURN OUTCOME
'------------------------------------------------------------------------------
    'Return the outcome
        CheckResults = "RESULT=" & Outcome

End Function

Private Function LineOfRun( _
    ByVal Shown As String) _
    As String
'
'------------------------------------------------------------------------------
'                             EXTRACT LINE OF RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the line a run that ended wrote, from the status line shown
'
' WHY THIS EXISTS
'   The status line may carry the STATUS_OUTDATED prefix, an empty-sheet text
'   or the line of a failed run, and the check needs the line of the run alone
'
' INPUTS
'   Shown
'     Status line on the sheet (String)
'
' RETURNS
'   The line without STATUS_OUTDATED before it; "" when the sheet shows no
'   result: an empty line, STATUS_EMPTY or a line of a failed run
'
' BEHAVIOR
'   Strips the STATUS_OUTDATED prefix when present, returns an empty string for
'   STATUS_EMPTY or a line that starts with STATUS_FAILED, and returns the line
'   as it is otherwise
'
' ERROR POLICY
'   Never raises
'
' DEPENDENCIES
'   STATUS_OUTDATED, STATUS_EMPTY and STATUS_FAILED
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLASSIFY STATUS LINE
'------------------------------------------------------------------------------
    'Strip the OUTDATED prefix, blank a line that shows no result, or keep the line
        If Left$(Shown, Len(STATUS_OUTDATED)) = STATUS_OUTDATED Then
            LineOfRun = Mid$(Shown, Len(STATUS_OUTDATED) + 1)
        ElseIf Shown = STATUS_EMPTY Or Left$(Shown, Len(STATUS_FAILED)) = STATUS_FAILED Then
            LineOfRun = vbNullString
        Else
            LineOfRun = Shown
        End If

End Function

Private Function MarkFailed( _
    ByVal Reason As String) _
    As String
'
'------------------------------------------------------------------------------
'                              MARK RUN AS FAILED
'------------------------------------------------------------------------------
' PURPOSE
'   After a run raised an error: clears every output area, so that no result of
'   that run or of an earlier one stays, and writes a FAILED status line
'
' WHY THIS EXISTS
'   A failed run must not leave the results of an earlier run next to the
'   failure
'
' INPUTS
'   Reason
'     Description of the error the run raised (String)
'
' RETURNS
'   "" when the line was written; else why it was not
'
' BEHAVIOR
'   Clears every output area, adds the areas that could not be cleared to the
'   line, and writes STATUS_FAILED, the time of the run, the reason and the
'   uncleared areas on the status line
'
' ERROR POLICY
'   Contains its own error; never raises. An area that could not be cleared is
'   named on the status line
'
' DEPENDENCIES
'   ClearEveryOutput, NamedRange, TimeText; NM_RESULTS_STATUS and SH_RESULTS
'   (CORE_Config); STATUS_FAILED
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NotCleared  As String   'Areas that could not be cleared

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the handler
        On Error GoTo Failed

'------------------------------------------------------------------------------
' CLEAR OUTPUTS
'------------------------------------------------------------------------------
    'Clear every output area and name those that could not be cleared
        NotCleared = ClearEveryOutput()
        If Len(NotCleared) > 0 Then NotCleared = " Not cleared: " & NotCleared

'------------------------------------------------------------------------------
' WRITE STATUS LINE
'------------------------------------------------------------------------------
    'Write the FAILED line with the time, the reason and the uncleared areas
        NamedRange(NM_RESULTS_STATUS).Value = STATUS_FAILED & "the run of " & _
            TimeText(Now) & " raised an error and no result is shown. " & Reason & NotCleared
    'Report that the line was written
        MarkFailed = vbNullString

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Report why the failure could not be marked
        MarkFailed = "the failure could not be marked on " & SH_RESULTS & ": " & Err.Description

End Function

Private Sub MarkUnchecked()

'
'------------------------------------------------------------------------------
'                           MARK RESULTS AS OUTDATED
'------------------------------------------------------------------------------
' PURPOSE
'   After a check raised an error in an event procedure: marks the results
'   OUTDATED, because they could not be confirmed
'
' WHY THIS EXISTS
'   A check that failed leaves it unknown whether the results belong to the
'   inputs
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads the status line and, when it is the plain line of a run, writes
'   STATUS_OUTDATED before it
'
' ERROR POLICY
'   Contains its own error; never raises. When the status line cannot be read
'   or written, nothing is marked
'
' DEPENDENCIES
'   LineOfRun, NamedRange, PlainText (CORE_Util); NM_RESULTS_STATUS
'   (CORE_Config); STATUS_OUTDATED
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Shown   As String   'Status line on the sheet
    Dim RunLine As String   'Line the run wrote

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the handler
        On Error GoTo Failed

'------------------------------------------------------------------------------
' READ STATUS LINE
'------------------------------------------------------------------------------
    'Read the status line shown and the line of the run in it
        Shown = PlainText(NamedRange(NM_RESULTS_STATUS).Value)
        RunLine = LineOfRun(Shown)

'------------------------------------------------------------------------------
' MARK OUTDATED
'------------------------------------------------------------------------------
    'Put STATUS_OUTDATED before the line of a run that is not marked yet
        If Len(RunLine) > 0 And Shown = RunLine Then
            NamedRange(NM_RESULTS_STATUS).Value = STATUS_OUTDATED & RunLine
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Nothing to report to: see the error policy

End Sub

'
'------------------------------------------------------------------------------
'
'                                   READING
'
'------------------------------------------------------------------------------
'

Private Sub ReadInputs( _
    ByRef Inputs() As Variant)
'
'------------------------------------------------------------------------------
'                           READ INPUTS FROM SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the engine input vector from Params, OptionCharges, the four
'   position sheets and Eligibility
'
' WHY THIS EXISTS
'   CORE_Engine takes its inputs as one vector of values, and this procedure is
'   where the sheets and workbook names are turned into it
'
' INPUTS
'   Inputs
'     Vector to fill, sized here to 1 To IN_COUNT (Variant array, output)
'
' RETURNS
'   Nothing; Inputs(1 To IN_COUNT) is filled as CORE_Config describes
'
' BEHAVIOR
'   Sizes Inputs, then reads the numeric parameters and the other Params
'   values, the ladder matrix and its labels, the specific-risk and foreign
'   exchange tables, the option charges, the four position tables with the
'   sheet row of each first data row, the capacity of the equity and
'   commodity blocks of Buckets, and the ELIGIBILITY_ROWS rows of Eligibility
'   below its header, whatever the regime
'
' ERROR POLICY
'   Raises ERR_LAYOUT when a workbook name is missing
'
' DEPENDENCIES
'   NamedRange, ColumnValues, PairedColumns, TableValues, PlainText and
'   ListItem (CORE_Util); the IN_, PI_, PRM_, NM_, LD_, RC_, EL_ and
'   table-width constants of CORE_Config
'
' NOTES
'   The rows of Eligibility are read under the Basel regime too, so that the
'   fingerprint of the inputs covers them; the engine assesses them under the
'   CRR only
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Numeric(1 To PI_COUNT)  As Variant  'Numeric parameters
    Dim Options(1 To RC_COUNT)  As Variant  'Option charges
    Dim Ladder()                As Variant  'Ladder matrix
    Dim Labels()                As String   'Ladder labels
    Dim Column                  As Variant  'One named column as read
    Dim FirstRow                As Long     'First data row of a table
    Dim b                       As Long     'Band
    Dim i                       As Long     'Parameter or class

'------------------------------------------------------------------------------
' SIZE INPUT VECTOR
'------------------------------------------------------------------------------
    'Size the input vector
        ReDim Inputs(1 To IN_COUNT)

'------------------------------------------------------------------------------
' READ PARAMETERS
'------------------------------------------------------------------------------
    'Read each numeric parameter by its code
        For i = 1 To PI_COUNT
            Numeric(i) = NamedRange(ListItem(PRM_NUMERIC_CODES, i)).Value
        Next i
    'Store the numeric parameters
        Inputs(IN_NUMERIC) = Numeric
    'Read the as-of date, reporting currency, gold code, regime, transition end and own funds
        Inputs(IN_ASOF) = NamedRange(PRM_ASOF).Value
        Inputs(IN_REPCCY) = NamedRange(PRM_REPCCY).Value
        Inputs(IN_GOLD) = NamedRange(PRM_GOLD).Value
        Inputs(IN_REGIME) = NamedRange(PRM_REGIME).Value
        Inputs(IN_TRANSITION_END) = NamedRange(PRM_TRANSITION_END).Value
        Inputs(IN_OWN_FUNDS) = NamedRange(PRM_OWN_FUNDS).Value

'------------------------------------------------------------------------------
' READ LADDER
'------------------------------------------------------------------------------
    'Size the ladder matrix and its labels
        ReDim Ladder(1 To BAND_COUNT, 1 To LD_WIDTH)
        ReDim Labels(1 To BAND_COUNT, 1 To 2)
    'Read the zone of each band
        Column = ColumnValues(NM_LADDER_ZONE, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Ladder(b, LD_ZONE) = Column(b, 1)
        Next b
    'Read the upper limit of each band, high column
        Column = ColumnValues(NM_LADDER_UP_HIGH, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Ladder(b, LD_UPPER_HIGH) = Column(b, 1)
        Next b
    'Read the upper limit of each band, low column
        Column = ColumnValues(NM_LADDER_UP_LOW, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Ladder(b, LD_UPPER_LOW) = Column(b, 1)
        Next b
    'Read the risk weight of each band
        Column = ColumnValues(NM_LADDER_RW, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Ladder(b, LD_WEIGHT) = Column(b, 1)
        Next b
    'Read the label of each band, high column
        Column = ColumnValues(NM_LADDER_LBL_HIGH, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Labels(b, 1) = PlainText(Column(b, 1))
        Next b
    'Read the label of each band, low column
        Column = ColumnValues(NM_LADDER_LBL_LOW, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Labels(b, 2) = PlainText(Column(b, 1))
        Next b
    'Store the ladder and its labels
        Inputs(IN_LADDER) = Ladder
        Inputs(IN_LADDER_LABELS) = Labels

'------------------------------------------------------------------------------
' READ TABLES
'------------------------------------------------------------------------------
    'Read the specific-risk table and the foreign exchange table
        Inputs(IN_SR_TABLE) = PairedColumns(NM_SR_KEY, NM_SR_CHARGE)
        Inputs(IN_FX_TABLE) = PairedColumns(NM_FX_CCY, NM_FX_RATE)
    'Read each option charge by its code
        For i = 1 To RC_COUNT
            Options(i) = NamedRange(NM_OPTIONS_PREFIX & ListItem(RC_CODES, i)).Value
        Next i
        Inputs(IN_OPTIONS) = Options

'------------------------------------------------------------------------------
' READ POSITIONS
'------------------------------------------------------------------------------
    'Read the interest rate positions and their first row
        Inputs(IN_IR_ROWS) = TableValues(NM_IR_POS_HEADER, IR_INPUT_WIDTH, FirstRow)
        Inputs(IN_IR_FIRST_ROW) = FirstRow
    'Read the equity positions and their first row
        Inputs(IN_EQ_ROWS) = TableValues(NM_EQ_POS_HEADER, EQ_INPUT_WIDTH, FirstRow)
        Inputs(IN_EQ_FIRST_ROW) = FirstRow
    'Read the commodity positions and their first row
        Inputs(IN_CO_ROWS) = TableValues(NM_CO_POS_HEADER, CO_INPUT_WIDTH, FirstRow)
        Inputs(IN_CO_FIRST_ROW) = FirstRow
    'Read the foreign exchange positions and their first row
        Inputs(IN_FX_ROWS) = TableValues(NM_FX_POS_HEADER, FX_INPUT_WIDTH, FirstRow)
        Inputs(IN_FX_FIRST_ROW) = FirstRow

'------------------------------------------------------------------------------
' READ CAPACITIES
'------------------------------------------------------------------------------
    'Count the bucket rows that fit in the equity and commodity blocks of Buckets
        Inputs(IN_EQ_CAPACITY) = NamedRange(NM_EQ_GENERAL).Row - NamedRange(NM_BK_EQ_HEADER).Row - 1
        Inputs(IN_CO_CAPACITY) = NamedRange(NM_CO_DIRECTIONAL).Row - NamedRange(NM_BK_CO_HEADER).Row - 1

'------------------------------------------------------------------------------
' READ ELIGIBILITY
'------------------------------------------------------------------------------
    'Read every row of Eligibility below its header, input columns only
        Inputs(IN_EL_ROWS) = NamedRange(NM_EL_HEADER).Offset(1, 0).Resize(ELIGIBILITY_ROWS, _
                                                                          EL_INPUT_WIDTH).Value

End Sub

Private Function TableValues( _
    ByVal HeaderName As String, _
    ByVal Width As Long, _
    ByRef FirstRow As Long) _
    As Variant
'
'------------------------------------------------------------------------------
'                             READ POSITION TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Reads an input table from the row below its header to the last row with a
'   value in its first column
'
' WHY THIS EXISTS
'   The position tables have no fixed length, and the engine needs their rows
'   and the sheet row where they start
'
' INPUTS
'   HeaderName
'     Workbook name of the table's first header cell (String)
'
'   Width
'     Columns to read (Long, >= 2)
'
'   FirstRow
'     Sheet row of the first data row (Long, ByRef, output)
'
' RETURNS
'   Variant (1 To n, 1 To Width), n >= 1 (one blank row for an empty table);
'   FirstRow = sheet row of the first data row
'
' BEHAVIOR
'   Locates the header cell, takes the row below it as the first data row,
'   finds the last row with a value in the first column of the table (at least
'   the first data row) and reads the block
'
' ERROR POLICY
'   Raises ERR_LAYOUT when the name is missing
'
' DEPENDENCIES
'   NamedRange; Range.End(xlUp)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header  As Range        'First header cell
    Dim Sheet   As Worksheet    'Sheet of the table
    Dim LastRow As Long         'Last row with a PositionID

'------------------------------------------------------------------------------
' LOCATE TABLE
'------------------------------------------------------------------------------
    'Find the header cell, its sheet and the first data row
        Set Header = NamedRange(HeaderName)
        Set Sheet = Header.Worksheet
        FirstRow = Header.Row + 1

'------------------------------------------------------------------------------
' FIND LAST ROW
'------------------------------------------------------------------------------
    'Take the last row with a value in the first column, at least the first data row
        LastRow = Sheet.Cells(Sheet.Rows.Count, Header.Column).End(xlUp).Row
        If LastRow < FirstRow Then LastRow = FirstRow

'------------------------------------------------------------------------------
' READ TABLE
'------------------------------------------------------------------------------
    'Read the block of values
        TableValues = Sheet.Range(Sheet.Cells(FirstRow, Header.Column), _
                                  Sheet.Cells(LastRow, Header.Column + Width - 1)).Value

End Function

Private Function ColumnValues( _
    ByVal RangeName As String, _
    ByVal ExpectedCount As Long) _
    As Variant
'
'------------------------------------------------------------------------------
'                              READ NAMED COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   Reads a named single-column range of a known number of cells
'
' WHY THIS EXISTS
'   The ladder and the paired tables are columns of fixed size, and a range of
'   another size must be reported, not read
'
' INPUTS
'   RangeName
'     Workbook name of the column (String)
'
'   ExpectedCount
'     Cells the range must have (Long, >= 2)
'
' RETURNS
'   Variant (1 To ExpectedCount, 1 To 1)
'
' BEHAVIOR
'   Locates the named range, checks that it has ExpectedCount rows and one
'   column, and returns its values
'
' ERROR POLICY
'   Raises ERR_LAYOUT when the name is missing or has another size
'
' DEPENDENCIES
'   NamedRange; ERR_LAYOUT (CORE_Config)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target  As Range    'Named range

'------------------------------------------------------------------------------
' LOCATE RANGE
'------------------------------------------------------------------------------
    'Find the named range
        Set Target = NamedRange(RangeName)

'------------------------------------------------------------------------------
' VALIDATE SIZE
'------------------------------------------------------------------------------
    'Reject a range that is not one column of the expected size
        If Target.Rows.Count <> ExpectedCount Or Target.Columns.Count <> 1 Then
            Err.Raise ERR_LAYOUT, "M_Main.ColumnValues", _
                "Workbook name " & RangeName & " must refer to " & ExpectedCount & " cells in one column."
        End If

'------------------------------------------------------------------------------
' RETURN VALUES
'------------------------------------------------------------------------------
    'Return the values of the range
        ColumnValues = Target.Value

End Function

Private Function PairedColumns( _
    ByVal FirstName As String, _
    ByVal SecondName As String) _
    As Variant
'
'------------------------------------------------------------------------------
'                            READ TWO NAMED COLUMNS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads two named columns of equal length as one two-column table
'
' WHY THIS EXISTS
'   The specific-risk and foreign exchange tables are kept in two columns that
'   the engine takes as one table
'
' INPUTS
'   FirstName
'     Workbook name of the first column (String)
'
'   SecondName
'     Workbook name of the second column (String)
'
' RETURNS
'   Variant (1 To n, 1 To 2)
'
' BEHAVIOR
'   Takes the row count of the first column, reads both columns with that count
'   and copies them into one table
'
' ERROR POLICY
'   Raises ERR_LAYOUT when a name is missing or the lengths differ
'
' DEPENDENCIES
'   NamedRange, ColumnValues
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FirstValues     As Variant  'First column values
    Dim SecondValues    As Variant  'Second column values
    Dim Table()         As Variant  'Result
    Dim RowCount        As Long     'Rows
    Dim r               As Long     'Row

'------------------------------------------------------------------------------
' READ COLUMNS
'------------------------------------------------------------------------------
    'Read both columns with the row count of the first
        RowCount = NamedRange(FirstName).Rows.Count
        FirstValues = ColumnValues(FirstName, RowCount)
        SecondValues = ColumnValues(SecondName, RowCount)

'------------------------------------------------------------------------------
' BUILD TABLE
'------------------------------------------------------------------------------
    'Copy the two columns into one table
        ReDim Table(1 To RowCount, 1 To 2)
        For r = 1 To RowCount
            Table(r, 1) = FirstValues(r, 1)
            Table(r, 2) = SecondValues(r, 1)
        Next r

'------------------------------------------------------------------------------
' RETURN TABLE
'------------------------------------------------------------------------------
    'Return the table
        PairedColumns = Table

End Function

'
'------------------------------------------------------------------------------
'
'                                   WRITING
'
'------------------------------------------------------------------------------
'

Private Sub WriteOutputs( _
    ByRef Outputs() As Variant, _
    ByVal InputFingerprint As String)
'
'------------------------------------------------------------------------------
'                           WRITE OUTPUTS TO SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the engine outputs to their sheets; the sheets were cleared before.
'   The result tables come first, then Checks, then the fingerprint of the
'   inputs, and the status line last: a write that raises an error leaves the
'   line ClearAllOutputs wrote
'
' WHY THIS EXISTS
'   A write that fails half way must not leave the status line of a complete
'   run, so the status line is written after every table
'
' INPUTS
'   Outputs
'     Vector (1 To OUT_COUNT) from CORE_Engine.Calculate (Variant array)
'
'   InputFingerprint
'     Fingerprint of the inputs the outputs come from (String)
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the Results tables unless the run stopped, then the Checks rules,
'   the error and warning counts and the findings, then the fingerprint, and
'   last the status line with the time of the run
'
' ERROR POLICY
'   Propagates the errors of its callees (ERR_LAYOUT when a workbook name or
'   rule is missing)
'
' DEPENDENCIES
'   WriteResults, WriteRules, WriteBlock, NamedRange, TimeText; the OUT_
'   indexes and NM_ names of CORE_Config
'
' NOTES
'   State ownership: writes IRCalc, EQCalc, COCalc, FXCalc, IR_Ladder,
'   IR_Zones, Buckets, Results and Checks, and the calculated columns and the
'   result block of Eligibility
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE RESULTS AND CHECKS
'------------------------------------------------------------------------------
    'Write the result tables of a run that was not stopped
        If Not Outputs(OUT_STOPPED) Then WriteResults Outputs
    'Write the rule results and the error and warning counts
        WriteRules Outputs(OUT_RULES)
        NamedRange(NM_CHECKS_ERRORS).Value = Outputs(OUT_ERRORS)
        NamedRange(NM_CHECKS_WARNINGS).Value = Outputs(OUT_WARNINGS)
    'Write the findings table
        WriteBlock NamedRange(NM_FINDINGS_HEADER).Offset(1, 0), Outputs(OUT_FINDINGS)

'------------------------------------------------------------------------------
' WRITE FINGERPRINT AND STATUS
'------------------------------------------------------------------------------
    'Store the fingerprint of the inputs
        NamedRange(NM_FINGERPRINT).Value = InputFingerprint
    'Write the status line last, with the time of the run
        NamedRange(NM_RESULTS_STATUS).Value = Outputs(OUT_STATUS) & " Run " & TimeText(Now) & "."

End Sub

Private Sub WriteResults( _
    ByRef Outputs() As Variant)
'
'------------------------------------------------------------------------------
'                             WRITE RESULT TABLES
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the result tables of a run that was not stopped
'
' WHY THIS EXISTS
'   A stopped run has no result tables, and the other runs write the same set
'   of blocks
'
' INPUTS
'   Outputs
'     Vector (1 To OUT_COUNT) from CORE_Engine.Calculate (Variant array)
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the IRCalc, EQCalc, COCalc and FXCalc tables, the ladder, the zones
'   and the Results table, then the equity and commodity blocks of Buckets, the
'   currency block and the FX ONOP and charge cells, and last the calculated
'   columns of Eligibility, right of its input columns, and its result block
'
' ERROR POLICY
'   Propagates ERR_LAYOUT when a workbook name is missing. Raises ERR_LAYOUT
'   when Eligibility_Results is not EL_BLOCK_ROWS rows by EL_BLOCK_WIDTH
'   columns
'
' DEPENDENCIES
'   WriteBlock, WriteBuckets, NamedRange; the OUT_ indexes, NM_ names, EL_
'   widths and ERR_LAYOUT of CORE_Config
'
' NOTES
'   State ownership: writes IRCalc, EQCalc, COCalc, FXCalc, IR_Ladder,
'   IR_Zones, Buckets, the table of Results, and the columns I to U and the
'   result block of Eligibility
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header          As Range    'Header cell of a table
    Dim MarketTable     As Variant  'Buckets equity block
    Dim CommodityTable  As Variant  'Buckets commodity block
    Dim SummaryTable    As Variant  'FX ONOP and charge
    Dim Capacity        As Long     'Rows of a Buckets block
    Dim Shown           As Long     'Rows written to a Buckets block

'------------------------------------------------------------------------------
' WRITE CALCULATION TABLES
'------------------------------------------------------------------------------
    'Write the interest rate, equity, commodity and FX calculation tables
        WriteBlock NamedRange(NM_IR_CALC_HEADER).Offset(1, 0), Outputs(OUT_IR_CALC)
        WriteBlock NamedRange(NM_EQ_CALC_HEADER).Offset(1, 0), Outputs(OUT_EQ_CALC)
        WriteBlock NamedRange(NM_CO_CALC_HEADER).Offset(1, 0), Outputs(OUT_CO_CALC)
        WriteBlock NamedRange(NM_FX_CALC_HEADER).Offset(1, 0), Outputs(OUT_FX_CALC)

'------------------------------------------------------------------------------
' WRITE LADDER, ZONES AND RESULTS
'------------------------------------------------------------------------------
    'Write the ladder, the zones and the Results table
        WriteBlock NamedRange(NM_LADDER_HEADER).Offset(1, 0), Outputs(OUT_LADDER)
        WriteBlock NamedRange(NM_ZONES_HEADER).Offset(1, 0), Outputs(OUT_ZONES)
        WriteBlock NamedRange(NM_RESULTS_HEADER).Offset(1, RESULTS_COLUMN - 1), Outputs(OUT_RESULTS)

'------------------------------------------------------------------------------
' WRITE BUCKETS
'------------------------------------------------------------------------------
    'Write the Buckets blocks: as many rows as fit, then the total row at its name
        Set Header = NamedRange(NM_BK_EQ_HEADER)
        Capacity = NamedRange(NM_EQ_GENERAL).Row - Header.Row - 1
        MarketTable = Outputs(OUT_MARKETS)
        Shown = WriteBuckets(Header, MarketTable, Capacity)
    'Write the commodity block the same way
        Set Header = NamedRange(NM_BK_CO_HEADER)
        Capacity = NamedRange(NM_CO_DIRECTIONAL).Row - Header.Row - 1
        CommodityTable = Outputs(OUT_COMMODITIES)
        Shown = WriteBuckets(Header, CommodityTable, Capacity)
    'Write the currency block
        WriteBlock NamedRange(NM_BK_FX_HEADER).Offset(1, 0), Outputs(OUT_CURRENCIES)
    'Write the FX ONOP and charge cells
        SummaryTable = Outputs(OUT_FX_SUMMARY)
        NamedRange(NM_FX_ONOP).Value = SummaryTable(1, 1)
        NamedRange(NM_FX_CHARGE).Value = SummaryTable(2, 1)

'------------------------------------------------------------------------------
' WRITE ELIGIBILITY
'------------------------------------------------------------------------------
    'Write the calculated columns of each row of Eligibility, right of its input columns
        WriteBlock NamedRange(NM_EL_HEADER).Offset(1, EL_INPUT_WIDTH), Outputs(OUT_EL_CALC)
    'Reject a result block of another size, then write the status and the figures of each approach
        Set Header = NamedRange(NM_EL_RESULTS)
        If Header.Rows.Count <> EL_BLOCK_ROWS Or Header.Columns.Count <> EL_BLOCK_WIDTH Then
            Err.Raise ERR_LAYOUT, "M_Main.WriteResults", _
                "Workbook name " & NM_EL_RESULTS & " must refer to " & EL_BLOCK_ROWS & _
                " rows of " & EL_BLOCK_WIDTH & " cells."
        End If
        WriteBlock Header, Outputs(OUT_EL_RESULTS)

End Sub

Private Function WriteBuckets( _
    ByVal Header As Range, _
    ByRef Table As Variant, _
    ByVal Capacity As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                             WRITE BUCKETS BLOCK
'------------------------------------------------------------------------------
' PURPOSE
'   Writes an equity or commodity block of Buckets: up to Capacity bucket rows
'   below the header, and the total row (the table's last row) on the row after
'   the capacity. Totals include every bucket; rule C10 / C11 reports buckets
'   not shown
'
' WHY THIS EXISTS
'   The block has a fixed number of bucket rows and a total row below them, and
'   the tables of the engine can have more buckets than fit
'
' INPUTS
'   Header
'     First header cell of the block (Range)
'
'   Table
'     (1 To k + 1, 1 To width): k buckets and the total row (Variant)
'
'   Capacity
'     Bucket rows of the block (Long, >= 1)
'
' RETURNS
'   Bucket rows written (Long)
'
' BEHAVIOR
'   Writes the first Capacity bucket rows of Table, or all of them when fewer,
'   below the header, then writes the last row of Table, the total row, on the
'   row after the capacity
'
' ERROR POLICY
'   Propagates the error of a failed write
'
' DEPENDENCIES
'   Range.Offset and Range.Resize
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Shown   As Long     'Bucket rows written
    Dim Part()  As Variant  'Rows written
    Dim Width   As Long     'Columns
    Dim r       As Long     'Row
    Dim c       As Long     'Column

'------------------------------------------------------------------------------
' MEASURE TABLE
'------------------------------------------------------------------------------
    'Take the width and the number of bucket rows to show
        Width = UBound(Table, 2)
        Shown = UBound(Table, 1) - 1
        If Shown > Capacity Then Shown = Capacity

'------------------------------------------------------------------------------
' WRITE BUCKET ROWS
'------------------------------------------------------------------------------
    'Write the bucket rows that fit
        If Shown > 0 Then
            ReDim Part(1 To Shown, 1 To Width)
            For r = 1 To Shown
                For c = 1 To Width
                    Part(r, c) = Table(r, c)
                Next c
            Next r
            Header.Offset(1, 0).Resize(Shown, Width).Value = Part
        End If

'------------------------------------------------------------------------------
' WRITE TOTAL ROW
'------------------------------------------------------------------------------
    'Write the total row on the row after the capacity
        ReDim Part(1 To 1, 1 To Width)
        For c = 1 To Width
            Part(1, c) = Table(UBound(Table, 1), c)
        Next c
        Header.Offset(Capacity + 1, 0).Resize(1, Width).Value = Part

'------------------------------------------------------------------------------
' RETURN COUNT
'------------------------------------------------------------------------------
    'Return the bucket rows written
        WriteBuckets = Shown

End Function

Private Sub WriteRules( _
    ByVal Rules As Variant)
'
'------------------------------------------------------------------------------
'                              WRITE CHECKS RULES
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the count and result of each Checks rule on the row whose Rule cell
'   holds its identifier
'
' WHY THIS EXISTS
'   The Checks sheet lists the rules by identifier, and the engine returns them
'   in its own order
'
' INPUTS
'   Rules
'     (1 To RULE_COUNT, 1 To 2): identifier, count or Empty (Variant)
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Searches the rows under the Checks header for each rule identifier. For a
'   count of Empty it clears the count and writes NOT_RUN as the result;
'   otherwise it writes the count and ST_OK for a count of 0 or the severity of
'   the rule
'
' ERROR POLICY
'   Raises ERR_LAYOUT when a rule identifier is not found under the header
'
' DEPENDENCIES
'   NamedRange, PlainText and ListItem (CORE_Util); NM_CHECKS_HEADER,
'   RULE_COUNT, ST_OK, RULE_SEVERITIES and SH_CHECKS (CORE_Config)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header      As Range    'Rule header cell
    Dim Found       As Boolean  'Rule row found
    Dim r           As Long     'Rule
    Dim RowOffset   As Long     'Row offset under the header

'------------------------------------------------------------------------------
' LOCATE HEADER
'------------------------------------------------------------------------------
    'Find the header of the rule table
        Set Header = NamedRange(NM_CHECKS_HEADER)

'------------------------------------------------------------------------------
' WRITE EACH RULE
'------------------------------------------------------------------------------
    'Handle the rules in turn
        For r = 1 To UBound(Rules, 1)
            'Search the rows under the header for the identifier of the rule
            Found = False
            For RowOffset = 1 To RULE_COUNT + 5
                If PlainText(Header.Offset(RowOffset, 0).Value) = Rules(r, 1) Then
                    Found = True
                    Exit For
                End If
            Next RowOffset
            'Stop when the rule has no row
            If Not Found Then
                Err.Raise ERR_LAYOUT, "M_Main.WriteRules", _
                    "Rule " & Rules(r, 1) & " not found on " & SH_CHECKS & "."
            End If
            'Write the count and the result of the rule
            If IsEmpty(Rules(r, 2)) Then
                Header.Offset(RowOffset, CHECKS_COUNT_COLUMN - 1).Value = Empty
                Header.Offset(RowOffset, CHECKS_RESULT_COLUMN - 1).Value = NOT_RUN
            Else
                Header.Offset(RowOffset, CHECKS_COUNT_COLUMN - 1).Value = Rules(r, 2)
                If Rules(r, 2) = 0 Then
                    Header.Offset(RowOffset, CHECKS_RESULT_COLUMN - 1).Value = ST_OK
                Else
                    Header.Offset(RowOffset, CHECKS_RESULT_COLUMN - 1).Value = ListItem(RULE_SEVERITIES, r)
                End If
            End If
        Next r

End Sub

Private Sub WriteBlock( _
    ByVal TopLeft As Range, _
    ByVal Table As Variant)
'
'------------------------------------------------------------------------------
'                                 WRITE TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a two-dimensional array with its top-left cell at TopLeft
'
' WHY THIS EXISTS
'   Every output table is written the same way: resize the range to the table
'   and assign the values
'
' INPUTS
'   TopLeft
'     First cell to write (Range)
'
'   Table
'     Variant (1 To n, 1 To m), or Empty to write nothing
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Returns without writing when Table is Empty; otherwise writes Table to the
'   range of its size that starts at TopLeft
'
' ERROR POLICY
'   Propagates the error of a failed write
'
' DEPENDENCIES
'   Range.Resize
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE TABLE
'------------------------------------------------------------------------------
    'Write nothing for an empty table
        If IsEmpty(Table) Then Exit Sub
    'Write the table to the range of its size
        TopLeft.Resize(UBound(Table, 1), UBound(Table, 2)).Value = Table

End Sub

'
'------------------------------------------------------------------------------
'
'                                   CLEARING
'
'------------------------------------------------------------------------------
'

Private Sub ClearAllOutputs()

'
'------------------------------------------------------------------------------
'                              CLEAR ALL OUTPUTS
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the contents (not the formats) of every output area
'
' WHY THIS EXISTS
'   Gives the operations a clearing step that raises when an area could not be
'   cleared
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears every output area with ClearEveryOutput and raises an error that
'   names the areas that could not be cleared
'
' ERROR POLICY
'   Raises ERR_LAYOUT, after every other area was cleared, when an area could
'   not be
'
' DEPENDENCIES
'   ClearEveryOutput; ERR_LAYOUT (CORE_Config)
'
' NOTES
'   State ownership: clears the output sheets; never touches inputs, Params or
'   formulas outside the output areas
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NotCleared  As String   'Areas that could not be cleared

'------------------------------------------------------------------------------
' CLEAR OUTPUT AREAS
'------------------------------------------------------------------------------
    'Clear every area and collect the failures
        NotCleared = ClearEveryOutput()

'------------------------------------------------------------------------------
' REPORT UNCLEARED AREAS
'------------------------------------------------------------------------------
    'Raise an error that names the areas that could not be cleared
        If Len(NotCleared) > 0 Then Err.Raise ERR_LAYOUT, "M_Main.ClearAllOutputs", NotCleared

End Sub

Private Function ClearEveryOutput() As String

'
'------------------------------------------------------------------------------
'                           CLEAR EVERY OUTPUT AREA
'------------------------------------------------------------------------------
' PURPOSE
'   Clears every output area, going on after one that cannot be cleared
'
' WHY THIS EXISTS
'   MarkFailed and ClearAllOutputs both need every area cleared as far as
'   possible and the failures reported, not raised
'
' INPUTS
'   None
'
' RETURNS
'   "" when every area was cleared; else why each of the others was not,
'   separated by "; "
'
' BEHAVIOR
'   Calls ClearPart for each part from 1 to CLEAR_PARTS and joins the failure
'   messages
'
' ERROR POLICY
'   Never raises: ClearPart contains the error of its part
'
' DEPENDENCIES
'   ClearPart, AppendDetail; CLEAR_PARTS
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NotCleared  As String   'Areas that could not be cleared
    Dim Part        As Long     'Part of the outputs

'------------------------------------------------------------------------------
' CLEAR EACH PART
'------------------------------------------------------------------------------
    'Clear the parts in turn and collect the failures
        For Part = 1 To CLEAR_PARTS
            NotCleared = AppendDetail(NotCleared, ClearPart(Part))
        Next Part

'------------------------------------------------------------------------------
' RETURN FAILURES
'------------------------------------------------------------------------------
    'Return the failures, empty when every part was cleared
        ClearEveryOutput = NotCleared

End Function

Private Function ClearPart( _
    ByVal Part As Long) _
    As String
'
'------------------------------------------------------------------------------
'                                CLEAR ONE PART
'------------------------------------------------------------------------------
' PURPOSE
'   Clears one part of the outputs. The status line is the first part, so that
'   no later failure leaves the line of a run over cleared tables
'
' WHY THIS EXISTS
'   One failed area must not stop the clearing of the others, so each part has
'   its own error handling
'
' INPUTS
'   Part
'     Part number (Long, 1 To CLEAR_PARTS)
'
' RETURNS
'   "" on success, else what failed
'
' BEHAVIOR
'   Clears the part that Part selects: the status line, the fingerprint, the
'   calculation tables, the findings, the zones, the Buckets blocks (with the
'   label of each total row put back), the FX ONOP and charge cells, the
'   Results table, the Checks cells, and the calculated columns and the result
'   block of Eligibility
'
' ERROR POLICY
'   Contains its own error; never raises
'
' DEPENDENCIES
'   NamedRange, ClearToEnd, ClearRows; the NM_ names, table widths, LBL_TOTAL_
'   labels, RULE_COUNT, RESULTS_, ELIGIBILITY_ROWS and EL_ constants of
'   CORE_Config; STATUS_EMPTY and CHECKS_COUNT_COLUMN
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
    Dim Header  As Range    'Header cell of a table
    Dim r       As Long     'Rule row

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the handler
        On Error GoTo Failed

'------------------------------------------------------------------------------
' CLEAR PART
'------------------------------------------------------------------------------
    'Choose the area by part number
        Select Case Part
            'Reset the status line first
            Case 1
                NamedRange(NM_RESULTS_STATUS).Value = STATUS_EMPTY
            'Clear the fingerprint of the inputs
            Case 2
                NamedRange(NM_FINGERPRINT).ClearContents
            'Clear the interest rate, equity, commodity and FX calculation tables
            Case 3
                ClearToEnd NM_IR_CALC_HEADER, IR_CALC_WIDTH
            Case 4
                ClearToEnd NM_EQ_CALC_HEADER, EQ_CALC_WIDTH
            Case 5
                ClearToEnd NM_CO_CALC_HEADER, CO_CALC_WIDTH
            Case 6
                ClearToEnd NM_FX_CALC_HEADER, FX_CALC_WIDTH
            'Clear the ladder and the findings
            Case 7
                ClearToEnd NM_LADDER_HEADER, LADDER_WIDTH
            Case 8
                ClearToEnd NM_FINDINGS_HEADER, FINDINGS_WIDTH
            'The label of a total row sits in the cleared area: put it back
            Case 9
                ClearRows NM_ZONES_HEADER, NamedRange(NM_IR_GMR_TOTAL).Row, ZONES_WIDTH
                NamedRange(NM_IR_GMR_TOTAL).Worksheet.Cells(NamedRange(NM_IR_GMR_TOTAL).Row, _
                    NamedRange(NM_ZONES_HEADER).Column).Value = LBL_TOTAL_ZONES
            Case 10
                ClearRows NM_BK_EQ_HEADER, NamedRange(NM_EQ_GENERAL).Row, BK_EQ_WIDTH
                NamedRange(NM_BK_EQ_HEADER).Worksheet.Cells(NamedRange(NM_EQ_GENERAL).Row, _
                    NamedRange(NM_BK_EQ_HEADER).Column + 1).Value = LBL_TOTAL_EQ
            Case 11
                ClearRows NM_BK_CO_HEADER, NamedRange(NM_CO_DIRECTIONAL).Row, BK_CO_WIDTH
                NamedRange(NM_BK_CO_HEADER).Worksheet.Cells(NamedRange(NM_CO_DIRECTIONAL).Row, _
                    NamedRange(NM_BK_CO_HEADER).Column + 1).Value = LBL_TOTAL_CO
            Case 12
                ClearRows NM_BK_FX_HEADER, NamedRange(NM_BK_FX_HEADER).Row + _
                          NamedRange(NM_FX_CCY).Rows.Count + 1, BK_FX_WIDTH
                NamedRange(NM_BK_FX_HEADER).Offset(NamedRange(NM_FX_CCY).Rows.Count + 1, 1).Value = _
                    LBL_TOTAL_FX
            'Clear the FX ONOP and charge cells
            Case 13
                NamedRange(NM_FX_ONOP).ClearContents
            Case 14
                NamedRange(NM_FX_CHARGE).ClearContents
            'Clear the Results table
            Case 15
                Set Header = NamedRange(NM_RESULTS_HEADER)
                Header.Offset(1, RESULTS_COLUMN - 1).Resize(RESULTS_ROWS, RESULTS_WIDTH).ClearContents
            'Clear the count and result of each Checks rule
            Case 16
                Set Header = NamedRange(NM_CHECKS_HEADER)
                For r = 1 To RULE_COUNT
                    Header.Offset(r, CHECKS_COUNT_COLUMN - 1).Resize(1, 2).ClearContents
                Next r
            'Clear the error and warning counts of Checks
            Case 17
                NamedRange(NM_CHECKS_ERRORS).ClearContents
            Case 18
                NamedRange(NM_CHECKS_WARNINGS).ClearContents
            'Clear the calculated columns of Eligibility, never its input columns, and its result block
            Case 19
                NamedRange(NM_EL_HEADER).Offset(1, EL_INPUT_WIDTH).Resize(ELIGIBILITY_ROWS, _
                                                                          EL_CALC_WIDTH).ClearContents
            Case 20
                NamedRange(NM_EL_RESULTS).ClearContents
        End Select

'------------------------------------------------------------------------------
' REPORT SUCCESS
'------------------------------------------------------------------------------
    'Report that the part was cleared
        ClearPart = vbNullString

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return the description of the error
        ClearPart = Err.Description

End Function

Private Sub ClearToEnd( _
    ByVal HeaderName As String, _
    ByVal Width As Long)
'
'------------------------------------------------------------------------------
'                           CLEAR TABLE TO LAST ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Clears a table from the row below its header to the last used row of its
'   sheet
'
' WHY THIS EXISTS
'   A table that grows with its data must be cleared however many rows it has
'
' INPUTS
'   HeaderName
'     Workbook name of the first header cell (String)
'
'   Width
'     Columns of the table (Long)
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Finds the last used row of the header's sheet and clears the contents of
'   the block below the header, Width columns wide, down to that row
'
' ERROR POLICY
'   Propagates ERR_LAYOUT when the name is missing
'
' DEPENDENCIES
'   NamedRange; Worksheet.UsedRange
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header  As Range    'First header cell
    Dim LastRow As Long     'Last used row of the sheet

'------------------------------------------------------------------------------
' LOCATE TABLE
'------------------------------------------------------------------------------
    'Find the header cell
        Set Header = NamedRange(HeaderName)

'------------------------------------------------------------------------------
' FIND LAST ROW
'------------------------------------------------------------------------------
    'Take the last used row of the sheet
        With Header.Worksheet.UsedRange
            LastRow = .Row + .Rows.Count - 1
        End With

'------------------------------------------------------------------------------
' CLEAR ROWS
'------------------------------------------------------------------------------
    'Clear the block below the header, when there is one
        If LastRow > Header.Row Then
            Header.Offset(1, 0).Resize(LastRow - Header.Row, Width).ClearContents
        End If

End Sub

Private Sub ClearRows( _
    ByVal HeaderName As String, _
    ByVal LastRow As Long, _
    ByVal Width As Long)
'
'------------------------------------------------------------------------------
'                           CLEAR TABLE TO GIVEN ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Clears a block from the row below its header to a given row
'
' WHY THIS EXISTS
'   A block with a fixed number of rows is cleared down to its own last row,
'   which the caller knows
'
' INPUTS
'   HeaderName
'     Workbook name of the first header cell (String)
'
'   LastRow
'     Last sheet row to clear (Long, > header row)
'
'   Width
'     Columns of the block (Long)
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the contents of the block below the header, Width columns wide, down
'   to LastRow; does nothing when LastRow is not below the header
'
' ERROR POLICY
'   Propagates ERR_LAYOUT when the name is missing
'
' DEPENDENCIES
'   NamedRange
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header  As Range    'First header cell

'------------------------------------------------------------------------------
' LOCATE HEADER
'------------------------------------------------------------------------------
    'Find the header cell
        Set Header = NamedRange(HeaderName)

'------------------------------------------------------------------------------
' CLEAR ROWS
'------------------------------------------------------------------------------
    'Clear the block below the header, when there is one
        If LastRow > Header.Row Then
            Header.Offset(1, 0).Resize(LastRow - Header.Row, Width).ClearContents
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function NamedRange( _
    ByVal RangeName As String) _
    As Range
'
'------------------------------------------------------------------------------
'                            RESOLVE WORKBOOK NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the range a workbook name refers to
'
' WHY THIS EXISTS
'   Every table and cell of the workbook is located through a workbook name, so
'   one lookup reports a missing name in one way
'
' INPUTS
'   RangeName
'     Workbook-level name (String)
'
' RETURNS
'   The range (Range)
'
' BEHAVIOR
'   Looks the name up in ThisWorkbook.Names and returns its range
'
' ERROR POLICY
'   Raises ERR_LAYOUT when the name does not exist or is not a range; the
'   lookup error is contained on the single statement that can raise it
'
' DEPENDENCIES
'   ThisWorkbook.Names; ERR_LAYOUT (CORE_Config)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target  As Range    'Result

'------------------------------------------------------------------------------
' LOOK UP NAME
'------------------------------------------------------------------------------
    'Contain the lookup error on the one statement that can raise it
        On Error Resume Next
    'Resolve the name to its range
        Set Target = ThisWorkbook.Names(RangeName).RefersToRange
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' REJECT MISSING NAME
'------------------------------------------------------------------------------
    'Reject a name that is missing or not a range
        If Target Is Nothing Then
            Err.Raise ERR_LAYOUT, "M_Main.NamedRange", _
                "Workbook name " & RangeName & " is missing or does not refer to a range."
        End If

'------------------------------------------------------------------------------
' RETURN RANGE
'------------------------------------------------------------------------------
    'Return the range
        Set NamedRange = Target

End Function

Private Function InvariantNumber( _
    ByVal Value As Variant) _
    As String
'
'------------------------------------------------------------------------------
'                                FORMAT NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Formats a number with a decimal point whatever the Windows locale, for the
'   machine-readable result line
'
' WHY THIS EXISTS
'   The result line is read by programs, and CStr would use the decimal
'   separator of the locale
'
' INPUTS
'   Value
'     A number, or Empty (read as 0) (Variant)
'
' RETURNS
'   The number as Str$ writes it, trimmed
'
' BEHAVIOR
'   Returns "0" for Empty; otherwise converts the value with CDbl and writes it
'   with Str$, which always uses a point
'
' ERROR POLICY
'   Raises the conversion error of CDbl for a value that is not a number
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FORMAT NUMBER
'------------------------------------------------------------------------------
    'Write an empty value as 0, and any other value with a decimal point
        If IsEmpty(Value) Then
            InvariantNumber = "0"
        Else
            InvariantNumber = Trim$(Str$(CDbl(Value)))
        End If

End Function

Private Function TimeText( _
    ByVal Moment As Date) _
    As String
'
'------------------------------------------------------------------------------
'                             FORMAT DATE AND TIME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a local date and time as "yyyy-mm-dd hh:mm:ss", built from its
'   parts so that no format pattern or regional setting changes it
'
' WHY THIS EXISTS
'   A format pattern depends on the regional setting, and the status line needs
'   one fixed form
'
' INPUTS
'   Moment
'     Date and time (Date)
'
' RETURNS
'   The text "yyyy-mm-dd hh:mm:ss"
'
' BEHAVIOR
'   Joins the year, month, day, hour, minute and second of the value, each of
'   the last five as two digits
'
' ERROR POLICY
'   Never raises for a valid date
'
' DEPENDENCIES
'   TwoDigits
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' BUILD TEXT
'------------------------------------------------------------------------------
    'Join the parts of the date and time
        TimeText = CStr(Year(Moment)) & "-" & TwoDigits(Month(Moment)) & "-" & _
                   TwoDigits(Day(Moment)) & " " & TwoDigits(Hour(Moment)) & ":" & _
                   TwoDigits(Minute(Moment)) & ":" & TwoDigits(Second(Moment))

End Function

Private Function TwoDigits( _
    ByVal Part As Long) _
    As String
'
'------------------------------------------------------------------------------
'                              FORMAT TWO DIGITS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a number from 0 to 99 as two digits
'
' WHY THIS EXISTS
'   TimeText needs every part of a date and time at a fixed width
'
' INPUTS
'   Part
'     Number from 0 to 99 (Long)
'
' RETURNS
'   Two characters, with a leading zero for a number below 10
'
' BEHAVIOR
'   Prefixes the number with a zero and keeps its last two characters
'
' ERROR POLICY
'   Never raises for a number from 0 to 99
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FORMAT NUMBER
'------------------------------------------------------------------------------
    'Pad with a zero and keep the last two characters
        TwoDigits = Right$("0" & CStr(Part), 2)

End Function

Private Function AppendDetail( _
    ByVal FirstDetail As String, _
    ByVal SecondDetail As String) _
    As String
'
'------------------------------------------------------------------------------
'                                JOIN MESSAGES
'------------------------------------------------------------------------------
' PURPOSE
'   Joins two restoration messages, either of which may be empty
'
' WHY THIS EXISTS
'   The restoration steps each return a message or an empty string, and the
'   caller needs the non-empty ones as one text
'
' INPUTS
'   FirstDetail
'     First message (String), may be empty
'
'   SecondDetail
'     Second message (String), may be empty
'
' RETURNS
'   The non-empty messages separated by "; "
'
' BEHAVIOR
'   Returns the other message when one is empty, and both joined with "; "
'   otherwise
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
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' JOIN MESSAGES
'------------------------------------------------------------------------------
    'Keep a message that stands alone, or join the two with a separator
        If Len(FirstDetail) = 0 Then
            AppendDetail = SecondDetail
        ElseIf Len(SecondDetail) = 0 Then
            AppendDetail = FirstDetail
        Else
            AppendDetail = FirstDetail & "; " & SecondDetail
        End If

End Function

Private Function RestoreCalculation( _
    ByVal SavedCalculation As Long) _
    As String
'
'------------------------------------------------------------------------------
'                           RESTORE CALCULATION MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Restores the calculation mode and, when it is automatic, recalculates so
'   that formulas reading the outputs (TestCatalogue, Results headline) update
'
' WHY THIS EXISTS
'   The operations switch to manual calculation, and the formulas that read the
'   outputs must be updated when the user's mode is automatic
'
' INPUTS
'   SavedCalculation
'     Mode captured before the operation (Long, an xlCalculation value)
'
' RETURNS
'   "" on success, else what failed
'
' BEHAVIOR
'   Sets Application.Calculation to the saved mode and calls
'   Application.Calculate when that mode is automatic
'
' ERROR POLICY
'   Contains its own error; never raises
'
' DEPENDENCIES
'   Application.Calculation and Application.Calculate
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the handler
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE MODE
'------------------------------------------------------------------------------
    'Set the calculation mode captured before the operation
        Application.Calculation = SavedCalculation

'------------------------------------------------------------------------------
' RECALCULATE
'------------------------------------------------------------------------------
    'Recalculate when the mode is automatic
        If SavedCalculation = xlCalculationAutomatic Then Application.Calculate

'------------------------------------------------------------------------------
' REPORT SUCCESS
'------------------------------------------------------------------------------
    'Report that the mode was restored
        RestoreCalculation = vbNullString

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return what failed
        RestoreCalculation = "Calculation not restored: " & Err.Description

End Function

Private Function RestoreFlag( _
    ByVal FlagName As String, _
    ByVal SavedValue As Boolean) _
    As String
'
'------------------------------------------------------------------------------
'                           RESTORE APPLICATION FLAG
'------------------------------------------------------------------------------
' PURPOSE
'   Restores ScreenUpdating or EnableEvents
'
' WHY THIS EXISTS
'   The two settings are restored the same way, each in its own call so that
'   one failure does not stop the other
'
' INPUTS
'   FlagName
'     "ScreenUpdating" or "EnableEvents" (String)
'
'   SavedValue
'     Value captured before the operation (Boolean)
'
' RETURNS
'   "" on success, else what failed
'
' BEHAVIOR
'   Sets Application.ScreenUpdating when FlagName is "ScreenUpdating" and
'   Application.EnableEvents for any other name
'
' ERROR POLICY
'   Contains its own error; never raises
'
' DEPENDENCIES
'   Application.ScreenUpdating and Application.EnableEvents
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route any error to the handler
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE FLAG
'------------------------------------------------------------------------------
    'Set the flag that the name selects to the captured value
        Select Case FlagName
            Case "ScreenUpdating"
                Application.ScreenUpdating = SavedValue
            Case Else
                Application.EnableEvents = SavedValue
        End Select

'------------------------------------------------------------------------------
' REPORT SUCCESS
'------------------------------------------------------------------------------
    'Report that the flag was restored
        RestoreFlag = vbNullString

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return what failed
        RestoreFlag = FlagName & " not restored: " & Err.Description

End Function
