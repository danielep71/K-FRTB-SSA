Attribute VB_Name = "M_Main"
'==============================================================================
' MODULE: M_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Host layer of K-FRTB-SSA: the command macros, reading the input sheets and
'   Params into the engine's input vector, calling CORE_Engine.Calculate and
'   writing every output table back, with Excel application state captured
'   and restored around each operation.
'
' PUBLIC SURFACE
'   RunSSA, RunSSA_Silent, ClearOutputs, ClearOutputs_Silent (command
'   macros; not part of the supported worksheet-function API).
'
' DEPENDENCIES
'   CORE_Config, CORE_Engine, CORE_Util. Locates every table through the
'   workbook names listed in CORE_Config (TABLE LOCATION).
'
' STATE OWNERSHIP
'   Reads the input sheets and Params; clears and writes the output sheets
'   (IRCalc, EQCalc, COCalc, FXCalc, IR_Ladder, IR_Zones, Buckets, Results,
'   Checks). Changes Calculation, ScreenUpdating and EnableEvents for the
'   duration of an operation and restores them. mSilent holds the silent
'   flag of the running operation.
'
' ERROR POLICY
'   The command macros report a failure in a message box (silent variants:
'   re-raise it). The primary error is preserved through cleanup; a
'   cleanup-only failure raises ERR_CLEANUP_FAILED.
'
' COMPATIBILITY
'   Excel for Windows, VBA7; no references beyond the defaults.
'
' UPDATED
'   2026-11-14
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations.
    Option Explicit

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const OP_RUN         As Long = 1     'Operation: calculate
        Private Const OP_CLEAR       As Long = 2     'Operation: clear outputs
        Private Const CHECKS_COUNT_COLUMN  As Long = 4    'Count, from the rule column (F)
        Private Const CHECKS_RESULT_COLUMN As Long = 5    'Result, from the rule column (G)
        Private Const NOT_RUN        As String = "NOT RUN"    'Result of a rule not evaluated

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mSilent              As Boolean      'True while a _Silent macro runs


'------------------------------------------------------------------------------
'                                COMMAND MACROS
'------------------------------------------------------------------------------

Public Sub RunSSA()
'
'==============================================================================
'                                    RunSSA
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate the MAR40 simplified standardised approach and write every
'   output sheet, then show Results (or Checks when the run stopped) and a
'   summary message.
'
' USAGE
'   Alt+F8 > RunSSA, or a button assigned to it.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Summary     As String    'Result line of the run

        On Error GoTo Failed
        mSilent = False
        Summary = ExecuteOperation(OP_RUN)
        MsgBox Replace(Summary, "; ", vbLf), vbInformation, PRODUCT_NAME
        Exit Sub

Failed:
        MsgBox "The run failed." & vbLf & Err.Description & vbLf & "(" & Err.Source & ")", _
               vbCritical, PRODUCT_NAME

End Sub


Public Function RunSSA_Silent() As String
'
'==============================================================================
'                                 RunSSA_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Run RunSSA unattended: no message box, no sheet activation; failures
'   are raised to the caller.
'
' RETURNS
'   "RESULT=OK|INCOMPLETE|STOPPED; errors=n; warnings=n; positions_used=n;
'   requirement=x; rwa=x; cleanup=PASS", numbers with a decimal point.
'
' STATE OWNERSHIP
'   Sets mSilent for the run and clears it afterwards, also on error.
'
' ERROR POLICY
'   Re-raises the operation's error after mSilent is cleared.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim ErrNumber      As Long      'Primary error number; 0 for none
    Dim ErrSource      As String    'Primary error source
    Dim ErrDescription As String    'Primary error description

        On Error GoTo Failed
        mSilent = True
        RunSSA_Silent = ExecuteOperation(OP_RUN)
        mSilent = False
        Exit Function

Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        mSilent = False
        Err.Raise ErrNumber, ErrSource, ErrDescription

End Function


Public Sub ClearOutputs()
'
'==============================================================================
'                                 ClearOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Clear every output written by RunSSA; inputs and Params are untouched.
'
' USAGE
'   Alt+F8 > ClearOutputs.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        On Error GoTo Failed
        mSilent = False
        ExecuteOperation OP_CLEAR
        Exit Sub

Failed:
        MsgBox "Clearing failed." & vbLf & Err.Description, vbCritical, PRODUCT_NAME

End Sub


Public Function ClearOutputs_Silent() As String
'
'==============================================================================
'                              ClearOutputs_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   Run ClearOutputs unattended.
'
' RETURNS
'   "RESULT=CLEARED; cleanup=PASS".
'
' ERROR POLICY
'   Re-raises the operation's error after mSilent is cleared.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim ErrNumber      As Long      'Primary error number; 0 for none
    Dim ErrSource      As String    'Primary error source
    Dim ErrDescription As String    'Primary error description

        On Error GoTo Failed
        mSilent = True
        ClearOutputs_Silent = ExecuteOperation(OP_CLEAR)
        mSilent = False
        Exit Function

Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        mSilent = False
        Err.Raise ErrNumber, ErrSource, ErrDescription

End Function


'------------------------------------------------------------------------------
'                                   OPERATION
'------------------------------------------------------------------------------

Private Function ExecuteOperation( _
    ByVal Operation As Long) _
    As String
'
'==============================================================================
'                               ExecuteOperation
'------------------------------------------------------------------------------
' PURPOSE
'   Run one operation with Excel state captured, changed and restored.
'
' INPUTS
'   Operation   OP_RUN or OP_CLEAR.
'
' RETURNS
'   The operation's result line, with "; cleanup=PASS" appended.
'
' STATE OWNERSHIP
'   Captures Calculation, ScreenUpdating and EnableEvents; sets manual
'   calculation, no screen updating and no events; restores all three on
'   success and on error, then recalculates when calculation was automatic.
'
' ERROR POLICY
'   Preserves the primary error through cleanup and re-raises it; raises
'   ERR_CLEANUP_FAILED when only the restoration failed.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim SavedCalculation As Long        'Calculation mode before the operation
    Dim SavedScreen      As Boolean     'ScreenUpdating before the operation
    Dim SavedEvents      As Boolean     'EnableEvents before the operation
    Dim StateCaptured    As Boolean     'True once all three were read
    Dim CleanupDetails   As String      'Restoration failures, "" for none
    Dim ResultLine       As String      'Result of the operation
    Dim ErrNumber        As Long        'Primary error number; 0 for none
    Dim ErrSource        As String      'Primary error source
    Dim ErrDescription   As String      'Primary error description

        On Error GoTo Failed
        SavedCalculation = Application.Calculation
        SavedScreen = Application.ScreenUpdating
        SavedEvents = Application.EnableEvents
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

        Select Case Operation
            Case OP_RUN
                ResultLine = RunCalculation()
            Case OP_CLEAR
                ClearAllOutputs
                ResultLine = "RESULT=CLEARED"
        End Select

CleanUp:
        On Error GoTo 0
        If StateCaptured Then
            CleanupDetails = AppendDetail(RestoreCalculation(SavedCalculation), _
                             AppendDetail(RestoreFlag("ScreenUpdating", SavedScreen), _
                                          RestoreFlag("EnableEvents", SavedEvents)))
        End If
        If ErrNumber <> 0 Then Err.Raise ErrNumber, ErrSource, ErrDescription
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "M_Main.ExecuteOperation", CleanupDetails
        End If
        ExecuteOperation = ResultLine & "; cleanup=PASS"
        Exit Function

Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume CleanUp

End Function


Private Function RunCalculation() As String
'
'==============================================================================
'                                RunCalculation
'------------------------------------------------------------------------------
' PURPOSE
'   Read the inputs, calculate, write the outputs and activate the sheet to
'   review.
'
' RETURNS
'   "RESULT=...; errors=n; warnings=n; positions_used=n; requirement=x;
'   rwa=x".
'
' ERROR POLICY
'   Propagates ERR_LAYOUT (a workbook name is missing) and any engine error.
'
' STATE OWNERSHIP
'   Clears and writes the output sheets; activates Results or Checks unless
'   silent.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Inputs()    As Variant   'Engine input vector
    Dim Outputs()   As Variant   'Engine output vector
    Dim Outcome     As String    'OK, INCOMPLETE or STOPPED

        ReadInputs Inputs
        Calculate Inputs, Outputs
        ClearAllOutputs
        WriteOutputs Outputs

        If Outputs(OUT_STOPPED) Then
            Outcome = "STOPPED"
        ElseIf Outputs(OUT_ERRORS) > 0 Then
            Outcome = "INCOMPLETE"
        Else
            Outcome = "OK"
        End If
        If Not mSilent Then
            If Outputs(OUT_STOPPED) Then
                ThisWorkbook.Worksheets(SH_CHECKS).Activate
            Else
                ThisWorkbook.Worksheets(SH_RESULTS).Activate
            End If
        End If
        RunCalculation = "RESULT=" & Outcome & "; errors=" & Outputs(OUT_ERRORS) & _
                         "; warnings=" & InvariantNumber(Outputs(OUT_WARNINGS)) & _
                         "; positions_used=" & InvariantNumber(Outputs(OUT_USED)) & _
                         "; requirement=" & InvariantNumber(Outputs(OUT_REQUIREMENT)) & _
                         "; rwa=" & InvariantNumber(Outputs(OUT_RWA))

End Function


'------------------------------------------------------------------------------
'                                    READING
'------------------------------------------------------------------------------

Private Sub ReadInputs( _
    ByRef Inputs() As Variant)
'
'==============================================================================
'                                  ReadInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Fill the engine input vector from Params, OptionCharges and the four
'   position sheets.
'
' RETURNS
'   Inputs(1 To IN_COUNT), as CORE_Config describes.
'
' ERROR POLICY
'   Raises ERR_LAYOUT when a workbook name is missing.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Numeric(1 To PI_COUNT)  As Variant      'Numeric parameters
    Dim Options(1 To RC_COUNT)  As Variant      'Option charges
    Dim Ladder()       As Variant   'Ladder matrix
    Dim Labels()       As String    'Ladder labels
    Dim Column         As Variant   'One named column as read
    Dim FirstRow       As Long      'First data row of a table
    Dim b              As Long      'Band
    Dim i              As Long      'Parameter or class

        ReDim Inputs(1 To IN_COUNT)
        For i = 1 To PI_COUNT
            Numeric(i) = NamedRange(ListItem(PRM_NUMERIC_CODES, i)).Value
        Next i
        Inputs(IN_NUMERIC) = Numeric
        Inputs(IN_ASOF) = NamedRange(PRM_ASOF).Value
        Inputs(IN_REPCCY) = NamedRange(PRM_REPCCY).Value
        Inputs(IN_GOLD) = NamedRange(PRM_GOLD).Value

        ReDim Ladder(1 To BAND_COUNT, 1 To LD_WIDTH)
        ReDim Labels(1 To BAND_COUNT, 1 To 2)
        Column = ColumnValues(NM_LADDER_ZONE, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Ladder(b, LD_ZONE) = Column(b, 1)
        Next b
        Column = ColumnValues(NM_LADDER_UP_HIGH, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Ladder(b, LD_UPPER_HIGH) = Column(b, 1)
        Next b
        Column = ColumnValues(NM_LADDER_UP_LOW, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Ladder(b, LD_UPPER_LOW) = Column(b, 1)
        Next b
        Column = ColumnValues(NM_LADDER_RW, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Ladder(b, LD_WEIGHT) = Column(b, 1)
        Next b
        Column = ColumnValues(NM_LADDER_LBL_HIGH, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Labels(b, 1) = PlainText(Column(b, 1))
        Next b
        Column = ColumnValues(NM_LADDER_LBL_LOW, BAND_COUNT)
        For b = 1 To BAND_COUNT
            Labels(b, 2) = PlainText(Column(b, 1))
        Next b
        Inputs(IN_LADDER) = Ladder
        Inputs(IN_LADDER_LABELS) = Labels

        Inputs(IN_SR_TABLE) = PairedColumns(NM_SR_KEY, NM_SR_CHARGE)
        Inputs(IN_FX_TABLE) = PairedColumns(NM_FX_CCY, NM_FX_RATE)
        For i = 1 To RC_COUNT
            Options(i) = NamedRange(NM_OPTIONS_PREFIX & ListItem(RC_CODES, i)).Value
        Next i
        Inputs(IN_OPTIONS) = Options

        Inputs(IN_IR_ROWS) = TableValues(NM_IR_POS_HEADER, IR_INPUT_WIDTH, FirstRow)
        Inputs(IN_IR_FIRST_ROW) = FirstRow
        Inputs(IN_EQ_ROWS) = TableValues(NM_EQ_POS_HEADER, EQ_INPUT_WIDTH, FirstRow)
        Inputs(IN_EQ_FIRST_ROW) = FirstRow
        Inputs(IN_CO_ROWS) = TableValues(NM_CO_POS_HEADER, CO_INPUT_WIDTH, FirstRow)
        Inputs(IN_CO_FIRST_ROW) = FirstRow
        Inputs(IN_FX_ROWS) = TableValues(NM_FX_POS_HEADER, FX_INPUT_WIDTH, FirstRow)
        Inputs(IN_FX_FIRST_ROW) = FirstRow

        Inputs(IN_EQ_CAPACITY) = NamedRange(NM_EQ_GENERAL).Row - NamedRange(NM_BK_EQ_HEADER).Row - 1
        Inputs(IN_CO_CAPACITY) = NamedRange(NM_CO_DIRECTIONAL).Row - NamedRange(NM_BK_CO_HEADER).Row - 1

End Sub


Private Function TableValues( _
    ByVal HeaderName As String, _
    ByVal Width As Long, _
    ByRef FirstRow As Long) _
    As Variant
'
'==============================================================================
'                                  TableValues
'------------------------------------------------------------------------------
' PURPOSE
'   Read an input table from the row below its header to the last row with
'   a value in its first column.
'
' INPUTS
'   HeaderName   Workbook name of the table's first header cell.
'   Width        Columns to read (>= 2).
'
' RETURNS
'   Variant (1 To n, 1 To Width), n >= 1 (one blank row for an empty
'   table); FirstRow = sheet row of the first data row.
'
' ERROR POLICY
'   Raises ERR_LAYOUT when the name is missing.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Header      As Range     'First header cell
    Dim Sheet       As Worksheet 'Sheet of the table
    Dim LastRow     As Long      'Last row with a PositionID

        Set Header = NamedRange(HeaderName)
        Set Sheet = Header.Worksheet
        FirstRow = Header.Row + 1
        LastRow = Sheet.Cells(Sheet.Rows.Count, Header.Column).End(xlUp).Row
        If LastRow < FirstRow Then LastRow = FirstRow
        TableValues = Sheet.Range(Sheet.Cells(FirstRow, Header.Column), _
                                  Sheet.Cells(LastRow, Header.Column + Width - 1)).Value

End Function


Private Function ColumnValues( _
    ByVal RangeName As String, _
    ByVal ExpectedCount As Long) _
    As Variant
'
'==============================================================================
'                                 ColumnValues
'------------------------------------------------------------------------------
' PURPOSE
'   Read a named single-column range of a known number of cells.
'
' INPUTS
'   RangeName       Workbook name of the column.
'   ExpectedCount   Cells the range must have (>= 2).
'
' RETURNS
'   Variant (1 To ExpectedCount, 1 To 1).
'
' ERROR POLICY
'   Raises ERR_LAYOUT when the name is missing or has another size.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Target      As Range     'Named range

        Set Target = NamedRange(RangeName)
        If Target.Rows.Count <> ExpectedCount Or Target.Columns.Count <> 1 Then
            Err.Raise ERR_LAYOUT, "M_Main.ColumnValues", _
                "Workbook name " & RangeName & " must refer to " & ExpectedCount & " cells in one column."
        End If
        ColumnValues = Target.Value

End Function


Private Function PairedColumns( _
    ByVal FirstName As String, _
    ByVal SecondName As String) _
    As Variant
'
'==============================================================================
'                                 PairedColumns
'------------------------------------------------------------------------------
' PURPOSE
'   Read two named columns of equal length as one two-column table.
'
' INPUTS
'   FirstName, SecondName   Workbook names of the two columns.
'
' RETURNS
'   Variant (1 To n, 1 To 2).
'
' ERROR POLICY
'   Raises ERR_LAYOUT when a name is missing or the lengths differ.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim FirstValues  As Variant  'First column values
    Dim SecondValues As Variant  'Second column values
    Dim Table()     As Variant   'Result
    Dim RowCount    As Long      'Rows
    Dim r           As Long      'Row

        RowCount = NamedRange(FirstName).Rows.Count
        FirstValues = ColumnValues(FirstName, RowCount)
        SecondValues = ColumnValues(SecondName, RowCount)
        ReDim Table(1 To RowCount, 1 To 2)
        For r = 1 To RowCount
            Table(r, 1) = FirstValues(r, 1)
            Table(r, 2) = SecondValues(r, 1)
        Next r
        PairedColumns = Table

End Function


'------------------------------------------------------------------------------
'                                    WRITING
'------------------------------------------------------------------------------

Private Sub WriteOutputs( _
    ByRef Outputs() As Variant)
'
'==============================================================================
'                                 WriteOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Write the engine outputs to their sheets; the sheets were cleared
'   before.
'
' INPUTS
'   Outputs   Vector (1 To OUT_COUNT) from CORE_Engine.Calculate.
'
' STATE OWNERSHIP
'   Writes IRCalc, EQCalc, COCalc, FXCalc, IR_Ladder, IR_Zones, Buckets,
'   Results and Checks.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Header      As Range     'Header cell of a table
    Dim MarketTable As Variant   'Buckets equity block
    Dim CommodityTable As Variant    'Buckets commodity block
    Dim SummaryTable As Variant  'FX ONOP and charge
    Dim Capacity    As Long      'Rows of a Buckets block
    Dim Shown       As Long      'Rows written to a Buckets block

        NamedRange(NM_RESULTS_STATUS).Value = Outputs(OUT_STATUS)
        WriteRules Outputs(OUT_RULES)
        NamedRange(NM_CHECKS_ERRORS).Value = Outputs(OUT_ERRORS)
        NamedRange(NM_CHECKS_WARNINGS).Value = Outputs(OUT_WARNINGS)
        WriteBlock NamedRange(NM_FINDINGS_HEADER).Offset(1, 0), Outputs(OUT_FINDINGS)
        If Outputs(OUT_STOPPED) Then Exit Sub

        WriteBlock NamedRange(NM_IR_CALC_HEADER).Offset(1, 0), Outputs(OUT_IR_CALC)
        WriteBlock NamedRange(NM_EQ_CALC_HEADER).Offset(1, 0), Outputs(OUT_EQ_CALC)
        WriteBlock NamedRange(NM_CO_CALC_HEADER).Offset(1, 0), Outputs(OUT_CO_CALC)
        WriteBlock NamedRange(NM_FX_CALC_HEADER).Offset(1, 0), Outputs(OUT_FX_CALC)
        WriteBlock NamedRange(NM_LADDER_HEADER).Offset(1, 0), Outputs(OUT_LADDER)
        WriteBlock NamedRange(NM_ZONES_HEADER).Offset(1, 0), Outputs(OUT_ZONES)
        WriteBlock NamedRange(NM_RESULTS_HEADER).Offset(1, RESULTS_COLUMN - 1), Outputs(OUT_RESULTS)

        'Buckets blocks: as many rows as fit, then the total row at its name.
        Set Header = NamedRange(NM_BK_EQ_HEADER)
        Capacity = NamedRange(NM_EQ_GENERAL).Row - Header.Row - 1
        MarketTable = Outputs(OUT_MARKETS)
        Shown = WriteBuckets(Header, MarketTable, Capacity)
        Set Header = NamedRange(NM_BK_CO_HEADER)
        Capacity = NamedRange(NM_CO_DIRECTIONAL).Row - Header.Row - 1
        CommodityTable = Outputs(OUT_COMMODITIES)
        Shown = WriteBuckets(Header, CommodityTable, Capacity)
        WriteBlock NamedRange(NM_BK_FX_HEADER).Offset(1, 0), Outputs(OUT_CURRENCIES)
        SummaryTable = Outputs(OUT_FX_SUMMARY)
        NamedRange(NM_FX_ONOP).Value = SummaryTable(1, 1)
        NamedRange(NM_FX_CHARGE).Value = SummaryTable(2, 1)

End Sub


Private Function WriteBuckets( _
    ByVal Header As Range, _
    ByRef Table As Variant, _
    ByVal Capacity As Long) _
    As Long
'
'==============================================================================
'                                 WriteBuckets
'------------------------------------------------------------------------------
' PURPOSE
'   Write an equity or commodity block of Buckets: up to Capacity bucket
'   rows below the header, and the total row (the table's last row) on the
'   row after the capacity. Totals include every bucket; rule C10 / C11
'   reports buckets not shown.
'
' INPUTS
'   Header     First header cell of the block.
'   Table      (1 To k + 1, 1 To width): k buckets and the total row.
'   Capacity   Bucket rows of the block (>= 1).
'
' RETURNS
'   Bucket rows written.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Shown       As Long      'Bucket rows written
    Dim Part()      As Variant   'Rows written
    Dim Width       As Long      'Columns
    Dim r           As Long      'Row
    Dim c           As Long      'Column

        Width = UBound(Table, 2)
        Shown = UBound(Table, 1) - 1
        If Shown > Capacity Then Shown = Capacity
        If Shown > 0 Then
            ReDim Part(1 To Shown, 1 To Width)
            For r = 1 To Shown
                For c = 1 To Width
                    Part(r, c) = Table(r, c)
                Next c
            Next r
            Header.Offset(1, 0).Resize(Shown, Width).Value = Part
        End If
        ReDim Part(1 To 1, 1 To Width)
        For c = 1 To Width
            Part(1, c) = Table(UBound(Table, 1), c)
        Next c
        Header.Offset(Capacity + 1, 0).Resize(1, Width).Value = Part
        WriteBuckets = Shown

End Function


Private Sub WriteRules( _
    ByVal Rules As Variant)
'
'==============================================================================
'                                  WriteRules
'------------------------------------------------------------------------------
' PURPOSE
'   Write the count and result of each Checks rule on the row whose Rule
'   cell holds its identifier.
'
' INPUTS
'   Rules   (1 To RULE_COUNT, 1 To 2): identifier, count or Empty.
'
' ERROR POLICY
'   Raises ERR_LAYOUT when a rule identifier is not found under the header.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Header      As Range     'Rule header cell
    Dim Found       As Boolean   'Rule row found
    Dim r           As Long      'Rule
    Dim RowOffset   As Long      'Row offset under the header

        Set Header = NamedRange(NM_CHECKS_HEADER)
        For r = 1 To UBound(Rules, 1)
            Found = False
            For RowOffset = 1 To RULE_COUNT + 5
                If PlainText(Header.Offset(RowOffset, 0).Value) = Rules(r, 1) Then
                    Found = True
                    Exit For
                End If
            Next RowOffset
            If Not Found Then
                Err.Raise ERR_LAYOUT, "M_Main.WriteRules", _
                    "Rule " & Rules(r, 1) & " not found on " & SH_CHECKS & "."
            End If
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
'==============================================================================
'                                  WriteBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Write a two-dimensional array with its top-left cell at TopLeft.
'
' INPUTS
'   TopLeft   First cell to write.
'   Table     Variant (1 To n, 1 To m), or Empty to write nothing.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        If IsEmpty(Table) Then Exit Sub
        TopLeft.Resize(UBound(Table, 1), UBound(Table, 2)).Value = Table

End Sub


'------------------------------------------------------------------------------
'                                   CLEARING
'------------------------------------------------------------------------------

Private Sub ClearAllOutputs()
'
'==============================================================================
'                                ClearAllOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the contents (not the formats) of every output area.
'
' STATE OWNERSHIP
'   Clears the output sheets; never touches inputs, Params or formulas
'   outside the output areas.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Header      As Range     'Header cell of a table
    Dim r           As Long      'Rule row

        ClearToEnd NM_IR_CALC_HEADER, IR_CALC_WIDTH
        ClearToEnd NM_EQ_CALC_HEADER, EQ_CALC_WIDTH
        ClearToEnd NM_CO_CALC_HEADER, CO_CALC_WIDTH
        ClearToEnd NM_FX_CALC_HEADER, FX_CALC_WIDTH
        ClearToEnd NM_LADDER_HEADER, LADDER_WIDTH
        ClearToEnd NM_FINDINGS_HEADER, FINDINGS_WIDTH
        ClearRows NM_ZONES_HEADER, NamedRange(NM_IR_GMR_TOTAL).Row, ZONES_WIDTH
        ClearRows NM_BK_EQ_HEADER, NamedRange(NM_EQ_GENERAL).Row, BK_EQ_WIDTH
        ClearRows NM_BK_CO_HEADER, NamedRange(NM_CO_DIRECTIONAL).Row, BK_CO_WIDTH
        ClearRows NM_BK_FX_HEADER, NamedRange(NM_BK_FX_HEADER).Row + _
                  NamedRange(NM_FX_CCY).Rows.Count + 1, BK_FX_WIDTH
        NamedRange(NM_FX_ONOP).ClearContents
        NamedRange(NM_FX_CHARGE).ClearContents

        'Labels of the total rows sit in the cleared areas: put them back.
        NamedRange(NM_IR_GMR_TOTAL).Worksheet.Cells(NamedRange(NM_IR_GMR_TOTAL).Row, _
            NamedRange(NM_ZONES_HEADER).Column).Value = LBL_TOTAL_ZONES
        NamedRange(NM_BK_EQ_HEADER).Worksheet.Cells(NamedRange(NM_EQ_GENERAL).Row, _
            NamedRange(NM_BK_EQ_HEADER).Column + 1).Value = LBL_TOTAL_EQ
        NamedRange(NM_BK_CO_HEADER).Worksheet.Cells(NamedRange(NM_CO_DIRECTIONAL).Row, _
            NamedRange(NM_BK_CO_HEADER).Column + 1).Value = LBL_TOTAL_CO
        NamedRange(NM_BK_FX_HEADER).Offset(NamedRange(NM_FX_CCY).Rows.Count + 1, 1).Value = LBL_TOTAL_FX

        Set Header = NamedRange(NM_RESULTS_HEADER)
        Header.Offset(1, RESULTS_COLUMN - 1).Resize(RC_COUNT + 1, RESULTS_WIDTH).ClearContents
        NamedRange(NM_RESULTS_STATUS).Value = "Not calculated: run RunSSA."

        Set Header = NamedRange(NM_CHECKS_HEADER)
        For r = 1 To RULE_COUNT
            Header.Offset(r, CHECKS_COUNT_COLUMN - 1).Resize(1, 2).ClearContents
        Next r
        NamedRange(NM_CHECKS_ERRORS).ClearContents
        NamedRange(NM_CHECKS_WARNINGS).ClearContents

End Sub


Private Sub ClearToEnd( _
    ByVal HeaderName As String, _
    ByVal Width As Long)
'
'==============================================================================
'                                  ClearToEnd
'------------------------------------------------------------------------------
' PURPOSE
'   Clear a table from the row below its header to the last used row of
'   its sheet.
'
' INPUTS
'   HeaderName   Workbook name of the first header cell.
'   Width        Columns of the table.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Header      As Range     'First header cell
    Dim LastRow     As Long      'Last used row of the sheet

        Set Header = NamedRange(HeaderName)
        With Header.Worksheet.UsedRange
            LastRow = .Row + .Rows.Count - 1
        End With
        If LastRow > Header.Row Then
            Header.Offset(1, 0).Resize(LastRow - Header.Row, Width).ClearContents
        End If

End Sub


Private Sub ClearRows( _
    ByVal HeaderName As String, _
    ByVal LastRow As Long, _
    ByVal Width As Long)
'
'==============================================================================
'                                   ClearRows
'------------------------------------------------------------------------------
' PURPOSE
'   Clear a block from the row below its header to a given row.
'
' INPUTS
'   HeaderName   Workbook name of the first header cell.
'   LastRow      Last sheet row to clear (> header row).
'   Width        Columns of the block.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Header      As Range     'First header cell

        Set Header = NamedRange(HeaderName)
        If LastRow > Header.Row Then
            Header.Offset(1, 0).Resize(LastRow - Header.Row, Width).ClearContents
        End If

End Sub


'------------------------------------------------------------------------------
'                                    HELPERS
'------------------------------------------------------------------------------

Private Function NamedRange( _
    ByVal RangeName As String) _
    As Range
'
'==============================================================================
'                                  NamedRange
'------------------------------------------------------------------------------
' PURPOSE
'   Return the range a workbook name refers to.
'
' INPUTS
'   RangeName   Workbook-level name.
'
' RETURNS
'   The range.
'
' ERROR POLICY
'   Raises ERR_LAYOUT when the name does not exist or is not a range; the
'   lookup error is contained on the single statement that can raise it.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Target      As Range     'Result

        On Error Resume Next
        Set Target = ThisWorkbook.Names(RangeName).RefersToRange
        On Error GoTo 0
        If Target Is Nothing Then
            Err.Raise ERR_LAYOUT, "M_Main.NamedRange", _
                "Workbook name " & RangeName & " is missing or does not refer to a range."
        End If
        Set NamedRange = Target

End Function


Private Function InvariantNumber( _
    ByVal Value As Variant) _
    As String
'
'==============================================================================
'                                InvariantNumber
'------------------------------------------------------------------------------
' PURPOSE
'   Format a number with a decimal point whatever the Windows locale, for
'   the machine-readable result line.
'
' INPUTS
'   Value   A number, or Empty (read as 0).
'
' RETURNS
'   The number as Str$ writes it, trimmed.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        If IsEmpty(Value) Then
            InvariantNumber = "0"
        Else
            InvariantNumber = Trim$(Str$(CDbl(Value)))
        End If

End Function


Private Function AppendDetail( _
    ByVal FirstDetail As String, _
    ByVal SecondDetail As String) _
    As String
'
'==============================================================================
'                                 AppendDetail
'------------------------------------------------------------------------------
' PURPOSE
'   Join two restoration messages, either of which may be empty.
'
' RETURNS
'   The non-empty messages separated by "; ".
'
' UPDATED
'   2026-11-14
'==============================================================================
'
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
'==============================================================================
'                              RestoreCalculation
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the calculation mode and, when it is automatic, recalculate so
'   that formulas reading the outputs (TestCatalogue, Results headline)
'   update.
'
' INPUTS
'   SavedCalculation   Mode captured before the operation.
'
' RETURNS
'   "" on success, else what failed.
'
' ERROR POLICY
'   Contains its own error; never raises.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        On Error GoTo Failed
        Application.Calculation = SavedCalculation
        If SavedCalculation = xlCalculationAutomatic Then Application.Calculate
        RestoreCalculation = vbNullString
        Exit Function

Failed:
        RestoreCalculation = "Calculation not restored: " & Err.Description

End Function


Private Function RestoreFlag( _
    ByVal FlagName As String, _
    ByVal SavedValue As Boolean) _
    As String
'
'==============================================================================
'                                  RestoreFlag
'------------------------------------------------------------------------------
' PURPOSE
'   Restore ScreenUpdating or EnableEvents.
'
' INPUTS
'   FlagName     "ScreenUpdating" or "EnableEvents".
'   SavedValue   Value captured before the operation.
'
' RETURNS
'   "" on success, else what failed.
'
' ERROR POLICY
'   Contains its own error; never raises.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        On Error GoTo Failed
        Select Case FlagName
            Case "ScreenUpdating"
                Application.ScreenUpdating = SavedValue
            Case Else
                Application.EnableEvents = SavedValue
        End Select
        RestoreFlag = vbNullString
        Exit Function

Failed:
        RestoreFlag = FlagName & " not restored: " & Err.Description

End Function
