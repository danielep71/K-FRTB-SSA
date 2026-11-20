Attribute VB_Name = "CORE_Engine"
'==============================================================================
' MODULE: CORE_Engine
'------------------------------------------------------------------------------
' PURPOSE
'   Run the whole MAR40 simplified standardised approach from raw cell
'   values to every output table: validate Params, stop when a parameter is
'   unusable or a position sheet holds more rows than the workbook supports,
'   otherwise calculate the four risk classes, scale and sum them
'   into the own funds requirement and RWA, and assemble the Checks rules,
'   the findings and the Results status line.
'
' PUBLIC SURFACE
'   None outside this VBA project (Option Private Module).
'     Calculate, PositionVolumeErrors.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util, CORE_Parameters, CORE_InterestRate, CORE_Equity,
'   CORE_Commodity, CORE_ForeignExchange.
'
' STATE OWNERSHIP
'   None. The procedure depends only on its arguments; it never touches the
'   Excel object model, so it can be tested without a workbook.
'
' ERROR POLICY
'   Cell content never raises: unusable values become findings. Arrays of
'   the wrong shape raise ERR_ARGUMENT from the module that detects them,
'   propagated unchanged.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' REFERENCE
'   BCBS MAR40.1-40.2 (aggregation and scaling), MAR40 throughout.
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
    'Require explicit declarations; keep the procedure inside this project.
    Option Explicit
    Option Private Module


'------------------------------------------------------------------------------
'                                  CALCULATION
'------------------------------------------------------------------------------

Public Sub Calculate( _
    ByRef Inputs() As Variant, _
    ByRef Outputs() As Variant)
'
'==============================================================================
'                                   Calculate
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate the own funds requirement and every intermediate table.
'
' INPUTS
'   Inputs   Vector (1 To IN_COUNT), filled as CORE_Config describes for
'            each IN_ index: raw Params values and tables, the four input
'            tables as cell values, the sheet row of each table's first data
'            row and the market and commodity capacity of Buckets.
'
' RETURNS
'   Outputs  Vector (1 To OUT_COUNT), as CORE_Config describes for each
'            OUT_ index. When OUTPUTS(OUT_STOPPED) is True, only OUT_RULES,
'            OUT_FINDINGS, OUT_STATUS and OUT_ERRORS are set.
'
' REFERENCE
'   MAR40.2: requirement = 1.30 IR + 3.50 EQ + 1.90 CO + 1.20 FX, each class
'   including its non-delta option charge; MAR40.1: RWA = 12.5 x it.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Parameters()   As Double     'Validated parameter vector
    Dim AsOfDate       As Double     'Reporting date serial
    Dim ReportingCcy   As String     'Reporting currency
    Dim GoldCode       As String     'Gold code
    Dim Regime         As String     'REGIME_BASEL or REGIME_CRR
    Dim TransitionEnd  As Double     'Last date of the transitional multiplier
    Dim Ladder()       As Double     'Validated ladder
    Dim LadderLabels() As String     'Band labels
    Dim SrKeys()       As String     'Specific-risk keys
    Dim SrCharges()    As Double     'Specific-risk charges
    Dim FxCodes()      As String     'FX codes
    Dim FxRates()      As Double     'FX rates
    Dim OptionCharges() As Double    'Non-delta option charges per class
    Dim Findings()     As String     'Findings (1 To 4, 1 To n)
    Dim FindingCount   As Long       'Findings in use
    Dim ParameterErrors As Long      'BuildParameters errors
    Dim VolumeErrors   As Long       'Position sheets over MAX_POSITION_ROWS
    Dim TableErrors    As Long       'Ladder, specific-risk and option errors
    Dim ReportingErrors As Long      'C05
    Dim DuplicateErrors As Long      'C06
    Dim RateErrors     As Long       'C07
    Dim RuleCounts(1 To RULE_COUNT) As Variant    'Count per rule, Empty when not evaluated
    Dim IrTotals()     As Double     'Interest rate totals
    Dim EqTotals()     As Double     'Equity totals
    Dim CoTotals()     As Double     'Commodity totals
    Dim FxTotals()     As Double     'FX totals
    Dim IrCalc()       As Variant    'IRCalc table
    Dim LadderTable()  As Variant    'IR_Ladder table
    Dim ZoneTable()    As Variant    'IR_Zones table
    Dim EqCalc()       As Variant    'EQCalc table
    Dim MarketTable()  As Variant    'Buckets equity block
    Dim CoCalc()       As Variant    'COCalc table
    Dim CommodityTable() As Variant  'Buckets commodity block
    Dim FxCalc()       As Variant    'FXCalc table
    Dim CurrencyTable() As Variant   'Buckets FX block
    Dim ResultTable    As Variant    'Results table
    Dim Summary(1 To 2, 1 To 1) As Variant    'ONOP and FX charge
    Dim Errors         As Long       'Error count
    Dim Warnings       As Long       'Warning count
    Dim RawNumeric()   As Variant    'Numeric parameters as read
    Dim RawLadder()    As Variant    'Ladder as read
    Dim RawSrTable()   As Variant    'Specific-risk table as read
    Dim RawOptions()   As Variant    'Option charges as read
    Dim RawFxTable()   As Variant    'FX table as read
    Dim IrRows()       As Variant    'IRPositions cells
    Dim EqRows()       As Variant    'EQPositions cells
    Dim CoRows()       As Variant    'COPositions cells
    Dim FxRows()       As Variant    'FXPositions cells
    Dim i              As Long       'Rule or class index

        ReDim Outputs(1 To OUT_COUNT)
        Outputs(OUT_STOPPED) = False

        '--- Params --------------------------------------------------------------
        RawNumeric = Inputs(IN_NUMERIC)
        ParameterErrors = BuildParameters(RawNumeric, Inputs(IN_ASOF), Inputs(IN_REPCCY), _
                                          Inputs(IN_GOLD), Parameters, AsOfDate, ReportingCcy, _
                                          GoldCode, Findings, FindingCount)
        ParameterErrors = ParameterErrors + BuildRegime(Inputs(IN_REGIME), Inputs(IN_TRANSITION_END), _
                                                        Regime, TransitionEnd, Findings, FindingCount)
        RawLadder = Inputs(IN_LADDER)
        TableErrors = BuildLadder(RawLadder, Ladder, Findings, FindingCount)
        RawSrTable = Inputs(IN_SR_TABLE)
        TableErrors = TableErrors + BuildSpecificRiskTable(RawSrTable, SrKeys, SrCharges, _
                                                           Findings, FindingCount)
        RawOptions = Inputs(IN_OPTIONS)
        TableErrors = TableErrors + BuildOptionCharges(RawOptions, OptionCharges, Findings, FindingCount)
        RawFxTable = Inputs(IN_FX_TABLE)
        BuildFxTable RawFxTable, ReportingCcy, FxCodes, FxRates, ReportingErrors, DuplicateErrors, _
                     RateErrors, Findings, FindingCount
        LadderLabels = Inputs(IN_LADDER_LABELS)

        RuleCounts(5) = ReportingErrors
        RuleCounts(6) = DuplicateErrors
        RuleCounts(7) = RateErrors
        VolumeErrors = PositionVolumeErrors(Inputs, Findings, FindingCount)
        RuleCounts(14) = ParameterErrors + TableErrors + VolumeErrors
        If ParameterErrors = 0 Then RuleCounts(8) = ChangedParameterCount(Parameters, GoldCode, TransitionEnd)

        If ParameterErrors + TableErrors + ReportingErrors + DuplicateErrors + RateErrors _
           + VolumeErrors > 0 Then
            '--- Stopped: no calculation on unusable inputs ------------------------
            Outputs(OUT_STOPPED) = True
            CountBySeverity RuleCounts, Errors, Warnings
            Outputs(OUT_RULES) = RuleTable(RuleCounts)
            Outputs(OUT_FINDINGS) = FindingTable(Findings, FindingCount)
            Outputs(OUT_ERRORS) = Errors
            Outputs(OUT_WARNINGS) = Warnings
            Outputs(OUT_STATUS) = "STOPPED: " & Errors & " unusable input(s); see Checks. " & _
                                  "No results were calculated."
            Exit Sub
        End If

        '--- Risk classes --------------------------------------------------------
        'Each table has its own variable: an array stored in Outputs must not be
        'reused, as LibreOffice (the test host) keeps arrays in a Variant by
        'reference.
        IrRows = Inputs(IN_IR_ROWS)
        CalculateInterestRate IrRows, AsOfDate, Parameters, FxCodes, FxRates, GoldCode, SrKeys, _
                              SrCharges, Ladder, LadderLabels, IrCalc, LadderTable, ZoneTable, IrTotals
        Outputs(OUT_IR_CALC) = IrCalc
        Outputs(OUT_LADDER) = LadderTable
        Outputs(OUT_ZONES) = ZoneTable
        AddRowFindings IrCalc, SH_IR_POS, CLng(Inputs(IN_IR_FIRST_ROW)), Findings, FindingCount

        EqRows = Inputs(IN_EQ_ROWS)
        CalculateEquity EqRows, Parameters, FxCodes, FxRates, GoldCode, EqCalc, MarketTable, EqTotals
        Outputs(OUT_EQ_CALC) = EqCalc
        Outputs(OUT_MARKETS) = MarketTable
        AddRowFindings EqCalc, SH_EQ_POS, CLng(Inputs(IN_EQ_FIRST_ROW)), Findings, FindingCount

        CoRows = Inputs(IN_CO_ROWS)
        CalculateCommodity CoRows, Parameters, FxCodes, FxRates, GoldCode, CoCalc, CommodityTable, CoTotals
        Outputs(OUT_CO_CALC) = CoCalc
        Outputs(OUT_COMMODITIES) = CommodityTable
        AddRowFindings CoCalc, SH_CO_POS, CLng(Inputs(IN_CO_FIRST_ROW)), Findings, FindingCount

        FxRows = Inputs(IN_FX_ROWS)
        CalculateForeignExchange FxRows, Parameters, FxCodes, FxRates, ReportingCcy, GoldCode, _
                                 FxCalc, CurrencyTable, FxTotals
        Outputs(OUT_FX_CALC) = FxCalc
        Outputs(OUT_CURRENCIES) = CurrencyTable
        AddRowFindings FxCalc, SH_FX_POS, CLng(Inputs(IN_FX_FIRST_ROW)), Findings, FindingCount
        Summary(1, 1) = FxTotals(TOT_ONOP)
        Summary(2, 1) = FxTotals(TOT_GENERAL)
        Outputs(OUT_FX_SUMMARY) = Summary

        '--- Results: MAR40.2 aggregation ----------------------------------------
        ResultTable = ResultsTable(IrTotals, EqTotals, CoTotals, FxTotals, OptionCharges, Parameters, _
                                   Regime, AsOfDate, TransitionEnd)
        Outputs(OUT_RESULTS) = ResultTable
        Outputs(OUT_REQUIREMENT_BEFORE) = ResultTable(RC_COUNT + 1, 7)
        Outputs(OUT_REQUIREMENT) = ResultTable(RC_COUNT + 2, 7)
        Outputs(OUT_RWA) = ResultTable(RC_COUNT + 2, 8)
        Outputs(OUT_USED) = ResultTable(RC_COUNT + 1, 10)

        '--- Checks rules --------------------------------------------------------
        RuleCounts(1) = IrTotals(TOT_REJECTED)
        RuleCounts(2) = EqTotals(TOT_REJECTED)
        RuleCounts(3) = CoTotals(TOT_REJECTED)
        RuleCounts(4) = FxTotals(TOT_REJECTED)
        RuleCounts(9) = IrTotals(TOT_MIXED)
        RuleCounts(10) = MaxOf(0, EqTotals(TOT_BUCKETS) - CDbl(Inputs(IN_EQ_CAPACITY)))
        RuleCounts(11) = MaxOf(0, CoTotals(TOT_BUCKETS) - CDbl(Inputs(IN_CO_CAPACITY)))
        RuleCounts(12) = FxTotals(TOT_EXCLUDED)
        RuleCounts(13) = 0
        For i = 1 To RC_COUNT
            If OptionCharges(i) <> 0 Then RuleCounts(13) = RuleCounts(13) + 1
        Next i
        CountBySeverity RuleCounts, Errors, Warnings
        Outputs(OUT_RULES) = RuleTable(RuleCounts)
        Outputs(OUT_FINDINGS) = FindingTable(Findings, FindingCount)
        Outputs(OUT_ERRORS) = Errors
        Outputs(OUT_WARNINGS) = Warnings
        If Errors > 0 Then
            Outputs(OUT_STATUS) = "INCOMPLETE: " & Errors & " error(s) on Checks; rejected rows " & _
                                  "are excluded from the totals."
        Else
            Outputs(OUT_STATUS) = "OK: " & Outputs(OUT_USED) & " positions used, no errors. " & _
                                  "Calculated by RunSSA, engine " & ENGINE_VERSION & "."
        End If

End Sub


Public Function PositionVolumeErrors( _
    ByRef Inputs() As Variant, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'==============================================================================
'                            PositionVolumeErrors
'------------------------------------------------------------------------------
' PURPOSE
'   Find the position sheets that hold more rows than the workbook supports
'   (MAX_POSITION_ROWS) and record a finding for each.
'
' INPUTS
'   Inputs         Vector (1 To IN_COUNT); only the four position tables
'                  (IN_IR_ROWS, IN_EQ_ROWS, IN_CO_ROWS, IN_FX_ROWS) are read.
'                  A table spans the row below its header to its last row
'                  with a PositionID, blank rows between included.
'   Findings       Findings list, appended to (CORE_Util.AddFinding).
'   FindingCount   Findings in use; incremented.
'
' RETURNS
'   The number of sheets over the limit: 0 when every table has at most
'   MAX_POSITION_ROWS rows. The finding names the sheet and the rows found.
'
' UPDATED
'   2026-11-19
'==============================================================================
'
    Dim SheetNames  As Variant   'Position sheet names
    Dim Slots       As Variant   'Input slot of each table
    Dim Table       As Variant   'One position table
    Dim RowsFound   As Long      'Rows of one table
    Dim i           As Long      'Table

        SheetNames = Array(SH_IR_POS, SH_EQ_POS, SH_CO_POS, SH_FX_POS)
        Slots = Array(IN_IR_ROWS, IN_EQ_ROWS, IN_CO_ROWS, IN_FX_ROWS)
        For i = 0 To 3
            Table = Inputs(Slots(i))
            RowsFound = UBound(Table, 1)
            If RowsFound > MAX_POSITION_ROWS Then
                PositionVolumeErrors = PositionVolumeErrors + 1
                AddFinding Findings, FindingCount, CStr(SheetNames(i)), "Rows", vbNullString, _
                           ST_INVALID & RowsFound & " rows found; the workbook supports at most " _
                           & MAX_POSITION_ROWS
            End If
        Next i

End Function


'------------------------------------------------------------------------------
'                                    HELPERS
'------------------------------------------------------------------------------

Private Function ResultsTable( _
    ByRef IrTotals() As Double, _
    ByRef EqTotals() As Double, _
    ByRef CoTotals() As Double, _
    ByRef FxTotals() As Double, _
    ByRef OptionCharges() As Double, _
    ByRef Parameters() As Double, _
    ByVal Regime As String, _
    ByVal AsOfDate As Double, _
    ByVal TransitionEnd As Double) _
    As Variant
'
'==============================================================================
'                                 ResultsTable
'------------------------------------------------------------------------------
' PURPOSE
'   Build the Results table: per risk class the specific, general and other
'   charges, the option charge, the requirement before and after scaling,
'   RWA, share, positions used and rows rejected; then a total row, and a
'   row for the transitional multiplier.
'
' INPUTS
'   IrTotals, EqTotals, CoTotals, FxTotals   Totals vectors of each class.
'   OptionCharges                            (1 To RC_COUNT), reporting ccy.
'   Parameters                               Vector; scaling, RWA and the
'                                            transitional multiplier used.
'   Regime, AsOfDate, TransitionEnd          Regime from BuildRegime, reporting
'                                            date and last date of the multiplier.
'
' RETURNS
'   Variant (1 To RC_COUNT + 2, 1 To RESULTS_WIDTH): Results columns E:O. The
'   row after the total holds the factor of TransitionalFactor in column 6, the
'   requirement multiplied by it in column 7 and its RWA in column 8; the
'   other columns are empty. Under the Basel regime the factor is 1.
'
' REFERENCE
'   MAR40.1-40.2; Article 325(2) and Article 495s(2) of Regulation (EU)
'   No 575/2013 for the transitional row.
'
' UPDATED
'   2026-11-20
'==============================================================================
'
    Dim Table(1 To 6, 1 To 11) As Variant    'Result (RESULTS_ROWS rows, RESULTS_WIDTH columns)
    Dim Total       As Double    'Column total
    Dim i           As Long      'Risk class
    Dim c           As Long      'Column

        Table(RC_IR, 1) = IrTotals(TOT_SPECIFIC)
        Table(RC_IR, 2) = IrTotals(TOT_GENERAL)
        Table(RC_IR, 3) = 0#
        Table(RC_EQ, 1) = EqTotals(TOT_SPECIFIC)
        Table(RC_EQ, 2) = EqTotals(TOT_GENERAL)
        Table(RC_EQ, 3) = EqTotals(TOT_OTHER)
        Table(RC_CO, 1) = 0#
        Table(RC_CO, 2) = CoTotals(TOT_GENERAL)
        Table(RC_CO, 3) = CoTotals(TOT_OTHER)
        Table(RC_FX, 1) = 0#
        Table(RC_FX, 2) = FxTotals(TOT_GENERAL)
        Table(RC_FX, 3) = 0#
        Table(RC_IR, 6) = Parameters(PI_SCALING_IR)
        Table(RC_EQ, 6) = Parameters(PI_SCALING_EQ)
        Table(RC_CO, 6) = Parameters(PI_SCALING_CO)
        Table(RC_FX, 6) = Parameters(PI_SCALING_FX)
        Table(RC_IR, 10) = IrTotals(TOT_USED)
        Table(RC_EQ, 10) = EqTotals(TOT_USED)
        Table(RC_CO, 10) = CoTotals(TOT_USED)
        Table(RC_FX, 10) = FxTotals(TOT_USED)
        Table(RC_IR, 11) = IrTotals(TOT_REJECTED)
        Table(RC_EQ, 11) = EqTotals(TOT_REJECTED)
        Table(RC_CO, 11) = CoTotals(TOT_REJECTED)
        Table(RC_FX, 11) = FxTotals(TOT_REJECTED)
        For i = 1 To RC_COUNT
            Table(i, 4) = OptionCharges(i)
            Table(i, 5) = Table(i, 1) + Table(i, 2) + Table(i, 3) + Table(i, 4)
            Table(i, 7) = Table(i, 5) * Table(i, 6)
            Table(i, 8) = Table(i, 7) * Parameters(PI_RWA_MULTIPLIER)
        Next i

        'Total row, then each class's share of the scaled total.
        For c = 1 To RESULTS_WIDTH
            If c = 6 Or c = 9 Then
                Table(RC_COUNT + 1, c) = Empty
            Else
                Total = 0#
                For i = 1 To RC_COUNT
                    Total = Total + Table(i, c)
                Next i
                Table(RC_COUNT + 1, c) = Total
            End If
        Next c
        For i = 1 To RC_COUNT
            If Table(RC_COUNT + 1, 7) = 0 Then
                Table(i, 9) = 0#
            Else
                Table(i, 9) = Table(i, 7) / Table(RC_COUNT + 1, 7)
            End If
        Next i
        Total = 0#
        For i = 1 To RC_COUNT
            Total = Total + Table(i, 9)
        Next i
        Table(RC_COUNT + 1, 9) = Total

        'Transitional row: the factor, the requirement after it and its RWA.
        For c = 1 To RESULTS_WIDTH
            Table(RC_COUNT + 2, c) = Empty
        Next c
        Table(RC_COUNT + 2, 6) = TransitionalFactor(Regime, AsOfDate, TransitionEnd, _
                                                    Parameters(PI_TRANSITIONAL))
        Table(RC_COUNT + 2, 7) = Table(RC_COUNT + 1, 7) * Table(RC_COUNT + 2, 6)
        Table(RC_COUNT + 2, 8) = Table(RC_COUNT + 2, 7) * Parameters(PI_RWA_MULTIPLIER)
        ResultsTable = Table

End Function


Private Sub AddRowFindings( _
    ByRef Table() As Variant, _
    ByVal SheetName As String, _
    ByVal FirstRow As Long, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long)
'
'==============================================================================
'                                AddRowFindings
'------------------------------------------------------------------------------
' PURPOSE
'   Add one finding per rejected row of a calculation table.
'
' INPUTS
'   Table         Calculation table; column 1 PositionID, column 2 status.
'   SheetName     Input sheet the rows come from.
'   FirstRow      Sheet row of the table's first data row.
'   Findings      Findings list, appended to; FindingCount incremented.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim r           As Long      'Table row

        For r = 1 To UBound(Table, 1)
            If IsInvalidStatus(CStr(Table(r, 2))) Then
                AddFinding Findings, FindingCount, SheetName, CStr(Table(r, 1)), _
                           CStr(FirstRow + r - 1), CStr(Table(r, 2))
            End If
        Next r

End Sub


Private Function FindingTable( _
    ByRef Findings() As String, _
    ByVal FindingCount As Long) _
    As Variant
'
'==============================================================================
'                                 FindingTable
'------------------------------------------------------------------------------
' PURPOSE
'   Lay the findings out as the Checks findings table: number within the
'   sheet, sheet, item, row, two blank columns, reason.
'
' INPUTS
'   Findings       (1 To 4, 1 To n) from CORE_Util.AddFinding.
'   FindingCount   Findings in use (>= 0).
'
' RETURNS
'   Variant (1 To FindingCount, 1 To FINDINGS_WIDTH), or Empty when there
'   are none.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Table()     As Variant   'Result
    Dim Number      As Long      'Number within the sheet
    Dim f           As Long      'Finding

        If FindingCount = 0 Then
            FindingTable = Empty
            Exit Function
        End If
        ReDim Table(1 To FindingCount, 1 To FINDINGS_WIDTH)
        For f = 1 To FindingCount
            If f = 1 Then
                Number = 1
            ElseIf Findings(1, f) = Findings(1, f - 1) Then
                Number = Number + 1
            Else
                Number = 1
            End If
            Table(f, 1) = Number
            Table(f, 2) = Findings(1, f)
            Table(f, 3) = Findings(2, f)
            If Len(Findings(3, f)) > 0 Then Table(f, 4) = CLng(Findings(3, f))
            Table(f, 7) = Findings(4, f)
        Next f
        FindingTable = Table

End Function


Private Function RuleTable( _
    ByRef RuleCounts() As Variant) _
    As Variant
'
'==============================================================================
'                                   RuleTable
'------------------------------------------------------------------------------
' PURPOSE
'   Pair each rule identifier with its count.
'
' INPUTS
'   RuleCounts   (1 To RULE_COUNT); Empty for a rule not evaluated.
'
' RETURNS
'   Variant (1 To RULE_COUNT, 1 To 2): "C01".."C14", count or Empty.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Table(1 To 14, 1 To 2) As Variant    'Result (RULE_COUNT rows)
    Dim i           As Long      'Rule

        For i = 1 To RULE_COUNT
            Table(i, 1) = "C" & Right$("0" & i, 2)
            Table(i, 2) = RuleCounts(i)
        Next i
        RuleTable = Table

End Function


Private Sub CountBySeverity( _
    ByRef RuleCounts() As Variant, _
    ByRef Errors As Long, _
    ByRef Warnings As Long)
'
'==============================================================================
'                                CountBySeverity
'------------------------------------------------------------------------------
' PURPOSE
'   Sum the rule counts by severity (RULE_SEVERITIES).
'
' INPUTS
'   RuleCounts   (1 To RULE_COUNT); Empty counts as 0.
'
' RETURNS
'   Errors, Warnings: sums of the ERROR and WARNING rules.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Count       As Double    'Count of one rule
    Dim i           As Long      'Rule

        Errors = 0
        Warnings = 0
        For i = 1 To RULE_COUNT
            If IsEmpty(RuleCounts(i)) Then Count = 0 Else Count = CDbl(RuleCounts(i))
            Select Case ListItem(RULE_SEVERITIES, i)
                Case SEV_ERROR
                    Errors = Errors + Count
                Case SEV_WARNING
                    Warnings = Warnings + Count
            End Select
        Next i

End Sub


Private Function MaxOf( _
    ByVal ValueA As Double, _
    ByVal ValueB As Double) _
    As Double
'
'==============================================================================
'                                     MaxOf
'------------------------------------------------------------------------------
' PURPOSE
'   Larger of two numbers, without the Excel object model.
'
' RETURNS
'   max(ValueA, ValueB).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        If ValueA > ValueB Then MaxOf = ValueA Else MaxOf = ValueB

End Function
