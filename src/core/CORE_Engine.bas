Attribute VB_Name = "CORE_Engine"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Engine
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the whole MAR40 simplified standardised approach from raw cell values
'   to every output table: validates Params, stops when a parameter is unusable
'   or a position sheet holds more rows than the workbook supports, otherwise
'   calculates the four risk classes, scales and sums them into the own funds
'   requirement and RWA, assesses the eligibility for the simplified
'   standardised approach and the small trading book derogation under the CRR,
'   and assembles the Checks rules, the findings and the Results status line
'
' WHY THIS EXISTS
'   Keeps the sequence of the whole calculation in one procedure that depends
'   only on its arguments and never touches the Excel object model, so it can
'   be tested without a workbook
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Calculate runs the steps in order:
'     - validates the parameters, the regime, the ladder, the specific-risk table,
'       the option charges and the FX table, and counts the rows of the position
'       sheets
'     - stops, with only the Checks outputs, when an input is unusable
'     - otherwise calculates the interest rate, equity, commodity and foreign
'       exchange classes, assesses the eligibility of the CRR approaches,
'       aggregates the classes in the Results table and evaluates the Checks
'       rules
'
'   Public surface, visible only inside this VBA project (Option Private Module):
'   Calculate and PositionVolumeErrors. The other procedures are private helpers:
'   ResultsTable, AddRowFindings, FindingTable, RuleTable, CountBySeverity and
'   MaxOf
'
' ERROR POLICY
'   Cell content never raises: unusable values become findings. Arrays of the
'   wrong shape raise ERR_ARGUMENT from the module that detects them, propagated
'   unchanged
'
' DEPENDENCIES
'   CORE_Config, CORE_Util, CORE_Parameters, CORE_InterestRate, CORE_Equity,
'   CORE_Commodity, CORE_ForeignExchange and CORE_Eligibility
'
'   Reference: BCBS MAR40.1-40.2 (aggregation and scaling), MAR40 throughout;
'   Articles 94 and 325a of Regulation (EU) No 575/2013 (calculation contract
'   CALC-09)
'
' NOTES
'   Option Explicit requires explicit declarations; Option Private Module keeps
'   the procedures inside this project
'
'   State ownership: none. The procedures depend only on their arguments and
'   never touch the Excel object model, so they can be tested without a workbook
'
'   Compatibility: Excel VBA; no references beyond the defaults
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'
'------------------------------------------------------------------------------
'
'                                 CALCULATION
'
'------------------------------------------------------------------------------
'

Public Sub Calculate( _
    ByRef Inputs() As Variant, _
    ByRef Outputs() As Variant)
'
'------------------------------------------------------------------------------
'                       CALCULATE OWN FUNDS REQUIREMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Calculates the own funds requirement and every intermediate table
'
' WHY THIS EXISTS
'   Gives the caller one entry point that turns the input vector into the full
'   set of output tables
'
' INPUTS
'   Inputs
'     Vector (1 To IN_COUNT), filled as CORE_Config describes for each IN_ index:
'     raw Params values and tables, the four input tables as cell values, the
'     sheet row of each table's first data row, the market and commodity
'     capacity of Buckets and the rows of Eligibility as cell values. Read only;
'     the procedure does not change it
'
'   Outputs
'     Vector (1 To OUT_COUNT), sized here and filled as CORE_Config describes
'     for each OUT_ index
'
' RETURNS
'   Nothing. The results are in Outputs. When Outputs(OUT_STOPPED) is True, only
'   OUT_STOPPED, OUT_RULES, OUT_FINDINGS, OUT_ERRORS, OUT_WARNINGS and OUT_STATUS
'   are set
'
' BEHAVIOR
'   Sizes Outputs, validates the parameters, the regime and the tables, and
'   counts the Checks rules that depend on the inputs. If any parameter, table
'   or position sheet is unusable, it sets the stopped outputs and exits. Otherwise
'   it calculates the interest rate, equity, commodity and foreign exchange
'   classes in turn, adds the findings of their rejected rows, assesses the
'   eligibility of the CRR approaches, builds the Results table, evaluates the
'   remaining Checks rules and sets the status line. The eligibility changes no
'   requirement: its invalid rows are errors of rule C15, counted and not
'   listed as findings, and its approaches not met or to cease are warnings
'   of rule C16
'
' ERROR POLICY
'   Cell content never raises: unusable values become findings. Arrays of the
'   wrong shape raise ERR_ARGUMENT from the module that detects them, propagated
'   unchanged. The procedure has no error handler
'
' DEPENDENCIES
'   BuildParameters, BuildRegime, BuildLadder, BuildSpecificRiskTable,
'   BuildOptionCharges, BuildFxTable, ChangedParameterCount, TransitionalFactor
'   (through ResultsTable), CalculateInterestRate, CalculateEquity,
'   CalculateCommodity, CalculateForeignExchange, DeMinimisWaived, PositionVolumeErrors,
'   CalculateEligibility, AddRowFindings, ResultsTable, CountBySeverity, RuleTable,
'   FindingTable and MaxOf; the IN_, OUT_ and other constants of CORE_Config
'
'   Reference: MAR40.2: requirement = 1.30 IR + 3.50 EQ + 1.90 CO + 1.20 FX, each
'   class including its non-delta option charge; MAR40.1: RWA = 12.5 x it; CRR
'   Article 351 for the foreign exchange de minimis; CRR Articles 94 and 325a
'   for the eligibility (CALC-09)
'
' NOTES
'   Each table has its own variable: LibreOffice (the test host) keeps arrays in
'   a Variant by reference, so an array stored in Outputs must not be reused
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parameters()                As Double   'Validated parameter vector
    Dim AsOfDate                    As Double   'Reporting date serial
    Dim ReportingCcy                As String   'Reporting currency
    Dim GoldCode                    As String   'Gold code
    Dim Regime                      As String   'REGIME_BASEL or REGIME_CRR
    Dim TransitionEnd               As Double   'Last date of the transitional multiplier
    Dim OwnFunds                    As Double   'Total own funds
    Dim Ladder()                    As Double   'Validated ladder
    Dim LadderLabels()              As String   'Band labels
    Dim SrKeys()                    As String   'Specific-risk keys
    Dim SrCharges()                 As Double   'Specific-risk charges
    Dim FxCodes()                   As String   'FX codes
    Dim FxRates()                   As Double   'FX rates
    Dim OptionCharges()             As Double   'Non-delta option charges per class
    Dim Findings()                  As String   'Findings (1 To 4, 1 To n)
    Dim FindingCount                As Long     'Findings in use
    Dim ParameterErrors             As Long     'BuildParameters errors
    Dim VolumeErrors                As Long     'Position sheets over MAX_POSITION_ROWS
    Dim TableErrors                 As Long     'Ladder, specific-risk and option errors
    Dim ReportingErrors             As Long     'C05
    Dim DuplicateErrors             As Long     'C06
    Dim RateErrors                  As Long     'C07
    Dim RuleCounts(1 To RULE_COUNT) As Variant  'Count per rule, Empty when not evaluated
    Dim IrTotals()                  As Double   'Interest rate totals
    Dim EqTotals()                  As Double   'Equity totals
    Dim CoTotals()                  As Double   'Commodity totals
    Dim FxTotals()                  As Double   'FX totals
    Dim IrCalc()                    As Variant  'IRCalc table
    Dim LadderTable()               As Variant  'IR_Ladder table
    Dim ZoneTable()                 As Variant  'IR_Zones table
    Dim EqCalc()                    As Variant  'EQCalc table
    Dim MarketTable()               As Variant  'Buckets equity block
    Dim CoCalc()                    As Variant  'COCalc table
    Dim CommodityTable()            As Variant  'Buckets commodity block
    Dim FxCalc()                    As Variant  'FXCalc table
    Dim CurrencyTable()             As Variant  'Buckets FX block
    Dim ResultTable                 As Variant  'Results table
    Dim Summary(1 To 2, 1 To 1)     As Variant  'ONOP and FX charge
    Dim Errors                      As Long     'Error count
    Dim Warnings                    As Long     'Warning count
    Dim RawNumeric()                As Variant  'Numeric parameters as read
    Dim RawLadder()                 As Variant  'Ladder as read
    Dim RawSrTable()                As Variant  'Specific-risk table as read
    Dim RawOptions()                As Variant  'Option charges as read
    Dim RawFxTable()                As Variant  'FX table as read
    Dim IrRows()                    As Variant  'IRPositions cells
    Dim EqRows()                    As Variant  'EQPositions cells
    Dim CoRows()                    As Variant  'COPositions cells
    Dim FxRows()                    As Variant  'FXPositions cells
    Dim EligibilityRows()           As Variant  'Eligibility cells
    Dim EligibilityCalc()           As Variant  'Eligibility calculated columns
    Dim EligibilityBlock()          As Variant  'Eligibility result block
    Dim RejectedMonths              As Long     'C15
    Dim FlaggedApproaches           As Long     'C16
    Dim i                           As Long     'Rule or class index

'------------------------------------------------------------------------------
' INITIALIZE OUTPUTS
'------------------------------------------------------------------------------
    'Size the output vector and mark the run as not stopped
        ReDim Outputs(1 To OUT_COUNT)
        Outputs(OUT_STOPPED) = False

'------------------------------------------------------------------------------
' VALIDATE PARAMETERS
'------------------------------------------------------------------------------
    'Read the numeric parameters and validate them with the as-of date, reporting currency and gold code
        RawNumeric = Inputs(IN_NUMERIC)
        ParameterErrors = BuildParameters(RawNumeric, Inputs(IN_ASOF), Inputs(IN_REPCCY), _
            Inputs(IN_GOLD), Parameters, AsOfDate, ReportingCcy, _
            GoldCode, Findings, FindingCount)
    'Add the errors of the regime, the transitional end date and the total own funds
        ParameterErrors = ParameterErrors + BuildRegime(Inputs(IN_REGIME), Inputs(IN_TRANSITION_END), _
            Inputs(IN_OWN_FUNDS), Regime, TransitionEnd, _
            OwnFunds, Findings, FindingCount)

'------------------------------------------------------------------------------
' VALIDATE TABLES
'------------------------------------------------------------------------------
    'Validate the ladder
        RawLadder = Inputs(IN_LADDER)
        TableErrors = BuildLadder(RawLadder, Ladder, Findings, FindingCount)
    'Validate the specific-risk table
        RawSrTable = Inputs(IN_SR_TABLE)
        TableErrors = TableErrors + BuildSpecificRiskTable(RawSrTable, SrKeys, SrCharges, _
            Findings, FindingCount)
    'Validate the option charges
        RawOptions = Inputs(IN_OPTIONS)
        TableErrors = TableErrors + BuildOptionCharges(RawOptions, OptionCharges, Findings, FindingCount)
    'Validate the FX table against the reporting currency, counting the errors of rules C05, C06 and C07
        RawFxTable = Inputs(IN_FX_TABLE)
        BuildFxTable RawFxTable, ReportingCcy, FxCodes, FxRates, ReportingErrors, DuplicateErrors, _
            RateErrors, Findings, FindingCount
    'Read the ladder band labels
        LadderLabels = Inputs(IN_LADDER_LABELS)

'------------------------------------------------------------------------------
' COUNT INPUT RULES
'------------------------------------------------------------------------------
    'Record the FX table error counts as the counts of rules C05, C06 and C07
        RuleCounts(5) = ReportingErrors
        RuleCounts(6) = DuplicateErrors
        RuleCounts(7) = RateErrors
    'Count the position sheets that hold more rows than the workbook supports
        VolumeErrors = PositionVolumeErrors(Inputs, Findings, FindingCount)
    'Rule C14 counts the unusable parameters, tables and position sheets
        RuleCounts(14) = ParameterErrors + TableErrors + VolumeErrors
    'Count the changed parameters in rule C08, only when the parameters have no errors
        If ParameterErrors = 0 Then RuleCounts(8) = ChangedParameterCount(Parameters, GoldCode, TransitionEnd)

'------------------------------------------------------------------------------
' STOP ON UNUSABLE INPUTS
'------------------------------------------------------------------------------
    'When any input is unusable, return only the stopped flag, the Checks rules, the findings and the status line
        If ParameterErrors + TableErrors + ReportingErrors + DuplicateErrors + RateErrors _
            + VolumeErrors > 0 Then
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

'------------------------------------------------------------------------------
' APPLY REGIME
'------------------------------------------------------------------------------
    'Flag the CRR regime in the parameter vector
        If Regime = REGIME_CRR Then Parameters(PI_REGIME_FLAG) = 1#

'------------------------------------------------------------------------------
' CALCULATE INTEREST RATE
'------------------------------------------------------------------------------
    'Use a separate variable per table, because LibreOffice (the test host) keeps arrays in a Variant by reference
        IrRows = Inputs(IN_IR_ROWS)
    'Calculate the interest rate charges, the ladder and the zone tables and the totals
        CalculateInterestRate IrRows, AsOfDate, Parameters, FxCodes, FxRates, GoldCode, SrKeys, _
            SrCharges, Ladder, LadderLabels, IrCalc, LadderTable, ZoneTable, IrTotals
    'Store the interest rate tables in the outputs
        Outputs(OUT_IR_CALC) = IrCalc
        Outputs(OUT_LADDER) = LadderTable
        Outputs(OUT_ZONES) = ZoneTable
    'Record a finding for each rejected interest rate row
        AddRowFindings IrCalc, SH_IR_POS, CLng(Inputs(IN_IR_FIRST_ROW)), Findings, FindingCount

'------------------------------------------------------------------------------
' CALCULATE EQUITY
'------------------------------------------------------------------------------
    'Read the equity position rows and calculate the equity charges, the market table and the totals
        EqRows = Inputs(IN_EQ_ROWS)
        CalculateEquity EqRows, Parameters, FxCodes, FxRates, GoldCode, EqCalc, MarketTable, EqTotals
    'Store the equity tables in the outputs
        Outputs(OUT_EQ_CALC) = EqCalc
        Outputs(OUT_MARKETS) = MarketTable
    'Record a finding for each rejected equity row
        AddRowFindings EqCalc, SH_EQ_POS, CLng(Inputs(IN_EQ_FIRST_ROW)), Findings, FindingCount

'------------------------------------------------------------------------------
' CALCULATE COMMODITY
'------------------------------------------------------------------------------
    'Read the commodity position rows and calculate the commodity charges, the commodity table and the totals
        CoRows = Inputs(IN_CO_ROWS)
        CalculateCommodity CoRows, Parameters, FxCodes, FxRates, GoldCode, CoCalc, CommodityTable, CoTotals
    'Store the commodity tables in the outputs
        Outputs(OUT_CO_CALC) = CoCalc
        Outputs(OUT_COMMODITIES) = CommodityTable
    'Record a finding for each rejected commodity row
        AddRowFindings CoCalc, SH_CO_POS, CLng(Inputs(IN_CO_FIRST_ROW)), Findings, FindingCount

'------------------------------------------------------------------------------
' CALCULATE FOREIGN EXCHANGE
'------------------------------------------------------------------------------
    'Read the FX position rows and calculate the FX charges, the currency table and the totals
        FxRows = Inputs(IN_FX_ROWS)
        CalculateForeignExchange FxRows, Parameters, FxCodes, FxRates, ReportingCcy, GoldCode, _
            FxCalc, CurrencyTable, FxTotals
    'Store the FX tables in the outputs
        Outputs(OUT_FX_CALC) = FxCalc
        Outputs(OUT_CURRENCIES) = CurrencyTable
    'Record a finding for each rejected FX row
        AddRowFindings FxCalc, SH_FX_POS, CLng(Inputs(IN_FX_FIRST_ROW)), Findings, FindingCount
    'Apply CRR Article 351: no requirement below the share of the total own funds
        If DeMinimisWaived(Regime, FxTotals(TOT_ONOP), OwnFunds, Parameters(PI_FX_DE_MINIMIS)) Then
            FxTotals(TOT_GENERAL) = 0#
        End If
    'Store the net open position and the general FX charge as the FX summary
        Summary(1, 1) = FxTotals(TOT_ONOP)
        Summary(2, 1) = FxTotals(TOT_GENERAL)
        Outputs(OUT_FX_SUMMARY) = Summary

'------------------------------------------------------------------------------
' ASSESS ELIGIBILITY
'------------------------------------------------------------------------------
    'Assess the eligibility of the CRR approaches, which changes no requirement
        EligibilityRows = Inputs(IN_EL_ROWS)
        CalculateEligibility EligibilityRows, Regime, Parameters, EligibilityCalc, EligibilityBlock, _
            RejectedMonths, FlaggedApproaches
    'Store the calculated columns and the result block of Eligibility
        Outputs(OUT_EL_CALC) = EligibilityCalc
        Outputs(OUT_EL_RESULTS) = EligibilityBlock

'------------------------------------------------------------------------------
' AGGREGATE RESULTS
'------------------------------------------------------------------------------
    'Aggregate the four classes into the Results table (MAR40.2)
        ResultTable = ResultsTable(IrTotals, EqTotals, CoTotals, FxTotals, OptionCharges, Parameters, _
            Regime, AsOfDate, TransitionEnd)
        Outputs(OUT_RESULTS) = ResultTable
    'Take the requirement before and after the transitional multiplier, the RWA and the positions used from the table
        Outputs(OUT_REQUIREMENT_BEFORE) = ResultTable(RC_COUNT + 1, 7)
        Outputs(OUT_REQUIREMENT) = ResultTable(RC_COUNT + 2, 7)
        Outputs(OUT_RWA) = ResultTable(RC_COUNT + 2, 8)
        Outputs(OUT_USED) = ResultTable(RC_COUNT + 1, 10)

'------------------------------------------------------------------------------
' EVALUATE CHECKS RULES
'------------------------------------------------------------------------------
    'Count the rejected rows of each class as the counts of rules C01 to C04
        RuleCounts(1) = IrTotals(TOT_REJECTED)
        RuleCounts(2) = EqTotals(TOT_REJECTED)
        RuleCounts(3) = CoTotals(TOT_REJECTED)
        RuleCounts(4) = FxTotals(TOT_REJECTED)
    'Count the mixed interest rate rows (C09) and the buckets over the market and commodity capacity (C10, C11)
        RuleCounts(9) = IrTotals(TOT_MIXED)
        RuleCounts(10) = MaxOf(0, EqTotals(TOT_BUCKETS) - CDbl(Inputs(IN_EQ_CAPACITY)))
        RuleCounts(11) = MaxOf(0, CoTotals(TOT_BUCKETS) - CDbl(Inputs(IN_CO_CAPACITY)))
    'Count the excluded FX positions (C12)
        RuleCounts(12) = FxTotals(TOT_EXCLUDED)
    'Count the classes that have a non-delta option charge (C13)
        RuleCounts(13) = 0
        For i = 1 To RC_COUNT
            If OptionCharges(i) <> 0 Then RuleCounts(13) = RuleCounts(13) + 1
        Next i
    'Count the rejected Eligibility rows (C15) and the approaches not met or to cease (C16)
        RuleCounts(15) = RejectedMonths
        RuleCounts(16) = FlaggedApproaches

'------------------------------------------------------------------------------
' SET CHECKS OUTPUTS
'------------------------------------------------------------------------------
    'Sum the rule counts by severity and store the rules, the findings and the counts
        CountBySeverity RuleCounts, Errors, Warnings
        Outputs(OUT_RULES) = RuleTable(RuleCounts)
        Outputs(OUT_FINDINGS) = FindingTable(Findings, FindingCount)
        Outputs(OUT_ERRORS) = Errors
        Outputs(OUT_WARNINGS) = Warnings
    'Set the status line: incomplete when there are errors, otherwise ok with the positions used
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
'------------------------------------------------------------------------------
'                       COUNT OVERSIZED POSITION SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Finds the position sheets that hold more rows than the workbook supports
'   (MAX_POSITION_ROWS) and records a finding for each
'
' WHY THIS EXISTS
'   Calculate stops on a position sheet over MAX_POSITION_ROWS, so the sheets
'   are counted before any calculation
'
' INPUTS
'   Inputs
'     Vector (1 To IN_COUNT); only the four position tables (IN_IR_ROWS,
'     IN_EQ_ROWS, IN_CO_ROWS, IN_FX_ROWS) are read. A table spans the row below
'     its header to its last row with a PositionID, blank rows between included
'
'   Findings
'     Findings list, appended to (CORE_Util.AddFinding)
'
'   FindingCount
'     Findings in use; incremented
'
' RETURNS
'   The number of sheets over the limit: 0 when every table has at most
'   MAX_POSITION_ROWS rows. The finding names the sheet and the rows found
'
' BEHAVIOR
'   Takes the row count of each of the four position tables in turn. For a
'   table over MAX_POSITION_ROWS it adds one to the result and appends a finding
'   with the status ST_INVALID, the sheet and the rows found
'
' ERROR POLICY
'   Raises nothing explicitly. The procedure has no error handler
'
' DEPENDENCIES
'   AddFinding; the IN_ slots, SH_ sheet names, ST_INVALID and MAX_POSITION_ROWS
'   of CORE_Config
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
    Dim SheetNames  As Variant  'Position sheet names
    Dim Slots       As Variant  'Input slot of each table
    Dim Table       As Variant  'One position table
    Dim RowsFound   As Long     'Rows of one table
    Dim i           As Long     'Table

'------------------------------------------------------------------------------
' LIST POSITION TABLES
'------------------------------------------------------------------------------
    'List the four position sheets and the input slot of each table
        SheetNames = Array(SH_IR_POS, SH_EQ_POS, SH_CO_POS, SH_FX_POS)
        Slots = Array(IN_IR_ROWS, IN_EQ_ROWS, IN_CO_ROWS, IN_FX_ROWS)

'------------------------------------------------------------------------------
' COUNT OVERSIZED SHEETS
'------------------------------------------------------------------------------
    'Count each table over the row limit and record a finding naming the sheet and the rows found
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

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

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
'------------------------------------------------------------------------------
'                             BUILD RESULTS TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the Results table: per risk class the specific, general and other
'   charges, the option charge, the requirement before and after scaling, RWA,
'   share, positions used and rows rejected; then a total row, and a row for the
'   transitional multiplier
'
' WHY THIS EXISTS
'   Puts the aggregation of the four classes (MAR40.1-40.2) in one table that
'   Calculate stores as the Results output
'
' INPUTS
'   IrTotals, EqTotals, CoTotals, FxTotals
'     Totals vectors of each class
'
'   OptionCharges
'     (1 To RC_COUNT), reporting currency
'
'   Parameters
'     Vector; the scaling factors, the RWA multiplier and the transitional
'     multiplier parameter are used
'
'   Regime, AsOfDate, TransitionEnd
'     Regime from BuildRegime, reporting date and last date of the multiplier
'
' RETURNS
'   Variant (1 To RC_COUNT + 2, 1 To RESULTS_WIDTH): Results columns E:O. The
'   row after the total holds the factor of TransitionalFactor in column 6, the
'   requirement multiplied by it in column 7 and its RWA in column 8; the other
'   columns are empty. Under the Basel regime the factor is 1
'
' BEHAVIOR
'   Fills the charges, scaling factors, positions used and rows rejected of each
'   class from the totals and the parameters. Adds the option charge, then
'   calculates the requirement before scaling (column 5), after scaling
'   (column 7) and the RWA (column 8) of each class. Sums the columns into the
'   total row, sets the share of each class in the scaled total (column 9) and
'   fills the transitional row
'
' ERROR POLICY
'   Raises nothing explicitly: a zero total requirement gives a share of 0, so
'   there is no division by zero. The procedure has no error handler
'
' DEPENDENCIES
'   TransitionalFactor; the RC_, TOT_, PI_ constants and RESULTS_WIDTH of
'   CORE_Config
'
'   Reference: MAR40.1-40.2; Article 325(2) and Article 495s(2) of Regulation (EU)
'   No 575/2013 for the transitional row
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Table(1 To 6, 1 To 11)  As Variant  'Result (RESULTS_ROWS rows, RESULTS_WIDTH columns)
    Dim Total                   As Double   'Column total
    Dim i                       As Long     'Risk class
    Dim c                       As Long     'Column

'------------------------------------------------------------------------------
' FILL CLASS FIGURES
'------------------------------------------------------------------------------
    'Take the specific, general and other charges of each class from its totals, with zero where a class has none
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
    'Take the scaling factor of each class from the parameters
        Table(RC_IR, 6) = Parameters(PI_SCALING_IR)
        Table(RC_EQ, 6) = Parameters(PI_SCALING_EQ)
        Table(RC_CO, 6) = Parameters(PI_SCALING_CO)
        Table(RC_FX, 6) = Parameters(PI_SCALING_FX)
    'Take the positions used of each class from its totals
        Table(RC_IR, 10) = IrTotals(TOT_USED)
        Table(RC_EQ, 10) = EqTotals(TOT_USED)
        Table(RC_CO, 10) = CoTotals(TOT_USED)
        Table(RC_FX, 10) = FxTotals(TOT_USED)
    'Take the rows rejected of each class from its totals
        Table(RC_IR, 11) = IrTotals(TOT_REJECTED)
        Table(RC_EQ, 11) = EqTotals(TOT_REJECTED)
        Table(RC_CO, 11) = CoTotals(TOT_REJECTED)
        Table(RC_FX, 11) = FxTotals(TOT_REJECTED)

'------------------------------------------------------------------------------
' CALCULATE REQUIREMENT
'------------------------------------------------------------------------------
    'Add the option charge, then the requirement before and after scaling and the RWA of each class
        For i = 1 To RC_COUNT
            Table(i, 4) = OptionCharges(i)
            Table(i, 5) = Table(i, 1) + Table(i, 2) + Table(i, 3) + Table(i, 4)
            Table(i, 7) = Table(i, 5) * Table(i, 6)
            Table(i, 8) = Table(i, 7) * Parameters(PI_RWA_MULTIPLIER)
        Next i

'------------------------------------------------------------------------------
' ADD TOTAL ROW
'------------------------------------------------------------------------------
    'Sum each column into the total row, leaving the scaling and share columns empty
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
    'Set each class's share of the scaled total, zero when the total is zero
        For i = 1 To RC_COUNT
            If Table(RC_COUNT + 1, 7) = 0 Then
                Table(i, 9) = 0#
            Else
                Table(i, 9) = Table(i, 7) / Table(RC_COUNT + 1, 7)
            End If
        Next i
    'Sum the shares into the total row
        Total = 0#
        For i = 1 To RC_COUNT
            Total = Total + Table(i, 9)
        Next i
        Table(RC_COUNT + 1, 9) = Total

'------------------------------------------------------------------------------
' ADD TRANSITIONAL ROW
'------------------------------------------------------------------------------
    'Empty every column of the transitional row
        For c = 1 To RESULTS_WIDTH
            Table(RC_COUNT + 2, c) = Empty
        Next c
    'Set the transitional factor, the requirement after it and its RWA
        Table(RC_COUNT + 2, 6) = TransitionalFactor(Regime, AsOfDate, TransitionEnd, _
            Parameters(PI_TRANSITIONAL))
        Table(RC_COUNT + 2, 7) = Table(RC_COUNT + 1, 7) * Table(RC_COUNT + 2, 6)
        Table(RC_COUNT + 2, 8) = Table(RC_COUNT + 2, 7) * Parameters(PI_RWA_MULTIPLIER)
    'Return the table
        ResultsTable = Table

End Function

Private Sub AddRowFindings( _
    ByRef Table() As Variant, _
    ByVal SheetName As String, _
    ByVal FirstRow As Long, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long)
'
'------------------------------------------------------------------------------
'                               ADD ROW FINDINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Adds one finding per rejected row of a calculation table
'
' WHY THIS EXISTS
'   Each rejected position row becomes a finding that names its sheet,
'   PositionID and sheet row
'
' INPUTS
'   Table
'     Calculation table; column 1 PositionID, column 2 status
'
'   SheetName
'     Input sheet the rows come from
'
'   FirstRow
'     Sheet row of the table's first data row
'
'   Findings
'     Findings list, appended to
'
'   FindingCount
'     Findings in use; incremented
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Goes through the rows of the table. For each row whose status is invalid it
'   appends a finding with the sheet, the PositionID, the sheet row (FirstRow
'   plus the table row minus 1) and the status
'
' ERROR POLICY
'   Raises nothing explicitly. The procedure has no error handler
'
' DEPENDENCIES
'   IsInvalidStatus and AddFinding (CORE_Util)
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
    Dim r   As Long 'Table row

'------------------------------------------------------------------------------
' ADD FINDINGS
'------------------------------------------------------------------------------
    'Record a finding for each row with an invalid status
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
'------------------------------------------------------------------------------
'                             BUILD FINDINGS TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Lays the findings out as the Checks findings table: number within the
'   sheet, sheet, item, row, two blank columns, reason
'
' WHY THIS EXISTS
'   Calculate returns the findings as the Checks findings table, in this column
'   order
'
' INPUTS
'   Findings
'     (1 To 4, 1 To n) from CORE_Util.AddFinding
'
'   FindingCount
'     Findings in use (>= 0)
'
' RETURNS
'   Variant (1 To FindingCount, 1 To FINDINGS_WIDTH), or Empty when there are
'   none
'
' BEHAVIOR
'   Returns Empty when there are no findings. Otherwise it sizes the table and
'   fills one row per finding: the number of the finding within its sheet
'   (restarting at 1 when the sheet changes), the sheet, the item, the row when
'   there is one (column 4) and the reason (column 7)
'
' ERROR POLICY
'   Raises nothing explicitly. The procedure has no error handler
'
' DEPENDENCIES
'   FINDINGS_WIDTH (CORE_Config)
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
    Dim Table() As Variant  'Result
    Dim Number  As Long     'Number within the sheet
    Dim f       As Long     'Finding

'------------------------------------------------------------------------------
' HANDLE NO FINDINGS
'------------------------------------------------------------------------------
    'Return Empty when there are no findings
        If FindingCount = 0 Then
            FindingTable = Empty
            Exit Function
        End If

'------------------------------------------------------------------------------
' BUILD TABLE
'------------------------------------------------------------------------------
    'Size the table to one row per finding
        ReDim Table(1 To FindingCount, 1 To FINDINGS_WIDTH)
    'Fill one row per finding, numbering the findings within their sheet and leaving the row blank when there is none
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
    'Return the table
        FindingTable = Table

End Function

Private Function RuleTable( _
    ByRef RuleCounts() As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                               BUILD RULE TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Pairs each rule identifier with its count
'
' WHY THIS EXISTS
'   The Checks rules table lists every rule with its count, in the order of the
'   rule identifiers
'
' INPUTS
'   RuleCounts
'     (1 To RULE_COUNT); Empty for a rule not evaluated
'
' RETURNS
'   Variant (1 To RULE_COUNT, 1 To 2): "C01".."C16", count or Empty
'
' BEHAVIOR
'   Writes the identifier "C" followed by the two-digit rule number in column 1
'   and the count of the rule in column 2
'
' ERROR POLICY
'   Raises nothing explicitly. The procedure has no error handler
'
' DEPENDENCIES
'   RULE_COUNT (CORE_Config)
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
    Dim Table(1 To 16, 1 To 2)  As Variant  'Result (RULE_COUNT rows)
    Dim i                       As Long     'Rule

'------------------------------------------------------------------------------
' BUILD TABLE
'------------------------------------------------------------------------------
    'Pair each rule identifier with its count
        For i = 1 To RULE_COUNT
            Table(i, 1) = "C" & Right$("0" & i, 2)
            Table(i, 2) = RuleCounts(i)
        Next i
    'Return the table
        RuleTable = Table

End Function

Private Sub CountBySeverity( _
    ByRef RuleCounts() As Variant, _
    ByRef Errors As Long, _
    ByRef Warnings As Long)
'
'------------------------------------------------------------------------------
'                              COUNT BY SEVERITY
'------------------------------------------------------------------------------
' PURPOSE
'   Sums the rule counts by severity (RULE_SEVERITIES)
'
' WHY THIS EXISTS
'   Calculate needs the number of error rules and of warning rules found, for
'   the Errors and Warnings outputs and the status line
'
' INPUTS
'   RuleCounts
'     (1 To RULE_COUNT); Empty counts as 0
'
'   Errors
'     Output: sum of the counts of the rules with severity SEV_ERROR
'
'   Warnings
'     Output: sum of the counts of the rules with severity SEV_WARNING
'
' RETURNS
'   Nothing. Errors and Warnings hold the sums of the ERROR and WARNING rules
'
' BEHAVIOR
'   Resets Errors and Warnings to 0, then adds the count of each rule to the sum
'   of its severity, taking an Empty count as 0
'
' ERROR POLICY
'   Raises nothing explicitly. The procedure has no error handler
'
' DEPENDENCIES
'   ListItem (CORE_Util); RULE_SEVERITIES, SEV_ERROR and SEV_WARNING (CORE_Config)
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
    Dim Count   As Double   'Count of one rule
    Dim i       As Long     'Rule

'------------------------------------------------------------------------------
' RESET TOTALS
'------------------------------------------------------------------------------
    'Start both sums at zero
        Errors = 0
        Warnings = 0

'------------------------------------------------------------------------------
' SUM BY SEVERITY
'------------------------------------------------------------------------------
    'Add the count of each rule to the sum of its severity, taking an Empty count as 0
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
'------------------------------------------------------------------------------
'                             RETURN LARGER NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the larger of two numbers, without the Excel object model
'
' WHY THIS EXISTS
'   Calculate needs the larger of two numbers, and this module does not use the
'   Excel object model
'
' INPUTS
'   ValueA
'     First number
'
'   ValueB
'     Second number
'
' RETURNS
'   max(ValueA, ValueB)
'
' BEHAVIOR
'   Compares the two numbers and returns the larger
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
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the larger of the two numbers
        If ValueA > ValueB Then MaxOf = ValueA Else MaxOf = ValueB

End Function
