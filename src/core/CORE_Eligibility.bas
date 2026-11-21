Attribute VB_Name = "CORE_Eligibility"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Eligibility
'------------------------------------------------------------------------------
' PURPOSE
'   Eligibility for the simplified standardised approach (Article 325a of the
'   CRR) and the derogation for small trading book business (Article 94),
'   under the CRR regime: validates each month of the Eligibility sheet, tests
'   the size of the business of each approach against its share of the total
'   assets and its amount, and gives each approach the status of the last
'   month entered. Returns the calculated columns of the sheet and its result
'   block
'
' WHY THIS EXISTS
'   Holds the assessment of calculation contract CALC-09 in one component whose
'   procedures depend only on their arguments and never touch the Excel object
'   model. The result is a finding; it changes no requirement
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Public surface: CalculateEligibility, visible only inside this VBA project
'   (Option Private Module). The private helpers are MonthStatus, which tests
'   one month, and MonthEndSerial, which finds the last day of a month
'
' ERROR POLICY
'   Cell content never raises: a month that fails a test gets an INVALID
'   status and contributes nothing. ERR_ARGUMENT is raised for an input array
'   of the wrong shape
'
' DEPENDENCIES
'   CORE_Config and CORE_Util
'
'   Reference: Regulation (EU) No 575/2013, Article 325a(1), (2), (5) and (6)
'   and Article 94(1), (3), (7) and (8); calculation contract CALC-09
'
' NOTES
'   Option Explicit requires explicit declarations and Option Private Module
'   keeps the procedures inside this project. The module owns no state: every
'   procedure depends only on its arguments
'
'   Compatibility: Excel VBA; no references beyond the defaults. The dates are
'   handled with DateSerial, Year and Month, which LibreOffice's VBA mode also
'   provides
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Columns of the row output (Eligibility I:U), counted from column I
    Private Const OC_STATUS         As Long = 1     'Month status
    Private Const OC_APPROACH_WIDTH As Long = 6     'Columns of one approach, 325a from J, 94 from P

    'Column of each figure within the columns of an approach
    Private Const AO_SIZE           As Long = 1     'Size of the business, euro
    Private Const AO_SHARE          As Long = 2     'Size as a share of the total assets
    Private Const AO_RESULT         As Long = 3     'MET or NOT MET
    Private Const AO_RUN            As Long = 4     'Consecutive months not met, this one included
    Private Const AO_COUNT          As Long = 5     'Months not met among the last 12 valid months
    Private Const AO_EVENT          As Long = 6     'YES or NO

    'Columns of the result block (Eligibility_Results D:J), one row per approach
    Private Const BC_STATUS         As Long = 1     'Status of the approach
    Private Const BC_MONTH          As Long = 2     'Month end of the last month entered
    Private Const BC_SIZE           As Long = 3     'Size of the last month
    Private Const BC_SHARE          As Long = 4     'Share of the last month
    Private Const BC_RUN            As Long = 5     'Run of the last month
    Private Const BC_COUNT          As Long = 6     'Count of the last month
    Private Const BC_RESTART        As Long = 7     'Restart condition

    'Articles 325a(5) and (6) and 94(7) and (8): a run of 3 months, or more than 6
    'of the last 12, is an event; 12 months met in a row allow a restart
    Private Const EVENT_RUN         As Long = 3     'Run that is an event
    Private Const EVENT_COUNT       As Long = 6     'Count that an event exceeds
    Private Const WINDOW_MONTHS     As Long = 12    'Months of the count and of the restart

    'Texts of the outputs
    Private Const TXT_MET           As String = "MET"               'Conditions met
    Private Const TXT_NOT_MET       As String = "NOT MET"           'Conditions not met
    Private Const TXT_CEASE         As String = "CEASE"             'Last month is an event
    Private Const TXT_YES           As String = "YES"               'Event or restart condition
    Private Const TXT_NO            As String = "NO"                'No event or restart condition
    Private Const TXT_NOT_ASSESSED  As String = "NOT ASSESSED: "    'Prefix, approach not assessed


'
'------------------------------------------------------------------------------
'
'                                 CALCULATION
'
'------------------------------------------------------------------------------
'

Public Sub CalculateEligibility( _
    ByRef InputRows() As Variant, _
    ByVal Regime As String, _
    ByRef Parameters() As Double, _
    ByRef RowOutput() As Variant, _
    ByRef ResultOutput() As Variant, _
    ByRef RejectedRows As Long, _
    ByRef FlaggedApproaches As Long)
'
'------------------------------------------------------------------------------
'                            ASSESS ELIGIBILITY
'------------------------------------------------------------------------------
' PURPOSE
'   Assesses the eligibility of both approaches from the months entered and
'   returns the calculated columns of each row and the result block
'
' WHY THIS EXISTS
'   The run, the count and the status of a month depend on the months before
'   it, so the rows are assessed in one pass in the order of the sheet
'
' INPUTS
'   InputRows
'     Eligibility cells (1 To n, 1 To EL_INPUT_WIDTH), as read: month end,
'     total assets, and the long and short totals of Article 325a and of
'     Article 94, in euro. A row whose month end is blank is skipped
'
'   Regime
'     REGIME_BASEL or REGIME_CRR, from BuildRegime; only the CRR is assessed
'
'   Parameters
'     Vector (1 To PI_VECTOR) from BuildParameters; PI_EL_SHARE_325A,
'     PI_EL_CAP_325A, PI_EL_SHARE_94 and PI_EL_CAP_94 are used
'
'   RowOutput
'     Output; dimensioned by the procedure, see RETURNS
'
'   ResultOutput
'     Output; dimensioned by the procedure, see RETURNS
'
'   RejectedRows
'     Output: the rows whose month status is INVALID (rule C15)
'
'   FlaggedApproaches
'     Output: the approaches whose status is NOT MET or CEASE (rule C16)
'
' RETURNS
'   Nothing; the results are written to the output arguments:
'     RowOutput      (1 To n, 1 To EL_CALC_WIDTH): Eligibility I:U, the month
'                    status, then size, share, conditions, run, count and
'                    event of Article 325a and the same of Article 94. A
'                    skipped row, and every row under the Basel regime, is
'                    Empty; an invalid row holds its status only
'     ResultOutput   (1 To EL_BLOCK_ROWS, 1 To EL_BLOCK_WIDTH): row 1 Article
'                    325a, row 2 Article 94; status, last month, size, share,
'                    run, count and restart condition. Where the status reads
'                    NOT ASSESSED, every other column is Empty
'
' BEHAVIOR
'   Validates the width of InputRows and sizes the outputs. Under the Basel
'   regime it sets both statuses to NOT ASSESSED: Basel regime and reads no
'   row. Otherwise, in the order of the rows:
'     - skips a row whose month end is blank
'     - gives each other row its month status with MonthStatus, and keeps its
'       month end as the previous one when it is a date
'     - for a valid row and each approach: the size, the absolute long total
'       plus the absolute short total; the share, the size over the total
'       assets; the conditions, met where the size is at most the share
'       parameter times the total assets and at most the amount; the run and
'       the count of the months not met over the valid rows; and the event,
'       a run of EVENT_RUN or more or a count over EVENT_COUNT
'   Then, with no month entered or an invalid row, both statuses read NOT
'   ASSESSED with the reason. Otherwise each approach takes the figures of the
'   last month entered: CEASE where it is an event, NOT MET where its
'   conditions are not met, MET otherwise; the restart condition is YES where
'   WINDOW_MONTHS valid months are entered and the count of the last is 0
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when InputRows has fewer than EL_INPUT_WIDTH columns.
'   Cell content never raises. The procedure has no error handler
'
' DEPENDENCIES
'   MonthStatus; CORE_Util: PlainText, IsDateSerial and NumberValue;
'   CORE_Config: the EL_, PI_EL_, REGIME_ and ST_ constants and ERR_ARGUMENT
'
'   Reference: Article 325a(1), (2), (5) and (6) and Article 94(1), (3), (7)
'   and (8) of Regulation (EU) No 575/2013; calculation contract CALC-09
'
' NOTES
'   Owns no state: it writes only to its output arguments. The status changes
'   no requirement and stops no run
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RowCount            As Long     'Input rows
    Dim LongColumn(1 To 2)  As Long     'Input column of the long total, by approach
    Dim ShortColumn(1 To 2) As Long     'Input column of the short total, by approach
    Dim ShareLimit(1 To 2)  As Double   'Share of the total assets, by approach
    Dim AmountLimit(1 To 2) As Double   'Amount in euro, by approach
    Dim Missed()            As Boolean  'Conditions not met, by approach and valid row
    Dim Runs(1 To 2)        As Long     'Run of the last valid row, by approach
    Dim Entered             As Long     'Rows with a month end
    Dim Valid               As Long     'Rows with status OK
    Dim LastRow             As Long     'Row of the last month entered
    Dim Previous            As Double   'Month end of the row entered before, 0 when not a date
    Dim Status              As String   'Month status of a row
    Dim Reason              As String   'Why the approaches are not assessed
    Dim Assets              As Double   'Total assets of a month
    Dim Size                As Double   'Size of the business of a month
    Dim Count               As Long     'Months not met among the last valid ones
    Dim ColumnBefore        As Long     'Row output column before the columns of an approach
    Dim r                   As Long     'Input row
    Dim a                   As Long     'Approach
    Dim v                   As Long     'Valid row

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an input array narrower than the Eligibility input columns
        If UBound(InputRows, 2) < EL_INPUT_WIDTH Then
            Err.Raise ERR_ARGUMENT, "CORE_Eligibility.CalculateEligibility", _
                "InputRows must be " & EL_INPUT_WIDTH & " columns wide."
        End If

'------------------------------------------------------------------------------
' SIZE OUTPUTS
'------------------------------------------------------------------------------
    'Size the row output and the result block, and start both counts at zero
        RowCount = UBound(InputRows, 1)
        ReDim RowOutput(1 To RowCount, 1 To EL_CALC_WIDTH)
        ReDim ResultOutput(1 To EL_BLOCK_ROWS, 1 To EL_BLOCK_WIDTH)
        RejectedRows = 0
        FlaggedApproaches = 0

'------------------------------------------------------------------------------
' SKIP OTHER REGIMES
'------------------------------------------------------------------------------
    'Under the Basel regime read no row and give both approaches the reason
        If Regime <> REGIME_CRR Then
            For a = 1 To EL_BLOCK_ROWS
                ResultOutput(a, BC_STATUS) = TXT_NOT_ASSESSED & "Basel regime"
            Next a
            Exit Sub
        End If

'------------------------------------------------------------------------------
' LOAD APPROACHES
'------------------------------------------------------------------------------
    'Article 325a: its totals, its share of the total assets and its amount
        LongColumn(1) = EL_LONG_325A
        ShortColumn(1) = EL_SHORT_325A
        ShareLimit(1) = Parameters(PI_EL_SHARE_325A)
        AmountLimit(1) = Parameters(PI_EL_CAP_325A)
    'Article 94: its totals, its share of the total assets and its amount
        LongColumn(2) = EL_LONG_94
        ShortColumn(2) = EL_SHORT_94
        ShareLimit(2) = Parameters(PI_EL_SHARE_94)
        AmountLimit(2) = Parameters(PI_EL_CAP_94)
    'One flag per approach and row, of which the valid rows use the first ones
        ReDim Missed(1 To EL_BLOCK_ROWS, 1 To RowCount)

'------------------------------------------------------------------------------
' ASSESS MONTHS
'------------------------------------------------------------------------------
    'Give each row with a month end its status, and a valid row its figures
        For r = 1 To RowCount
            If Len(PlainText(InputRows(r, EL_MONTH))) > 0 Then
                'Test the month against the month end of the row entered before it
                Entered = Entered + 1
                LastRow = r
                Status = MonthStatus(InputRows, r, Previous)
                RowOutput(r, OC_STATUS) = Status
                'Keep the month end as the previous one when it is a date
                If IsDateSerial(InputRows(r, EL_MONTH)) Then
                    Previous = NumberValue(InputRows(r, EL_MONTH))
                Else
                    Previous = 0#
                End If
                'Count an invalid row, and give a valid row the figures of each approach
                If Status <> ST_OK Then
                    RejectedRows = RejectedRows + 1
                Else
                    Valid = Valid + 1
                    Assets = NumberValue(InputRows(r, EL_ASSETS))
                    For a = 1 To EL_BLOCK_ROWS
                        'The size, and whether it exceeds the share of the total assets or the amount
                        Size = Abs(NumberValue(InputRows(r, LongColumn(a)))) + _
                               Abs(NumberValue(InputRows(r, ShortColumn(a))))
                        Missed(a, Valid) = Not (Size <= ShareLimit(a) * Assets And Size <= AmountLimit(a))
                        'The run of consecutive months not met, this one included
                        If Missed(a, Valid) Then
                            Runs(a) = Runs(a) + 1
                        Else
                            Runs(a) = 0
                        End If
                        'The months not met among the last WINDOW_MONTHS valid rows, this one included
                        Count = 0
                        For v = Valid - WINDOW_MONTHS + 1 To Valid
                            If v >= 1 Then
                                If Missed(a, v) Then Count = Count + 1
                            End If
                        Next v
                        'Write the figures of the approach
                        ColumnBefore = OC_STATUS + (a - 1) * OC_APPROACH_WIDTH
                        RowOutput(r, ColumnBefore + AO_SIZE) = Size
                        RowOutput(r, ColumnBefore + AO_SHARE) = Size / Assets
                        If Missed(a, Valid) Then
                            RowOutput(r, ColumnBefore + AO_RESULT) = TXT_NOT_MET
                        Else
                            RowOutput(r, ColumnBefore + AO_RESULT) = TXT_MET
                        End If
                        RowOutput(r, ColumnBefore + AO_RUN) = Runs(a)
                        RowOutput(r, ColumnBefore + AO_COUNT) = Count
                        If Runs(a) >= EVENT_RUN Or Count > EVENT_COUNT Then
                            RowOutput(r, ColumnBefore + AO_EVENT) = TXT_YES
                        Else
                            RowOutput(r, ColumnBefore + AO_EVENT) = TXT_NO
                        End If
                    Next a
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' HANDLE NOT ASSESSED
'------------------------------------------------------------------------------
    'With no month entered or an invalid row, give both approaches the reason and nothing else
        If Entered = 0 Or RejectedRows > 0 Then
            If Entered = 0 Then
                Reason = "no month entered"
            Else
                Reason = "a row is invalid"
            End If
            For a = 1 To EL_BLOCK_ROWS
                ResultOutput(a, BC_STATUS) = TXT_NOT_ASSESSED & Reason
            Next a
            Exit Sub
        End If

'------------------------------------------------------------------------------
' BUILD RESULT BLOCK
'------------------------------------------------------------------------------
    'Give each approach the status and the figures of the last month entered
        For a = 1 To EL_BLOCK_ROWS
            ColumnBefore = OC_STATUS + (a - 1) * OC_APPROACH_WIDTH
            'CEASE for an event, NOT MET for conditions not met, MET otherwise
            If RowOutput(LastRow, ColumnBefore + AO_EVENT) = TXT_YES Then
                ResultOutput(a, BC_STATUS) = TXT_CEASE
            ElseIf RowOutput(LastRow, ColumnBefore + AO_RESULT) = TXT_NOT_MET Then
                ResultOutput(a, BC_STATUS) = TXT_NOT_MET
            Else
                ResultOutput(a, BC_STATUS) = TXT_MET
            End If
            'Count an approach that is not met or must cease (rule C16)
            If ResultOutput(a, BC_STATUS) <> TXT_MET Then FlaggedApproaches = FlaggedApproaches + 1
            'The month end as entered, and the size, share, run and count of that month
            ResultOutput(a, BC_MONTH) = InputRows(LastRow, EL_MONTH)
            ResultOutput(a, BC_SIZE) = RowOutput(LastRow, ColumnBefore + AO_SIZE)
            ResultOutput(a, BC_SHARE) = RowOutput(LastRow, ColumnBefore + AO_SHARE)
            ResultOutput(a, BC_RUN) = RowOutput(LastRow, ColumnBefore + AO_RUN)
            ResultOutput(a, BC_COUNT) = RowOutput(LastRow, ColumnBefore + AO_COUNT)
            'YES where WINDOW_MONTHS months are entered and each of the last WINDOW_MONTHS is met
            If Valid >= WINDOW_MONTHS And RowOutput(LastRow, ColumnBefore + AO_COUNT) = 0 Then
                ResultOutput(a, BC_RESTART) = TXT_YES
            Else
                ResultOutput(a, BC_RESTART) = TXT_NO
            End If
        Next a

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function MonthStatus( _
    ByRef InputRows() As Variant, _
    ByVal RowIndex As Long, _
    ByVal Previous As Double) _
    As String
'
'------------------------------------------------------------------------------
'                              TEST ONE MONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the month status of one row: ST_OK, or the first test of CALC-09
'   that the row fails
'
' WHY THIS EXISTS
'   The tests of a month are made in the order of the contract, and the first
'   that fails gives the status
'
' INPUTS
'   InputRows
'     Eligibility cells (1 To n, 1 To EL_INPUT_WIDTH); read, not modified
'
'   RowIndex
'     Row to test (1 To n); its month end is not blank
'
'   Previous
'     Month end of the row entered before it as a date serial, or 0 when there
'     is none or it is not a date
'
' RETURNS
'   ST_OK, or ST_INVALID followed by the reason of the first test that fails
'
' BEHAVIOR
'   Tests in this order: the month end is a date; it is the last day of its
'   month; when Previous is a date, it is the last day of the month after the
'   month of Previous; the total assets are a number greater than 0; the four
'   long and short totals are numbers
'
' ERROR POLICY
'   Never raises on cell content. The procedure has no error handler
'
' DEPENDENCIES
'   MonthEndSerial; CORE_Util: IsDateSerial, IsNumberValue and NumberValue;
'   CORE_Config: the EL_ columns, ST_OK and ST_INVALID
'
'   Reference: calculation contract CALC-09, Failure behaviour
'
' NOTES
'   A month end with a time of day is not the last day of a month: the last
'   day is a whole serial
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MonthEnd    As Double   'Month end of the row as a date serial

'------------------------------------------------------------------------------
' TEST MONTH END
'------------------------------------------------------------------------------
    'The month end must be a date
        If Not IsDateSerial(InputRows(RowIndex, EL_MONTH)) Then
            MonthStatus = ST_INVALID & "MonthEnd must be a date"
            Exit Function
        End If
    'It must be the last day of its month
        MonthEnd = NumberValue(InputRows(RowIndex, EL_MONTH))
        If MonthEnd <> MonthEndSerial(MonthEnd, 0) Then
            MonthStatus = ST_INVALID & "MonthEnd must be the last day of a month"
            Exit Function
        End If
    'It must be the month end after the one of the row entered before it, when that is a date
        If Previous > 0 Then
            If MonthEnd <> MonthEndSerial(Previous, 1) Then
                MonthStatus = ST_INVALID & "MonthEnd must be one month after the previous row"
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' TEST AMOUNTS
'------------------------------------------------------------------------------
    'The total assets must be a number greater than 0
        If Not IsNumberValue(InputRows(RowIndex, EL_ASSETS)) Then
            MonthStatus = ST_INVALID & "TotalAssets must be a number greater than 0"
            Exit Function
        End If
        If NumberValue(InputRows(RowIndex, EL_ASSETS)) <= 0 Then
            MonthStatus = ST_INVALID & "TotalAssets must be a number greater than 0"
            Exit Function
        End If
    'The four long and short totals must be numbers
        If Not (IsNumberValue(InputRows(RowIndex, EL_LONG_325A)) And _
                IsNumberValue(InputRows(RowIndex, EL_SHORT_325A)) And _
                IsNumberValue(InputRows(RowIndex, EL_LONG_94)) And _
                IsNumberValue(InputRows(RowIndex, EL_SHORT_94))) Then
            MonthStatus = ST_INVALID & "a long or short total must be a number"
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Every test passed
        MonthStatus = ST_OK

End Function

Private Function MonthEndSerial( _
    ByVal Serial As Double, _
    ByVal Months As Long) _
    As Double
'
'------------------------------------------------------------------------------
'                            FIND LAST DAY OF MONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the serial of the last day of the month that lies Months months
'   after the month of a date serial, as Excel's EOMONTH does
'
' WHY THIS EXISTS
'   The core has no worksheet function, and the month after December 9999 has
'   no date: DateSerial alone would raise there
'
' INPUTS
'   Serial
'     A date serial from DATE_SERIAL_MIN to DATE_SERIAL_MAX; its time of day is
'     ignored
'
'   Months
'     Months to add, 0 or more
'
' RETURNS
'   The serial of the last day of that month, a whole number; -1 when the
'   month lies after December 9999
'
' BEHAVIOR
'   Counts the months from year 0 to the month of the date, adds Months, and
'   returns the day before the first of the next month, or 31 December for a
'   month of December
'
' ERROR POLICY
'   Never raises for a serial in its range. The procedure has no error handler
'
' DEPENDENCIES
'   DateSerial, Year, Month, CDate and CLng
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
    Dim WholeDay    As Date     'The date, without its time of day
    Dim MonthIndex  As Long     'Months from year 0 to the target month
    Dim TargetYear  As Long     'Year of the target month
    Dim TargetMonth As Long     'Month of the target month, 1 to 12

'------------------------------------------------------------------------------
' FIND TARGET MONTH
'------------------------------------------------------------------------------
    'Count the months of the date and add Months
        WholeDay = CDate(Int(Serial))
        MonthIndex = CLng(Year(WholeDay)) * 12 + CLng(Month(WholeDay)) - 1 + Months
        TargetYear = MonthIndex \ 12
        TargetMonth = MonthIndex Mod 12 + 1

'------------------------------------------------------------------------------
' RETURN LAST DAY
'------------------------------------------------------------------------------
    'No date lies after December 9999
        If TargetYear > 9999 Then
            MonthEndSerial = -1#
        ElseIf TargetMonth = 12 Then
            MonthEndSerial = CDbl(DateSerial(TargetYear, 12, 31))
        Else
            MonthEndSerial = CDbl(DateSerial(TargetYear, TargetMonth + 1, 1)) - 1#
        End If

End Function
