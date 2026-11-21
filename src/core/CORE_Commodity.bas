Attribute VB_Name = "CORE_Commodity"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Commodity
'------------------------------------------------------------------------------
' PURPOSE
'   Commodity risk under the MAR40 simplified approach: validate each position
'   row, value it at spot in reporting currency, and charge 15% of the net and
'   3% of the gross position per commodity. Returns the COCalc table and the
'   commodity block of Buckets
'
' WHY THIS EXISTS
'   It holds the commodity step of the calculation as one procedure that
'   takes its positions and parameters as arguments and returns its result
'   tables through arguments; it never touches the Excel object model
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Public surface: CalculateCommodity, which is visible only inside this VBA
'   project (Option Private Module). It validates the position rows, values
'   the valid ones in reporting currency, groups them by commodity and
'   computes the directional and gross charges of each commodity
'
' ERROR POLICY
'   Cell content never raises: an unusable row gets an INVALID status and
'   contributes nothing. ERR_ARGUMENT is raised for arrays of the wrong shape
'
' DEPENDENCIES
'   CORE_Config and CORE_Util
'
'   Reference: BCBS MAR40.64-40.73 (simplified approach: MAR40.72-40.73);
'   calculation contract CALC-05
'
' NOTES
'   Option Explicit requires explicit declarations and Option Private Module
'   keeps the procedures inside this project. The module owns no state: every
'   procedure depends only on its arguments
'
'   Compatibility: Excel VBA; no references beyond the defaults
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Columns of the COCalc output, counted from column C
    Private Const OC_ID             As Long = 1 'PositionID
    Private Const OC_STATUS         As Long = 2 'Status
    Private Const OC_COMMODITY      As Long = 3 'Commodity code
    Private Const OC_FX             As Long = 4 'FX rate
    Private Const OC_SIGNED         As Long = 5 'Signed value, reporting currency
    Private Const OC_GROSS          As Long = 6 'Gross value
    Private Const OC_COMMODITY_NO   As Long = 7 'Commodity number
    Private Const OC_ERROR_NO       As Long = 8 'Running number of rejected rows


'
'------------------------------------------------------------------------------
'
'                                 CALCULATION
'
'------------------------------------------------------------------------------
'

Public Sub CalculateCommodity( _
    ByRef InputRows() As Variant, _
    ByRef Parameters() As Double, _
    ByRef FxCodes() As String, _
    ByRef FxRates() As Double, _
    ByVal GoldCode As String, _
    ByRef RowOutput() As Variant, _
    ByRef CommodityOutput() As Variant, _
    ByRef Totals() As Double)
'
'------------------------------------------------------------------------------
'                             CALCULATE COMMODITY
'------------------------------------------------------------------------------
' PURPOSE
'   Calculates the commodity charges for all positions and returns every
'   intermediate
'
' WHY THIS EXISTS
'   The commodity step needs the row table, the commodity block of Buckets
'   and the totals from one pass over the same positions; it reads and
'   writes only its arguments
'
' INPUTS
'   InputRows
'     COPositions cells (1 To n, 1 To CO_INPUT_WIDTH); a row whose
'     PositionID is blank is ignored. Quantities are delta-equivalent; any
'     sign of spot price is accepted
'
'   Parameters
'     Vector (1 To PI_COUNT); PI_CO_DIRECTIONAL and PI_CO_GROSS are used
'
'   FxCodes
'     FX table codes, upper case
'
'   FxRates
'     Rates of the FxCodes, one per code and > 0, used to value a position
'     in reporting currency
'
'   GoldCode
'     Gold code, upper case; not a valid price currency
'
'   RowOutput
'     Output; dimensioned by the procedure, see RETURNS
'
'   CommodityOutput
'     Output; dimensioned by the procedure, see RETURNS
'
'   Totals
'     Output; dimensioned by the procedure, see RETURNS
'
' RETURNS
'   Nothing; the results are written to the output arguments:
'     RowOutput        (1 To n, 1 To CO_CALC_WIDTH): COCalc columns C:J
'     CommodityOutput  (1 To k + 1, 1 To BK_CO_WIDTH): Buckets C:L, one row
'                      per commodity in order of first valid row, then a
'                      total row (k >= 0)
'     Totals           (1 To TOT_COUNT): TOT_GENERAL (directional),
'                      TOT_OTHER (gross charge), TOT_USED, TOT_REJECTED,
'                      TOT_BUCKETS (commodities)
'
' BEHAVIOR
'   Validates the width of InputRows, then in order:
'     - counts the uses of each PositionID
'     - classifies each row with a non-blank PositionID as OK or INVALID and
'       values an OK row in reporting currency (direction sign, quantity,
'       spot price and FX rate)
'     - groups the OK rows by commodity, in order of first valid row
'     - computes the directional charge (PI_CO_DIRECTIONAL times the absolute
'       net position) and the gross charge (PI_CO_GROSS times the gross
'       position) of each commodity, and the total row
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when InputRows is not CO_INPUT_WIDTH wide. A row whose
'   commodity is the gold code is rejected: gold is FX risk
'
' DEPENDENCIES
'   CORE_Util: KeyPosition, CodeText, PlainText, DirectionSign, BookStatus,
'   IsNumberValue, NumberValue, FindCode and WriteTotalRow; CORE_Config: the
'   column, width, index, status, label and error constants
'
'   Reference: MAR40.72 (15% of the net position per commodity), MAR40.73
'   (3% of the gross position, valued at spot), MAR40.63 (gold is FX risk);
'   calculation contract CALC-05
'
' NOTES
'   A non-trading book position is calculated as a trading book one
'   (Article 325(1) of the CRR; BCBS MAR12.6). A row is rejected for the first
'   test that fails, in this order: duplicate PositionID, Book not Trading,
'   Non-trading or blank, missing commodity, gold code, Direction not Long or
'   Short, Quantity not a number, Quantity < 0, spot price not a number,
'   currency not in the FX table or equal to the gold code
'
'   CommodityOutput columns: 1 number, 2 commodity, 3 valid rows, 4 sum of
'   positive signed values, 5 sum of negative signed values as a positive
'   amount, 6 net (4 less 5), 7 gross, 8 directional charge, 9 gross charge,
'   10 sum of 8 and 9
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RowCount            As Long     'Input rows
    Dim Ids()               As String   'Distinct PositionIDs
    Dim IdCount             As Long     'Distinct PositionIDs in use
    Dim IdUses()            As Long     'InputRows per PositionID
    Dim CommodityNames()    As String   'Distinct commodities, in order
    Dim NameCount           As Long     'Commodities in use
    Dim Status              As String   'Row status
    Dim Commodity           As String   'Row commodity code
    Dim CurrencyCode        As String   'Row currency code
    Dim Position            As Long     'Position in a lookup
    Dim ErrorNumber         As Long     'Running number of rejected rows
    Dim k                   As Long     'Commodity
    Dim r                   As Long     'Input row
    Dim c                   As Long     'Output column

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an input array narrower than the COPositions layout
        If UBound(InputRows, 2) < CO_INPUT_WIDTH Then
            Err.Raise ERR_ARGUMENT, "CORE_Commodity.CalculateCommodity", _
                "InputRows must be " & CO_INPUT_WIDTH & " columns wide."
        End If

'------------------------------------------------------------------------------
' INITIALIZE OUTPUTS
'------------------------------------------------------------------------------
    'Count the input rows
        RowCount = UBound(InputRows, 1)
    'Size the row output and the totals vector
        ReDim RowOutput(1 To RowCount, 1 To CO_CALC_WIDTH)
        ReDim Totals(1 To TOT_COUNT)

'------------------------------------------------------------------------------
' COUNT POSITION IDS
'------------------------------------------------------------------------------
    'Register every distinct PositionID
        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, CO_ID)), True)
        Next r
    'Size the use counters for the registered PositionIDs
        ReDim IdUses(1 To IdCount + 1)
    'Count the input rows that use each PositionID
        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, CO_ID)), False)
            If Position > 0 Then IdUses(Position) = IdUses(Position) + 1
        Next r

'------------------------------------------------------------------------------
' EVALUATE ROWS
'------------------------------------------------------------------------------
    'Classify and value each row that has a PositionID
        For r = 1 To RowCount
            If Len(PlainText(InputRows(r, CO_ID))) > 0 Then
                'Read the commodity and currency codes of the row
                Commodity = CodeText(InputRows(r, CO_COMMODITY))
                CurrencyCode = CodeText(InputRows(r, CO_CCY))
                'Classify the row: the first failing test gives the rejection reason
                If IdUses(KeyPosition(Ids, IdCount, CodeText(InputRows(r, CO_ID)), False)) > 1 Then
                    Status = ST_INVALID & "duplicate PositionID"
                ElseIf Len(BookStatus(InputRows(r, CO_BOOK), True)) > 0 Then
                    Status = BookStatus(InputRows(r, CO_BOOK), True)
                ElseIf Len(Commodity) = 0 Then
                    Status = ST_INVALID & "Commodity missing"
                ElseIf Commodity = GoldCode Then
                    Status = ST_INVALID & "gold is FX risk; enter it on FXPositions"
                ElseIf DirectionSign(InputRows(r, CO_DIRECTION)) = 0 Then
                    Status = ST_INVALID & "Direction must be Long or Short"
                ElseIf Not IsNumberValue(InputRows(r, CO_QUANTITY)) Then
                    Status = ST_INVALID & "Quantity must be a number"
                ElseIf NumberValue(InputRows(r, CO_QUANTITY)) < 0 Then
                    Status = ST_INVALID & "Quantity must be >= 0 (Direction gives the sign)"
                ElseIf Not IsNumberValue(InputRows(r, CO_PRICE)) Then
                    Status = ST_INVALID & "Spot price must be a number"
                ElseIf FindCode(FxCodes, CurrencyCode) < 1 Or CurrencyCode = GoldCode Then
                    Status = ST_INVALID & "currency not in FX_Table"
                Else
                    Status = ST_OK
                End If
                'Write the identity, status and commodity of the row, with zero values until it is valued
                RowOutput(r, OC_ID) = PlainText(InputRows(r, CO_ID))
                RowOutput(r, OC_STATUS) = Status
                RowOutput(r, OC_COMMODITY) = Commodity
                RowOutput(r, OC_SIGNED) = 0#
                RowOutput(r, OC_GROSS) = 0#
                'Value a valid row; number and count a rejected one
                If Status = ST_OK Then
                    'Look up the FX rate of the price currency
                    Position = FindCode(FxCodes, CurrencyCode)
                    RowOutput(r, OC_FX) = FxRates(Position)
                    'Signed value in reporting currency: direction sign times quantity, spot price and FX rate
                    RowOutput(r, OC_SIGNED) = DirectionSign(InputRows(r, CO_DIRECTION)) _
                        * NumberValue(InputRows(r, CO_QUANTITY)) * NumberValue(InputRows(r, CO_PRICE)) _
                        * FxRates(Position)
                    'Gross value is the absolute signed value
                    RowOutput(r, OC_GROSS) = Abs(RowOutput(r, OC_SIGNED))
                    'Number the commodity the first time a valid row shows it
                    If KeyPosition(CommodityNames, NameCount, Commodity, False) = 0 Then
                        RowOutput(r, OC_COMMODITY_NO) = KeyPosition(CommodityNames, NameCount, Commodity, True)
                    End If
                    'Count the row as used
                    Totals(TOT_USED) = Totals(TOT_USED) + 1
                Else
                    'Number the rejected row
                    ErrorNumber = ErrorNumber + 1
                    RowOutput(r, OC_ERROR_NO) = ErrorNumber
                    'Count the row as rejected
                    Totals(TOT_REJECTED) = Totals(TOT_REJECTED) + 1
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' BUILD COMMODITY TABLE
'------------------------------------------------------------------------------
    'Size the commodity table with one row per commodity and a total row
        ReDim CommodityOutput(1 To NameCount + 1, 1 To BK_CO_WIDTH)
    'Number and name each commodity and clear its figures
        For k = 1 To NameCount
            CommodityOutput(k, 1) = k
            CommodityOutput(k, 2) = CommodityNames(k)
            For c = 3 To BK_CO_WIDTH
                CommodityOutput(k, c) = 0#
            Next c
        Next k

'------------------------------------------------------------------------------
' ACCUMULATE POSITIONS
'------------------------------------------------------------------------------
    'Add each valid row to its commodity
        For r = 1 To RowCount
            If RowOutput(r, OC_STATUS) = ST_OK Then
                'Find the commodity of the row
                k = KeyPosition(CommodityNames, NameCount, CStr(RowOutput(r, OC_COMMODITY)), False)
                'Count the valid row
                CommodityOutput(k, 3) = CommodityOutput(k, 3) + 1
                'Add a positive signed value to the long column, a negative one to the short column
                If RowOutput(r, OC_SIGNED) > 0 Then CommodityOutput(k, 4) = CommodityOutput(k, 4) + _
                    RowOutput(r, OC_SIGNED)
                If RowOutput(r, OC_SIGNED) < 0 Then CommodityOutput(k, 5) = CommodityOutput(k, 5) - _
                    RowOutput(r, OC_SIGNED)
                'Add the gross value
                CommodityOutput(k, 7) = CommodityOutput(k, 7) + RowOutput(r, OC_GROSS)
            End If
        Next r

'------------------------------------------------------------------------------
' CALCULATE CHARGES
'------------------------------------------------------------------------------
    'Derive the net position and the charges of each commodity and add them to the totals
        For k = 1 To NameCount
            'Net position: long less short
            CommodityOutput(k, 6) = CommodityOutput(k, 4) - CommodityOutput(k, 5)
            'Directional charge: parameter times the absolute net position
            CommodityOutput(k, 8) = Parameters(PI_CO_DIRECTIONAL) * Abs(CommodityOutput(k, 6))
            'Gross charge: parameter times the gross position
            CommodityOutput(k, 9) = Parameters(PI_CO_GROSS) * CommodityOutput(k, 7)
            'Commodity charge: directional plus gross
            CommodityOutput(k, 10) = CommodityOutput(k, 8) + CommodityOutput(k, 9)
            'Add the commodity charges to the totals
            Totals(TOT_GENERAL) = Totals(TOT_GENERAL) + CommodityOutput(k, 8)
            Totals(TOT_OTHER) = Totals(TOT_OTHER) + CommodityOutput(k, 9)
        Next k

'------------------------------------------------------------------------------
' WRITE TOTAL ROW
'------------------------------------------------------------------------------
    'Record the number of commodities
        Totals(TOT_BUCKETS) = NameCount
    'Write the total row of the commodity block
        WriteTotalRow CommodityOutput, NameCount, LBL_TOTAL_CO

End Sub
