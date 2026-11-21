Attribute VB_Name = "CORE_Equity"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Equity
'------------------------------------------------------------------------------
' PURPOSE
'   Equity risk under the MAR40 simplified standardised approach: validate
'   each position row, net positions in the same stock or index of a
'   market, and calculate the specific, index and general market risk
'   charges. Returns the EQCalc table and the equity block of Buckets
'
' WHY THIS EXISTS
'   It keeps the whole equity calculation in one procedure that works only
'   on the arrays it receives, so the result does not depend on any
'   worksheet or workbook state
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Holds one procedure, visible only inside this VBA project:
'     - CalculateEquity: validates the position rows, nets them per stock or
'       index, and returns the EQCalc rows, the Buckets equity rows and the
'       totals
'
' ERROR POLICY
'   Cell content never raises: an unusable row gets an INVALID status and
'   contributes nothing. ERR_ARGUMENT is raised for arrays of the wrong
'   shape
'
' DEPENDENCIES
'   CORE_Config, CORE_Util
'
'   Reference: BCBS MAR40.41-40.47
'
' NOTES
'   Require explicit declarations (Option Explicit) and keep the procedures
'   inside this project (Option Private Module). The module owns no state:
'   every procedure depends only on its arguments
'
'   Compatibility: Excel VBA; no references beyond the defaults. Never
'   touches the Excel object model
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Columns of the EQCalc output, counted from column C
    Private Const OC_ID             As Long = 1                 'PositionID
    Private Const OC_STATUS         As Long = 2                 'Status
    Private Const OC_MARKET         As Long = 3                 'Market code
    Private Const OC_TYPE           As Long = 4                 'Single name or Index
    Private Const OC_IDENTIFIER     As Long = 5                 'Stock or index code
    Private Const OC_FX             As Long = 6                 'FX rate
    Private Const OC_SIGNED         As Long = 7                 'Signed value, reporting currency
    Private Const OC_NET_KEY        As Long = 8                 'Market|Type|Identifier
    Private Const OC_NET_VALUE      As Long = 9                 'Net value of the stock or index
    Private Const OC_CHARGE         As Long = 10                'Specific or index charge
    Private Const OC_MARKET_NO      As Long = 11                'Market number
    Private Const OC_ERROR_NO       As Long = 12                'Running number of rejected rows

    'Normalised values of the Type column
    Private Const TYPE_SINGLE       As String = "SINGLE NAME"   'Normalised type of a single name
    Private Const TYPE_INDEX        As String = "INDEX"         'Normalised type of an index


'
'------------------------------------------------------------------------------
'
'                                 CALCULATION
'
'------------------------------------------------------------------------------
'

Public Sub CalculateEquity( _
    ByRef InputRows() As Variant, _
    ByRef Parameters() As Double, _
    ByRef FxCodes() As String, _
    ByRef FxRates() As Double, _
    ByVal GoldCode As String, _
    ByRef RowOutput() As Variant, _
    ByRef MarketOutput() As Variant, _
    ByRef Totals() As Double)
'
'------------------------------------------------------------------------------
'                               CALCULATE EQUITY
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate the equity charges for all positions and return every
'   intermediate
'
' WHY THIS EXISTS
'   The workbook layer needs the EQCalc rows, the equity block of Buckets
'   and the totals from one call, computed from the arrays it passes and
'   from nothing else
'
' INPUTS
'   InputRows
'     EQPositions cells (1 To n, 1 To EQ_INPUT_WIDTH); a row whose
'     PositionID is blank is ignored. Values are delta-equivalent
'
'   Parameters
'     Vector (1 To PI_VECTOR); PI_EQ_SPECIFIC, PI_EQ_INDEX and
'     PI_EQ_GENERAL are used, and PI_EQ_INDEX_CRR in place of
'     PI_EQ_INDEX where PI_REGIME_FLAG is 1
'
'   FxCodes
'     FX table codes, upper case
'
'   FxRates
'     Rates of the FxCodes entries (> 0)
'
'   GoldCode
'     Gold code, upper case; not a valid equity currency
'
'   RowOutput, MarketOutput, Totals
'     Output arrays, resized and filled here; see RETURNS
'
' RETURNS
'   Nothing; the results are written to the output arguments:
'     - RowOutput: (1 To n, 1 To EQ_CALC_WIDTH), EQCalc columns C:N
'     - MarketOutput: (1 To k + 1, 1 To BK_EQ_WIDTH), Buckets C:L, one row
'       per market in order of first valid row, then a total row (k >= 0)
'     - Totals: (1 To TOT_COUNT): TOT_SPECIFIC (single names), TOT_GENERAL,
'       TOT_OTHER (index charge), TOT_USED, TOT_REJECTED, TOT_EXCLUDED
'       (non-trading book rows), TOT_BUCKETS (markets)
'
' BEHAVIOR
'   Works in this order:
'     - checks the width of InputRows
'     - counts the uses of each PositionID
'     - validates each row, the book first after the PositionID
'       (BookStatus: a non-trading book row is EXCLUDED), converts a valid row to the reporting currency
'       with its signed value and adds it to the net of its stock or index
'     - charges each net position once, on its first row: the index rate
'       for an index, PI_EQ_SPECIFIC for a single name
'     - aggregates per market and applies PI_EQ_GENERAL to the net position
'       of the market
'     - writes the total row of MarketOutput
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when InputRows is not EQ_INPUT_WIDTH wide. A row that
'   fails validation gets an INVALID status, is counted in TOT_REJECTED and
'   contributes nothing; it never raises. A non-trading book row gets an
'   EXCLUDED status, is counted in TOT_EXCLUDED and contributes nothing
'
' DEPENDENCIES
'   CodeText, PlainText, KeyPosition, FindCode, DirectionSign, BookStatus,
'   IsInvalidStatus, IsNumberValue, NumberValue and WriteTotalRow
'
'   Reference: MAR40.43 (8% specific and 8% general), MAR40.46 (identical
'   equity or index nets), MAR40.47 (further 2% on the net position per
'   index). Under the CRR regime, Article 344: an index future is one
'   equity, its general risk is in the market's net position, and its
'   specific risk is ignored where it is exchange traded and represents a
'   relevant, appropriately diversified index, so the charge on an index is
'   PI_EQ_INDEX_CRR, 0
'
' NOTES
'   Owns no state; writes only RowOutput, MarketOutput and Totals
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RowCount        As Long     'Input rows
    Dim IndexRate       As Double   'Charge on an index of the regime
    Dim Ids()           As String   'Distinct PositionIDs
    Dim IdCount         As Long     'Distinct PositionIDs in use
    Dim IdUses()        As Long     'InputRows per PositionID
    Dim NetKeys()       As String   'Distinct Market|Type|Identifier keys
    Dim NetCount        As Long     'Keys in use
    Dim NetValue()      As Double   'Net value per key
    Dim NetFirst()      As Long     'First row per key
    Dim Markets()       As String   'Distinct markets, in order
    Dim MarketCount     As Long     'Markets in use
    Dim Status          As String   'Row status
    Dim Market          As String   'Row market code
    Dim TypeCode        As String   'Row type, normalised
    Dim Identifier      As String   'Row identifier code
    Dim CurrencyCode    As String   'Row currency code
    Dim Position        As Long     'Position in a lookup
    Dim ErrorNumber     As Long     'Running number of rejected rows
    Dim m               As Long     'Market
    Dim r               As Long     'Input row
    Dim c               As Long     'Output column

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an InputRows array narrower than the EQPositions layout
        If UBound(InputRows, 2) < EQ_INPUT_WIDTH Then
            Err.Raise ERR_ARGUMENT, "CORE_Equity.CalculateEquity", _
                "InputRows must be " & EQ_INPUT_WIDTH & " columns wide."
        End If

'------------------------------------------------------------------------------
' SIZE OUTPUTS
'------------------------------------------------------------------------------
    'Take the number of input rows from the first dimension
        RowCount = UBound(InputRows, 1)
    'Size the row output, the totals and the netting arrays
        ReDim RowOutput(1 To RowCount, 1 To EQ_CALC_WIDTH)
        ReDim Totals(1 To TOT_COUNT)
        ReDim NetValue(1 To RowCount)
        ReDim NetFirst(1 To RowCount)

'------------------------------------------------------------------------------
' COUNT POSITION IDS
'------------------------------------------------------------------------------
    'Register each PositionID in the table of distinct IDs
        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, EQ_ID)), True)
        Next r
    'Size the use counters, then count how many rows carry each PositionID
        ReDim IdUses(1 To IdCount + 1)
        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, EQ_ID)), False)
            If Position > 0 Then IdUses(Position) = IdUses(Position) + 1
        Next r

'------------------------------------------------------------------------------
' VALIDATE AND NET ROWS
'------------------------------------------------------------------------------
    'Validate each non-blank row, write its output and add a valid row to the net of its stock or index
        For r = 1 To RowCount
            If Len(PlainText(InputRows(r, EQ_ID))) > 0 Then
                'Read the normalised key fields of the row
                Market = CodeText(InputRows(r, EQ_MARKET))
                TypeCode = CodeText(InputRows(r, EQ_TYPE))
                Identifier = CodeText(InputRows(r, EQ_IDENTIFIER))
                CurrencyCode = CodeText(InputRows(r, EQ_CCY))
                'Take the first validation rule that the row breaks as its status
                If IdUses(KeyPosition(Ids, IdCount, CodeText(InputRows(r, EQ_ID)), False)) > 1 Then
                    Status = ST_INVALID & "duplicate PositionID"
                ElseIf Len(BookStatus(InputRows(r, EQ_BOOK), False)) > 0 Then
                    Status = BookStatus(InputRows(r, EQ_BOOK), False)
                ElseIf Len(Market) = 0 Then
                    Status = ST_INVALID & "Market missing"
                ElseIf TypeCode <> TYPE_SINGLE And TypeCode <> TYPE_INDEX Then
                    Status = ST_INVALID & "Type must be Single name or Index"
                ElseIf Len(Identifier) = 0 Then
                    Status = ST_INVALID & "stock or index identifier missing"
                ElseIf FindCode(FxCodes, CurrencyCode) < 1 Or CurrencyCode = GoldCode Then
                    Status = ST_INVALID & "currency not in FX_Table"
                ElseIf DirectionSign(InputRows(r, EQ_DIRECTION)) = 0 Then
                    Status = ST_INVALID & "Direction must be Long or Short"
                ElseIf Not IsNumberValue(InputRows(r, EQ_VALUE)) Then
                    Status = ST_INVALID & "Market value must be a number"
                ElseIf NumberValue(InputRows(r, EQ_VALUE)) < 0 Then
                    Status = ST_INVALID & "Market value must be >= 0 (Direction gives the sign)"
                Else
                    Status = ST_OK
                End If

                'Write the identifying columns of the row
                RowOutput(r, OC_ID) = PlainText(InputRows(r, EQ_ID))
                RowOutput(r, OC_STATUS) = Status
                RowOutput(r, OC_MARKET) = Market
                RowOutput(r, OC_TYPE) = PlainText(InputRows(r, EQ_TYPE))
                RowOutput(r, OC_IDENTIFIER) = Identifier
                RowOutput(r, OC_SIGNED) = 0#
                RowOutput(r, OC_CHARGE) = 0#
                'Net a valid row and count it; number and count a rejected row
                If Status = ST_OK Then
                    'Convert the signed value to the reporting currency
                    Position = FindCode(FxCodes, CurrencyCode)
                    RowOutput(r, OC_FX) = FxRates(Position)
                    RowOutput(r, OC_SIGNED) = DirectionSign(InputRows(r, EQ_DIRECTION)) _
                        * NumberValue(InputRows(r, EQ_VALUE)) * FxRates(Position)
                    'Build the netting key and add the signed value to the net of that key
                    RowOutput(r, OC_NET_KEY) = Market & "|" & PlainText(InputRows(r, EQ_TYPE)) & "|" & Identifier
                    Position = KeyPosition(NetKeys, NetCount, CStr(RowOutput(r, OC_NET_KEY)), True)
                    NetValue(Position) = NetValue(Position) + RowOutput(r, OC_SIGNED)
                    'Remember the first row of the key, which carries its charge
                    If NetFirst(Position) = 0 Then NetFirst(Position) = r
                    'Register the market the first time a valid row uses it
                    If KeyPosition(Markets, MarketCount, Market, False) = 0 Then
                        RowOutput(r, OC_MARKET_NO) = KeyPosition(Markets, MarketCount, Market, True)
                    End If
                    'Count the row as used
                    Totals(TOT_USED) = Totals(TOT_USED) + 1
                ElseIf IsInvalidStatus(Status) Then
                    'Number the rejected row and count it
                    ErrorNumber = ErrorNumber + 1
                    RowOutput(r, OC_ERROR_NO) = ErrorNumber
                    Totals(TOT_REJECTED) = Totals(TOT_REJECTED) + 1
                Else
                    'Count the excluded row
                    Totals(TOT_EXCLUDED) = Totals(TOT_EXCLUDED) + 1
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' CALCULATE POSITION CHARGES
'------------------------------------------------------------------------------
    'Write the net of each key on its rows and the charge on the first row of the key
        For r = 1 To RowCount
            If Len(CStr(RowOutput(r, OC_NET_KEY))) > 0 Then
                'Look up the net position of the key
                Position = KeyPosition(NetKeys, NetCount, CStr(RowOutput(r, OC_NET_KEY)), False)
                RowOutput(r, OC_NET_VALUE) = NetValue(Position)
                If NetFirst(Position) = r Then
                    'Charge an index at the rate of the regime and add it to the index total
                    If CodeText(RowOutput(r, OC_TYPE)) = TYPE_INDEX Then
                        IndexRate = Parameters(PI_EQ_INDEX)
                        If Parameters(PI_REGIME_FLAG) = 1# Then IndexRate = Parameters(PI_EQ_INDEX_CRR)
                        RowOutput(r, OC_CHARGE) = IndexRate * Abs(NetValue(Position))
                        Totals(TOT_OTHER) = Totals(TOT_OTHER) + RowOutput(r, OC_CHARGE)
                    Else
                        'Charge a single name at the specific rate and add it to the specific total
                        RowOutput(r, OC_CHARGE) = Parameters(PI_EQ_SPECIFIC) * Abs(NetValue(Position))
                        Totals(TOT_SPECIFIC) = Totals(TOT_SPECIFIC) + RowOutput(r, OC_CHARGE)
                    End If
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' INITIALIZE MARKET ROWS
'------------------------------------------------------------------------------
    'Size the market rows and give each its number, its code and zero figures
        ReDim MarketOutput(1 To MarketCount + 1, 1 To BK_EQ_WIDTH)
        For m = 1 To MarketCount
            MarketOutput(m, 1) = m
            MarketOutput(m, 2) = Markets(m)
            For c = 3 To BK_EQ_WIDTH
                MarketOutput(m, c) = 0#
            Next c
        Next m

'------------------------------------------------------------------------------
' ACCUMULATE MARKET FIGURES
'------------------------------------------------------------------------------
    'Add each valid row to its market: count, long and short value, index or specific charge
        For r = 1 To RowCount
            If RowOutput(r, OC_STATUS) = ST_OK Then
                'Find the market of the row
                m = KeyPosition(Markets, MarketCount, CStr(RowOutput(r, OC_MARKET)), False)
                'Count the row and add its value to the long or the short side
                MarketOutput(m, 3) = MarketOutput(m, 3) + 1
                If RowOutput(r, OC_SIGNED) > 0 Then MarketOutput(m, 4) = MarketOutput(m, 4) + RowOutput(r, OC_SIGNED)
                If RowOutput(r, OC_SIGNED) < 0 Then MarketOutput(m, 5) = MarketOutput(m, 5) - RowOutput(r, OC_SIGNED)
                'Add the row charge to the index or to the specific column
                If CodeText(RowOutput(r, OC_TYPE)) = TYPE_INDEX Then
                    MarketOutput(m, 9) = MarketOutput(m, 9) + RowOutput(r, OC_CHARGE)
                Else
                    MarketOutput(m, 8) = MarketOutput(m, 8) + RowOutput(r, OC_CHARGE)
                End If
            End If
        Next r

'------------------------------------------------------------------------------
' CALCULATE MARKET CHARGES
'------------------------------------------------------------------------------
    'Derive the net position of each market, its general charge and its total
        For m = 1 To MarketCount
            MarketOutput(m, 6) = MarketOutput(m, 4) - MarketOutput(m, 5)
            MarketOutput(m, 7) = Parameters(PI_EQ_GENERAL) * Abs(MarketOutput(m, 6))
            MarketOutput(m, 10) = MarketOutput(m, 7) + MarketOutput(m, 8) + MarketOutput(m, 9)
            Totals(TOT_GENERAL) = Totals(TOT_GENERAL) + MarketOutput(m, 7)
        Next m

'------------------------------------------------------------------------------
' RETURN TOTALS
'------------------------------------------------------------------------------
    'Record the number of markets
        Totals(TOT_BUCKETS) = MarketCount
    'Write the total row of the market output
        WriteTotalRow MarketOutput, MarketCount, LBL_TOTAL_EQ

End Sub
