Attribute VB_Name = "CORE_Equity"
'==============================================================================
' MODULE: CORE_Equity
'------------------------------------------------------------------------------
' PURPOSE
'   Equity risk under the MAR40 simplified standardised approach: validate
'   each position row, net positions in the same stock or index of a
'   market, and calculate the specific, index and general market risk
'   charges. Returns the EQCalc table and the equity block of Buckets.
'
' PUBLIC SURFACE
'   None outside this VBA project (Option Private Module).
'     CalculateEquity.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util.
'
' STATE OWNERSHIP
'   None. Every procedure depends only on its arguments.
'
' ERROR POLICY
'   Cell content never raises: an unusable row gets an INVALID status and
'   contributes nothing. ERR_ARGUMENT is raised for arrays of the wrong
'   shape.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults. Never touches the Excel
'   object model.
'
' REFERENCE
'   BCBS MAR40.41-40.47.
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
    'Require explicit declarations; keep the procedures inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Columns of the EQCalc output, counted from column C.
        Private Const OC_ID          As Long = 1     'PositionID
        Private Const OC_STATUS      As Long = 2     'Status
        Private Const OC_MARKET      As Long = 3     'Market code
        Private Const OC_TYPE        As Long = 4     'Single name or Index
        Private Const OC_IDENTIFIER  As Long = 5     'Stock or index code
        Private Const OC_FX          As Long = 6     'FX rate
        Private Const OC_SIGNED      As Long = 7     'Signed value, reporting currency
        Private Const OC_NET_KEY     As Long = 8     'Market|Type|Identifier
        Private Const OC_NET_VALUE   As Long = 9     'Net value of the stock or index
        Private Const OC_CHARGE      As Long = 10    'Specific or index charge
        Private Const OC_MARKET_NO   As Long = 11    'Market number
        Private Const OC_ERROR_NO    As Long = 12    'Running number of rejected rows
        Private Const TYPE_SINGLE    As String = "SINGLE NAME"    'Normalised type
        Private Const TYPE_INDEX     As String = "INDEX"          'Normalised type


'------------------------------------------------------------------------------
'                                  CALCULATION
'------------------------------------------------------------------------------

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
'==============================================================================
'                                CalculateEquity
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate the equity charges for all positions and return every
'   intermediate.
'
' INPUTS
'   InputRows        EQPositions cells (1 To n, 1 To EQ_INPUT_WIDTH); a row whose
'               PositionID is blank is ignored. Values are delta-equivalent.
'   Parameters  Vector (1 To PI_COUNT); PI_EQ_SPECIFIC, PI_EQ_INDEX and
'               PI_EQ_GENERAL are used.
'   FxCodes     FX table codes, upper case; FxRates their rates (> 0).
'   GoldCode    Gold code, upper case; not a valid equity currency.
'
' RETURNS
'   RowOutput     (1 To n, 1 To EQ_CALC_WIDTH): EQCalc columns C:N.
'   MarketOutput  (1 To k + 1, 1 To BK_EQ_WIDTH): Buckets C:L, one row per
'                 market in order of first valid row, then a total row
'                 (k >= 0).
'   Totals        (1 To TOT_COUNT): TOT_SPECIFIC (single names),
'                 TOT_GENERAL, TOT_OTHER (index charge), TOT_USED,
'                 TOT_REJECTED, TOT_BUCKETS (markets).
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when InputRows is not EQ_INPUT_WIDTH wide.
'
' REFERENCE
'   MAR40.43 (8% specific and 8% general), MAR40.46 (identical equity or
'   index nets), MAR40.47 (further 2% on the net position per index).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim RowCount    As Long        'Input rows
    Dim Ids()       As String      'Distinct PositionIDs
    Dim IdCount     As Long        'Distinct PositionIDs in use
    Dim IdUses()    As Long        'InputRows per PositionID
    Dim NetKeys()   As String      'Distinct Market|Type|Identifier keys
    Dim NetCount    As Long        'Keys in use
    Dim NetValue()  As Double      'Net value per key
    Dim NetFirst()  As Long        'First row per key
    Dim Markets()   As String      'Distinct markets, in order
    Dim MarketCount As Long        'Markets in use
    Dim Status      As String      'Row status
    Dim Market      As String      'Row market code
    Dim TypeCode    As String      'Row type, normalised
    Dim Identifier  As String      'Row identifier code
    Dim CurrencyCode As String     'Row currency code
    Dim Position    As Long        'Position in a lookup
    Dim ErrorNumber As Long        'Running number of rejected rows
    Dim m           As Long        'Market
    Dim r           As Long        'Input row
    Dim c           As Long        'Output column

        If UBound(InputRows, 2) < EQ_INPUT_WIDTH Then
            Err.Raise ERR_ARGUMENT, "CORE_Equity.CalculateEquity", _
                "InputRows must be " & EQ_INPUT_WIDTH & " columns wide."
        End If
        RowCount = UBound(InputRows, 1)
        ReDim RowOutput(1 To RowCount, 1 To EQ_CALC_WIDTH)
        ReDim Totals(1 To TOT_COUNT)
        ReDim NetValue(1 To RowCount)
        ReDim NetFirst(1 To RowCount)

        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, EQ_ID)), True)
        Next r
        ReDim IdUses(1 To IdCount + 1)
        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, EQ_ID)), False)
            If Position > 0 Then IdUses(Position) = IdUses(Position) + 1
        Next r

        '--- Row pass ----------------------------------------------------------
        For r = 1 To RowCount
            If Len(PlainText(InputRows(r, EQ_ID))) > 0 Then
                Market = CodeText(InputRows(r, EQ_MARKET))
                TypeCode = CodeText(InputRows(r, EQ_TYPE))
                Identifier = CodeText(InputRows(r, EQ_IDENTIFIER))
                CurrencyCode = CodeText(InputRows(r, EQ_CCY))
                If IdUses(KeyPosition(Ids, IdCount, CodeText(InputRows(r, EQ_ID)), False)) > 1 Then
                    Status = ST_INVALID & "duplicate PositionID"
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

                RowOutput(r, OC_ID) = PlainText(InputRows(r, EQ_ID))
                RowOutput(r, OC_STATUS) = Status
                RowOutput(r, OC_MARKET) = Market
                RowOutput(r, OC_TYPE) = PlainText(InputRows(r, EQ_TYPE))
                RowOutput(r, OC_IDENTIFIER) = Identifier
                RowOutput(r, OC_SIGNED) = 0#
                RowOutput(r, OC_CHARGE) = 0#
                If Status = ST_OK Then
                    Position = FindCode(FxCodes, CurrencyCode)
                    RowOutput(r, OC_FX) = FxRates(Position)
                    RowOutput(r, OC_SIGNED) = DirectionSign(InputRows(r, EQ_DIRECTION)) _
                        * NumberValue(InputRows(r, EQ_VALUE)) * FxRates(Position)
                    RowOutput(r, OC_NET_KEY) = Market & "|" & PlainText(InputRows(r, EQ_TYPE)) & "|" & Identifier
                    Position = KeyPosition(NetKeys, NetCount, CStr(RowOutput(r, OC_NET_KEY)), True)
                    NetValue(Position) = NetValue(Position) + RowOutput(r, OC_SIGNED)
                    If NetFirst(Position) = 0 Then NetFirst(Position) = r
                    If KeyPosition(Markets, MarketCount, Market, False) = 0 Then
                        RowOutput(r, OC_MARKET_NO) = KeyPosition(Markets, MarketCount, Market, True)
                    End If
                    Totals(TOT_USED) = Totals(TOT_USED) + 1
                Else
                    ErrorNumber = ErrorNumber + 1
                    RowOutput(r, OC_ERROR_NO) = ErrorNumber
                    Totals(TOT_REJECTED) = Totals(TOT_REJECTED) + 1
                End If
            End If
        Next r

        '--- Net per stock or index; charge on the first row --------------------
        For r = 1 To RowCount
            If Len(CStr(RowOutput(r, OC_NET_KEY))) > 0 Then
                Position = KeyPosition(NetKeys, NetCount, CStr(RowOutput(r, OC_NET_KEY)), False)
                RowOutput(r, OC_NET_VALUE) = NetValue(Position)
                If NetFirst(Position) = r Then
                    If CodeText(RowOutput(r, OC_TYPE)) = TYPE_INDEX Then
                        RowOutput(r, OC_CHARGE) = Parameters(PI_EQ_INDEX) * Abs(NetValue(Position))
                        Totals(TOT_OTHER) = Totals(TOT_OTHER) + RowOutput(r, OC_CHARGE)
                    Else
                        RowOutput(r, OC_CHARGE) = Parameters(PI_EQ_SPECIFIC) * Abs(NetValue(Position))
                        Totals(TOT_SPECIFIC) = Totals(TOT_SPECIFIC) + RowOutput(r, OC_CHARGE)
                    End If
                End If
            End If
        Next r

        '--- Markets (Buckets) -------------------------------------------------
        ReDim MarketOutput(1 To MarketCount + 1, 1 To BK_EQ_WIDTH)
        For m = 1 To MarketCount
            MarketOutput(m, 1) = m
            MarketOutput(m, 2) = Markets(m)
            For c = 3 To BK_EQ_WIDTH
                MarketOutput(m, c) = 0#
            Next c
        Next m
        For r = 1 To RowCount
            If RowOutput(r, OC_STATUS) = ST_OK Then
                m = KeyPosition(Markets, MarketCount, CStr(RowOutput(r, OC_MARKET)), False)
                MarketOutput(m, 3) = MarketOutput(m, 3) + 1
                If RowOutput(r, OC_SIGNED) > 0 Then MarketOutput(m, 4) = MarketOutput(m, 4) + RowOutput(r, OC_SIGNED)
                If RowOutput(r, OC_SIGNED) < 0 Then MarketOutput(m, 5) = MarketOutput(m, 5) - RowOutput(r, OC_SIGNED)
                If CodeText(RowOutput(r, OC_TYPE)) = TYPE_INDEX Then
                    MarketOutput(m, 9) = MarketOutput(m, 9) + RowOutput(r, OC_CHARGE)
                Else
                    MarketOutput(m, 8) = MarketOutput(m, 8) + RowOutput(r, OC_CHARGE)
                End If
            End If
        Next r
        For m = 1 To MarketCount
            MarketOutput(m, 6) = MarketOutput(m, 4) - MarketOutput(m, 5)
            MarketOutput(m, 7) = Parameters(PI_EQ_GENERAL) * Abs(MarketOutput(m, 6))
            MarketOutput(m, 10) = MarketOutput(m, 7) + MarketOutput(m, 8) + MarketOutput(m, 9)
            Totals(TOT_GENERAL) = Totals(TOT_GENERAL) + MarketOutput(m, 7)
        Next m
        Totals(TOT_BUCKETS) = MarketCount
        WriteTotalRow MarketOutput, MarketCount, LBL_TOTAL_EQ

End Sub
