Attribute VB_Name = "CORE_ForeignExchange"
'==============================================================================
' MODULE: CORE_ForeignExchange
'------------------------------------------------------------------------------
' PURPOSE
'   Foreign exchange and gold risk under the MAR40 shorthand method:
'   validate each position row, net the positions per currency, and charge
'   8% of the overall net open position. Returns the FXCalc table and the
'   FX block of Buckets.
'
' PUBLIC SURFACE
'   None outside this VBA project (Option Private Module).
'     CalculateForeignExchange, NetOpenPositionCharge, DeMinimisWaived.
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
'   BCBS MAR40.53-40.61 (shorthand method: MAR40.60-40.61, Table 9).
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
    'Columns of the FXCalc output, counted from column C.
        Private Const OC_ID          As Long = 1     'PositionID
        Private Const OC_STATUS      As Long = 2     'Status
        Private Const OC_CCY         As Long = 3     'Currency code
        Private Const OC_FX          As Long = 4     'FX rate
        Private Const OC_AMOUNT      As Long = 5     'Signed amount, currency units
        Private Const OC_VALUE       As Long = 6     'Signed value, reporting currency
        Private Const OC_ERROR_NO    As Long = 7     'Running number of rejected rows


'------------------------------------------------------------------------------
'                                  CALCULATION
'------------------------------------------------------------------------------

Public Sub CalculateForeignExchange( _
    ByRef InputRows() As Variant, _
    ByRef Parameters() As Double, _
    ByRef FxCodes() As String, _
    ByRef FxRates() As Double, _
    ByVal ReportingCcy As String, _
    ByVal GoldCode As String, _
    ByRef RowOutput() As Variant, _
    ByRef CurrencyOutput() As Variant, _
    ByRef Totals() As Double)
'
'==============================================================================
'                           CalculateForeignExchange
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate the FX and gold charge for all positions and return every
'   intermediate.
'
' INPUTS
'   InputRows          FXPositions cells (1 To n, 1 To FX_INPUT_WIDTH); a row
'                 whose PositionID is blank is ignored. Amounts in units of
'                 the currency; gold in the unit its FX rate is quoted for.
'   Parameters    Vector (1 To PI_COUNT); PI_FX_CHARGE is used.
'   FxCodes       FX table codes (1 To m), upper case, "" for an empty slot;
'                 FxRates their rates (> 0).
'   ReportingCcy  Reporting currency, upper case: its rows are EXCLUDED.
'   GoldCode      Gold code, upper case.
'
' RETURNS
'   RowOutput       (1 To n, 1 To FX_CALC_WIDTH): FXCalc columns C:I.
'   CurrencyOutput  (1 To m + 1, 1 To BK_FX_WIDTH): Buckets C:K, one row per
'                   FX table slot, then a total row.
'   Totals          (1 To TOT_COUNT): TOT_LONG, TOT_SHORT, TOT_GOLD,
'                   TOT_ONOP, TOT_GENERAL (the charge), TOT_USED,
'                   TOT_REJECTED, TOT_EXCLUDED.
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when InputRows is not FX_INPUT_WIDTH wide.
'
' REFERENCE
'   MAR40.60 (greater of the sum of net shorts and the sum of net longs,
'   plus the net gold position regardless of sign), MAR40.61 (8%).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim RowCount    As Long        'Input rows
    Dim SlotCount   As Long        'FX table slots
    Dim Ids()       As String      'Distinct PositionIDs
    Dim IdCount     As Long        'Distinct PositionIDs in use
    Dim IdUses()    As Long        'InputRows per PositionID
    Dim Status      As String      'Row status
    Dim CurrencyCode As String     'Row currency code
    Dim Position    As Long        'Position in a lookup
    Dim ErrorNumber As Long        'Running number of rejected rows
    Dim r           As Long        'Input row
    Dim s           As Long        'FX slot

        If UBound(InputRows, 2) < FX_INPUT_WIDTH Then
            Err.Raise ERR_ARGUMENT, "CORE_ForeignExchange.CalculateForeignExchange", _
                "InputRows must be " & FX_INPUT_WIDTH & " columns wide."
        End If
        RowCount = UBound(InputRows, 1)
        SlotCount = UBound(FxCodes)
        ReDim RowOutput(1 To RowCount, 1 To FX_CALC_WIDTH)
        ReDim CurrencyOutput(1 To SlotCount + 1, 1 To BK_FX_WIDTH)
        ReDim Totals(1 To TOT_COUNT)

        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, FX_ID)), True)
        Next r
        ReDim IdUses(1 To IdCount + 1)
        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, FX_ID)), False)
            If Position > 0 Then IdUses(Position) = IdUses(Position) + 1
        Next r

        '--- Slots of the FX table (Buckets FX block) ---------------------------
        For s = 1 To SlotCount
            CurrencyOutput(s, 1) = s
            CurrencyOutput(s, 2) = FxCodes(s)
            If Len(FxCodes(s)) > 0 Then
                If FxCodes(s) = ReportingCcy Then
                    CurrencyOutput(s, 3) = FX_REPORTING
                ElseIf FxCodes(s) = GoldCode Then
                    CurrencyOutput(s, 3) = FX_GOLD_LABEL
                Else
                    CurrencyOutput(s, 3) = FX_CURRENCY
                End If
                CurrencyOutput(s, 5) = FxRates(s)
            End If
            CurrencyOutput(s, 4) = 0#
            CurrencyOutput(s, 6) = 0#
        Next s

        '--- Row pass ----------------------------------------------------------
        For r = 1 To RowCount
            If Len(PlainText(InputRows(r, FX_ID))) > 0 Then
                CurrencyCode = CodeText(InputRows(r, FX_CCY))
                Position = FindCode(FxCodes, CurrencyCode)
                If IdUses(KeyPosition(Ids, IdCount, CodeText(InputRows(r, FX_ID)), False)) > 1 Then
                    Status = ST_INVALID & "duplicate PositionID"
                ElseIf Position < 1 Then
                    Status = ST_INVALID & "currency not in FX_Table"
                ElseIf DirectionSign(InputRows(r, FX_DIRECTION)) = 0 Then
                    Status = ST_INVALID & "Direction must be Long or Short"
                ElseIf Not IsNumberValue(InputRows(r, FX_AMOUNT)) Then
                    Status = ST_INVALID & "Amount must be a number"
                ElseIf NumberValue(InputRows(r, FX_AMOUNT)) < 0 Then
                    Status = ST_INVALID & "Amount must be >= 0 (Direction gives the sign)"
                ElseIf CurrencyCode = ReportingCcy Then
                    Status = ST_EXCLUDED & "reporting currency carries no FX risk"
                Else
                    Status = ST_OK
                End If

                RowOutput(r, OC_ID) = PlainText(InputRows(r, FX_ID))
                RowOutput(r, OC_STATUS) = Status
                RowOutput(r, OC_CCY) = CurrencyCode
                RowOutput(r, OC_AMOUNT) = 0#
                RowOutput(r, OC_VALUE) = 0#
                If Status = ST_OK Then
                    RowOutput(r, OC_FX) = FxRates(Position)
                    RowOutput(r, OC_AMOUNT) = DirectionSign(InputRows(r, FX_DIRECTION)) _
                        * NumberValue(InputRows(r, FX_AMOUNT))
                    RowOutput(r, OC_VALUE) = RowOutput(r, OC_AMOUNT) * FxRates(Position)
                    CurrencyOutput(Position, 4) = CurrencyOutput(Position, 4) + RowOutput(r, OC_AMOUNT)
                    CurrencyOutput(Position, 6) = CurrencyOutput(Position, 6) + RowOutput(r, OC_VALUE)
                    Totals(TOT_USED) = Totals(TOT_USED) + 1
                ElseIf IsInvalidStatus(Status) Then
                    ErrorNumber = ErrorNumber + 1
                    RowOutput(r, OC_ERROR_NO) = ErrorNumber
                    Totals(TOT_REJECTED) = Totals(TOT_REJECTED) + 1
                Else
                    Totals(TOT_EXCLUDED) = Totals(TOT_EXCLUDED) + 1
                End If
            End If
        Next r

        '--- Net long, net short and gold per slot ------------------------------
        For s = 1 To SlotCount
            CurrencyOutput(s, 7) = 0#
            CurrencyOutput(s, 8) = 0#
            CurrencyOutput(s, 9) = 0#
            Select Case CurrencyOutput(s, 3)
                Case FX_CURRENCY
                    If CurrencyOutput(s, 6) > 0 Then CurrencyOutput(s, 7) = CurrencyOutput(s, 6)
                    If CurrencyOutput(s, 6) < 0 Then CurrencyOutput(s, 8) = -CurrencyOutput(s, 6)
                Case FX_GOLD_LABEL
                    CurrencyOutput(s, 9) = Abs(CurrencyOutput(s, 6))
            End Select
            Totals(TOT_LONG) = Totals(TOT_LONG) + CurrencyOutput(s, 7)
            Totals(TOT_SHORT) = Totals(TOT_SHORT) + CurrencyOutput(s, 8)
            Totals(TOT_GOLD) = Totals(TOT_GOLD) + CurrencyOutput(s, 9)
        Next s
        CurrencyOutput(SlotCount + 1, 2) = LBL_TOTAL_FX
        CurrencyOutput(SlotCount + 1, 7) = Totals(TOT_LONG)
        CurrencyOutput(SlotCount + 1, 8) = Totals(TOT_SHORT)
        CurrencyOutput(SlotCount + 1, 9) = Totals(TOT_GOLD)
        Totals(TOT_GENERAL) = NetOpenPositionCharge(Totals(TOT_LONG), Totals(TOT_SHORT), _
                                                    Totals(TOT_GOLD), Parameters(PI_FX_CHARGE), _
                                                    Totals(TOT_ONOP))

End Sub


Public Function DeMinimisWaived( _
    ByVal Regime As String, _
    ByVal OverallNetOpen As Double, _
    ByVal OwnFunds As Double, _
    ByVal Rate As Double) _
    As Boolean
'
'==============================================================================
'                               DeMinimisWaived
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether the FX requirement is waived by the de minimis of the CRR:
'   no requirement unless the overall net open position exceeds a share of
'   the total own funds.
'
' INPUTS
'   Regime          REGIME_BASEL or REGIME_CRR, from BuildRegime.
'   OverallNetOpen  Overall net open position, with gold, reporting currency.
'   OwnFunds        Total own funds, reporting currency.
'   Rate            Decimal share of the own funds; PI_FX_DE_MINIMIS (0.02).
'
' RETURNS
'   True when Regime is REGIME_CRR and OverallNetOpen does not exceed
'   Rate x OwnFunds; False otherwise, so that the Basel regime never waives.
'
' REFERENCE
'   Article 351 of Regulation (EU) No 575/2013, from Regulation (EU)
'   2024/1623: the requirement is calculated if the sum exceeds 2% of total
'   own funds.
'
' UPDATED
'   2026-11-20
'==============================================================================
'
        DeMinimisWaived = (Regime = REGIME_CRR) And (OverallNetOpen <= Rate * OwnFunds)

End Function


Public Function NetOpenPositionCharge( _
    ByVal SumNetLong As Double, _
    ByVal SumNetShort As Double, _
    ByVal GoldPosition As Double, _
    ByVal ChargeRate As Double, _
    ByRef OverallNetOpen As Double) _
    As Double
'
'==============================================================================
'                             NetOpenPositionCharge
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the shorthand method to the aggregated currency positions.
'
' INPUTS
'   SumNetLong    Sum of the net long currency positions, >= 0.
'   SumNetShort   Sum of the net short currency positions, as a positive
'                 amount, >= 0.
'   GoldPosition  Net gold position; its absolute value is used.
'   ChargeRate    Decimal, 0 to 1 (MAR40: 0.08).
'
' RETURNS
'   ChargeRate x OverallNetOpen; OverallNetOpen = max(SumNetLong,
'   SumNetShort) + |GoldPosition|. All in reporting currency.
'
' REFERENCE
'   MAR40.60-40.61, Table 9 (8% x (300 + 35) = 26.8).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        If SumNetLong > SumNetShort Then
            OverallNetOpen = SumNetLong + Abs(GoldPosition)
        Else
            OverallNetOpen = SumNetShort + Abs(GoldPosition)
        End If
        NetOpenPositionCharge = ChargeRate * OverallNetOpen

End Function
