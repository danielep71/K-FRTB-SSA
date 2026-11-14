Attribute VB_Name = "CORE_Commodity"
'==============================================================================
' MODULE: CORE_Commodity
'------------------------------------------------------------------------------
' PURPOSE
'   Commodity risk under the MAR40 simplified approach: validate each
'   position row, value it at spot in reporting currency, and charge 15% of
'   the net and 3% of the gross position per commodity. Returns the COCalc
'   table and the commodity block of Buckets.
'
' PUBLIC SURFACE
'   None outside this VBA project (Option Private Module).
'     CalculateCommodity.
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
'   BCBS MAR40.64-40.73 (simplified approach: MAR40.72-40.73).
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
    'Columns of the COCalc output, counted from column C.
        Private Const OC_ID          As Long = 1     'PositionID
        Private Const OC_STATUS      As Long = 2     'Status
        Private Const OC_COMMODITY   As Long = 3     'Commodity code
        Private Const OC_FX          As Long = 4     'FX rate
        Private Const OC_SIGNED      As Long = 5     'Signed value, reporting currency
        Private Const OC_GROSS       As Long = 6     'Gross value
        Private Const OC_COMMODITY_NO As Long = 7    'Commodity number
        Private Const OC_ERROR_NO    As Long = 8     'Running number of rejected rows


'------------------------------------------------------------------------------
'                                  CALCULATION
'------------------------------------------------------------------------------

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
'==============================================================================
'                              CalculateCommodity
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate the commodity charges for all positions and return every
'   intermediate.
'
' INPUTS
'   InputRows        COPositions cells (1 To n, 1 To CO_INPUT_WIDTH); a row whose
'               PositionID is blank is ignored. Quantities are
'               delta-equivalent; any sign of spot price is accepted.
'   Parameters  Vector (1 To PI_COUNT); PI_CO_DIRECTIONAL and PI_CO_GROSS
'               are used.
'   FxCodes     FX table codes, upper case; FxRates their rates (> 0).
'   GoldCode    Gold code, upper case; not a valid price currency.
'
' RETURNS
'   RowOutput        (1 To n, 1 To CO_CALC_WIDTH): COCalc columns C:J.
'   CommodityOutput  (1 To k + 1, 1 To BK_CO_WIDTH): Buckets C:L, one row
'                    per commodity in order of first valid row, then a
'                    total row (k >= 0).
'   Totals           (1 To TOT_COUNT): TOT_GENERAL (directional),
'                    TOT_OTHER (gross charge), TOT_USED, TOT_REJECTED,
'                    TOT_BUCKETS (commodities).
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when InputRows is not CO_INPUT_WIDTH wide.
'
' REFERENCE
'   MAR40.72 (15% of the net position per commodity), MAR40.73 (3% of the
'   gross position, valued at spot).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim RowCount    As Long        'Input rows
    Dim Ids()       As String      'Distinct PositionIDs
    Dim IdCount     As Long        'Distinct PositionIDs in use
    Dim IdUses()    As Long        'InputRows per PositionID
    Dim CommodityNames() As String    'Distinct commodities, in order
    Dim NameCount   As Long        'Commodities in use
    Dim Status      As String      'Row status
    Dim Commodity   As String      'Row commodity code
    Dim CurrencyCode As String     'Row currency code
    Dim Position    As Long        'Position in a lookup
    Dim ErrorNumber As Long        'Running number of rejected rows
    Dim k           As Long        'Commodity
    Dim r           As Long        'Input row
    Dim c           As Long        'Output column

        If UBound(InputRows, 2) < CO_INPUT_WIDTH Then
            Err.Raise ERR_ARGUMENT, "CORE_Commodity.CalculateCommodity", _
                "InputRows must be " & CO_INPUT_WIDTH & " columns wide."
        End If
        RowCount = UBound(InputRows, 1)
        ReDim RowOutput(1 To RowCount, 1 To CO_CALC_WIDTH)
        ReDim Totals(1 To TOT_COUNT)

        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, CO_ID)), True)
        Next r
        ReDim IdUses(1 To IdCount + 1)
        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, CO_ID)), False)
            If Position > 0 Then IdUses(Position) = IdUses(Position) + 1
        Next r

        '--- Row pass ----------------------------------------------------------
        For r = 1 To RowCount
            If Len(PlainText(InputRows(r, CO_ID))) > 0 Then
                Commodity = CodeText(InputRows(r, CO_COMMODITY))
                CurrencyCode = CodeText(InputRows(r, CO_CCY))
                If IdUses(KeyPosition(Ids, IdCount, CodeText(InputRows(r, CO_ID)), False)) > 1 Then
                    Status = ST_INVALID & "duplicate PositionID"
                ElseIf Len(Commodity) = 0 Then
                    Status = ST_INVALID & "Commodity missing"
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

                RowOutput(r, OC_ID) = PlainText(InputRows(r, CO_ID))
                RowOutput(r, OC_STATUS) = Status
                RowOutput(r, OC_COMMODITY) = Commodity
                RowOutput(r, OC_SIGNED) = 0#
                RowOutput(r, OC_GROSS) = 0#
                If Status = ST_OK Then
                    Position = FindCode(FxCodes, CurrencyCode)
                    RowOutput(r, OC_FX) = FxRates(Position)
                    RowOutput(r, OC_SIGNED) = DirectionSign(InputRows(r, CO_DIRECTION)) _
                        * NumberValue(InputRows(r, CO_QUANTITY)) * NumberValue(InputRows(r, CO_PRICE)) _
                        * FxRates(Position)
                    RowOutput(r, OC_GROSS) = Abs(RowOutput(r, OC_SIGNED))
                    If KeyPosition(CommodityNames, NameCount, Commodity, False) = 0 Then
                        RowOutput(r, OC_COMMODITY_NO) = KeyPosition(CommodityNames, NameCount, Commodity, True)
                    End If
                    Totals(TOT_USED) = Totals(TOT_USED) + 1
                Else
                    ErrorNumber = ErrorNumber + 1
                    RowOutput(r, OC_ERROR_NO) = ErrorNumber
                    Totals(TOT_REJECTED) = Totals(TOT_REJECTED) + 1
                End If
            End If
        Next r

        '--- Commodities (Buckets) ---------------------------------------------
        ReDim CommodityOutput(1 To NameCount + 1, 1 To BK_CO_WIDTH)
        For k = 1 To NameCount
            CommodityOutput(k, 1) = k
            CommodityOutput(k, 2) = CommodityNames(k)
            For c = 3 To BK_CO_WIDTH
                CommodityOutput(k, c) = 0#
            Next c
        Next k
        For r = 1 To RowCount
            If RowOutput(r, OC_STATUS) = ST_OK Then
                k = KeyPosition(CommodityNames, NameCount, CStr(RowOutput(r, OC_COMMODITY)), False)
                CommodityOutput(k, 3) = CommodityOutput(k, 3) + 1
                If RowOutput(r, OC_SIGNED) > 0 Then CommodityOutput(k, 4) = CommodityOutput(k, 4) + RowOutput(r, OC_SIGNED)
                If RowOutput(r, OC_SIGNED) < 0 Then CommodityOutput(k, 5) = CommodityOutput(k, 5) - RowOutput(r, OC_SIGNED)
                CommodityOutput(k, 7) = CommodityOutput(k, 7) + RowOutput(r, OC_GROSS)
            End If
        Next r
        For k = 1 To NameCount
            CommodityOutput(k, 6) = CommodityOutput(k, 4) - CommodityOutput(k, 5)
            CommodityOutput(k, 8) = Parameters(PI_CO_DIRECTIONAL) * Abs(CommodityOutput(k, 6))
            CommodityOutput(k, 9) = Parameters(PI_CO_GROSS) * CommodityOutput(k, 7)
            CommodityOutput(k, 10) = CommodityOutput(k, 8) + CommodityOutput(k, 9)
            Totals(TOT_GENERAL) = Totals(TOT_GENERAL) + CommodityOutput(k, 8)
            Totals(TOT_OTHER) = Totals(TOT_OTHER) + CommodityOutput(k, 9)
        Next k
        Totals(TOT_BUCKETS) = NameCount
        WriteTotalRow CommodityOutput, NameCount, LBL_TOTAL_CO

End Sub
