Attribute VB_Name = "FRTBSSA_Formulas"
'==============================================================================
' MODULE: FRTBSSA_Formulas
'------------------------------------------------------------------------------
' PURPOSE
'   Public worksheet functions for single steps of the MAR40 simplified
'   standardised approach, with the MAR40 reference parameters: the time
'   band of a debt position, the general market risk charge of one
'   currency's maturity ladder, the FX shorthand charge and the commodity
'   simplified charge. They let a user check one figure by hand; RunSSA does
'   not use them.
'
' PUBLIC SURFACE
'   FRTBSSA_TimeBand, FRTBSSA_LadderCharge, FRTBSSA_FXCharge, FRTBSSA_CommodityCharge
'   (listed in PUBLIC_API.txt).
'
' DEPENDENCIES
'   CORE_Config, CORE_Parameters, CORE_InterestRate, CORE_ForeignExchange,
'   CORE_Util.
'
' STATE OWNERSHIP
'   None.
'
' ERROR POLICY
'   Never raise: an argument outside its domain returns #VALUE! (wrong type
'   or shape) or #NUM! (number outside its domain).
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' REFERENCE
'   BCBS MAR40.26-40.28 (Tables 4 and 5), MAR40.60-40.61, MAR40.72-40.73.
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
    'Require explicit declarations. Public module: the functions are
    'visible to worksheets.
    Option Explicit


'------------------------------------------------------------------------------
'                              WORKSHEET FUNCTIONS
'------------------------------------------------------------------------------

Public Function FRTBSSA_TimeBand( _
    ByVal ResidualMaturity As Variant, _
    ByVal Coupon As Variant, _
    Optional ByVal Floating As Boolean = False) _
    As Variant
'
'==============================================================================
'                                 FRTBSSA_TimeBand
'------------------------------------------------------------------------------
' PURPOSE
'   Return the maturity-method time band (1 to 15) of a debt position.
'
' INPUTS
'   ResidualMaturity  Years to final maturity (fixed) or to the next
'                     repricing date (floating); >= 0.
'   Coupon            Annual coupon, decimal (0.03 = 3%); ignored when
'                     Floating is True.
'   Floating          True for a floating-rate position (slotted with the
'                     coupon >= 3% bands).
'
' RETURNS
'   Band number; #VALUE! for a non-numeric argument, #NUM! for a negative
'   maturity.
'
' REFERENCE
'   MAR40.25-40.26, Table 4.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Ladder()     As Double   'MAR40 ladder
    Dim Parameters() As Double   'MAR40 parameters

        If Not IsNumberValue(ResidualMaturity) Or (Not Floating And Not IsNumberValue(Coupon)) Then
            FRTBSSA_TimeBand = CVErr(xlErrValue)
            Exit Function
        End If
        If CDbl(ResidualMaturity) < 0 Then
            FRTBSSA_TimeBand = CVErr(xlErrNum)
            Exit Function
        End If
        ReferenceLadder Ladder
        ReferenceParameters Parameters
        If Floating Then
            FRTBSSA_TimeBand = TimeBand(CDbl(ResidualMaturity), True, Ladder)
        Else
            FRTBSSA_TimeBand = TimeBand(CDbl(ResidualMaturity), _
                                    CDbl(Coupon) >= Parameters(PI_COUPON_THRESHOLD), Ladder)
        End If

End Function


Public Function FRTBSSA_LadderCharge( _
    ByVal WeightedLong As Variant, _
    ByVal WeightedShort As Variant) _
    As Variant
'
'==============================================================================
'                               FRTBSSA_LadderCharge
'------------------------------------------------------------------------------
' PURPOSE
'   Return the general market risk charge of one currency from its 15
'   weighted long and 15 weighted short positions.
'
' INPUTS
'   WeightedLong    15 cells (or array) of weighted long positions, band 1
'                   to 15, each >= 0; blank = 0.
'   WeightedShort   15 cells of weighted short positions as positive
'                   amounts, each >= 0; blank = 0.
'
' RETURNS
'   Vertical + within-zone + between-zone disallowances + net position, in
'   the unit of the arguments; #VALUE! for another shape or text, #NUM! for
'   a negative position.
'
' REFERENCE
'   MAR40.27-40.28, Table 5.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim LongValues()  As Double  'Weighted longs
    Dim ShortValues() As Double  'Weighted shorts
    Dim Ladder()      As Double  'MAR40 ladder
    Dim Parameters()  As Double  'MAR40 parameters
    Dim Components()  As Double  'LadderCharges result
    Dim Problem       As Long    'Worksheet error code, 0 for none

        Problem = BandValues(WeightedLong, LongValues)
        If Problem = 0 Then Problem = BandValues(WeightedShort, ShortValues)
        If Problem <> 0 Then
            FRTBSSA_LadderCharge = CVErr(Problem)
            Exit Function
        End If
        ReferenceLadder Ladder
        ReferenceParameters Parameters
        LadderCharges LongValues, ShortValues, Ladder, Parameters, Components
        FRTBSSA_LadderCharge = Components(ZC_GENERAL)

End Function


Public Function FRTBSSA_FXCharge( _
    ByVal SumNetLong As Variant, _
    ByVal SumNetShort As Variant, _
    Optional ByVal GoldPosition As Variant = 0) _
    As Variant
'
'==============================================================================
'                                 FRTBSSA_FXCharge
'------------------------------------------------------------------------------
' PURPOSE
'   Return the FX and gold requirement under the shorthand method.
'
' INPUTS
'   SumNetLong     Sum of net long currency positions, >= 0.
'   SumNetShort    Sum of net short currency positions, positive, >= 0.
'   GoldPosition   Net gold position, any sign.
'
' RETURNS
'   8% x (max(SumNetLong, SumNetShort) + |GoldPosition|); #VALUE! for text,
'   #NUM! for a negative sum.
'
' REFERENCE
'   MAR40.60-40.61, Table 9.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Parameters() As Double   'MAR40 parameters
    Dim OverallNet   As Double   'Overall net open position

        If Not IsNumberValue(SumNetLong) Or Not IsNumberValue(SumNetShort) _
           Or Not IsNumberValue(GoldPosition) Then
            FRTBSSA_FXCharge = CVErr(xlErrValue)
            Exit Function
        End If
        If CDbl(SumNetLong) < 0 Or CDbl(SumNetShort) < 0 Then
            FRTBSSA_FXCharge = CVErr(xlErrNum)
            Exit Function
        End If
        ReferenceParameters Parameters
        FRTBSSA_FXCharge = NetOpenPositionCharge(CDbl(SumNetLong), CDbl(SumNetShort), _
                                             CDbl(GoldPosition), Parameters(PI_FX_CHARGE), OverallNet)

End Function


Public Function FRTBSSA_CommodityCharge( _
    ByVal NetPosition As Variant, _
    ByVal GrossPosition As Variant) _
    As Variant
'
'==============================================================================
'                              FRTBSSA_CommodityCharge
'------------------------------------------------------------------------------
' PURPOSE
'   Return the simplified-approach charge of one commodity.
'
' INPUTS
'   NetPosition     Net position at spot, any sign.
'   GrossPosition   Sum of the absolute positions at spot, >= |NetPosition|.
'
' RETURNS
'   15% x |NetPosition| + 3% x GrossPosition; #VALUE! for text, #NUM! when
'   GrossPosition < |NetPosition|.
'
' REFERENCE
'   MAR40.72-40.73.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Parameters() As Double   'MAR40 parameters

        If Not IsNumberValue(NetPosition) Or Not IsNumberValue(GrossPosition) Then
            FRTBSSA_CommodityCharge = CVErr(xlErrValue)
            Exit Function
        End If
        If CDbl(GrossPosition) < Abs(CDbl(NetPosition)) Then
            FRTBSSA_CommodityCharge = CVErr(xlErrNum)
            Exit Function
        End If
        ReferenceParameters Parameters
        FRTBSSA_CommodityCharge = Parameters(PI_CO_DIRECTIONAL) * Abs(CDbl(NetPosition)) _
                              + Parameters(PI_CO_GROSS) * CDbl(GrossPosition)

End Function


'------------------------------------------------------------------------------
'                                    HELPERS
'------------------------------------------------------------------------------

Private Function BandValues( _
    ByVal Source As Variant, _
    ByRef Values() As Double) _
    As Long
'
'==============================================================================
'                                  BandValues
'------------------------------------------------------------------------------
' PURPOSE
'   Read 15 band positions from a range or an array.
'
' INPUTS
'   Source   A range of 15 cells (one row or one column) or an array of 15.
'
' RETURNS
'   0 and Values(1 To BAND_COUNT) on success; xlErrValue for another shape
'   or a non-numeric cell; xlErrNum for a negative value.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Items       As Variant   'Values of the source
    Dim Item        As Variant   'One value
    Dim Count       As Long      'Values read

        If TypeName(Source) = "Range" Then Items = Source.Value Else Items = Source
        If Not IsArray(Items) Then
            BandValues = xlErrValue
            Exit Function
        End If
        ReDim Values(1 To BAND_COUNT)
        For Each Item In Items
            Count = Count + 1
            If Count > BAND_COUNT Then
                BandValues = xlErrValue
                Exit Function
            End If
            If IsEmpty(Item) Then
                Values(Count) = 0#
            ElseIf Not IsNumberValue(Item) Then
                BandValues = xlErrValue
                Exit Function
            ElseIf CDbl(Item) < 0 Then
                BandValues = xlErrNum
                Exit Function
            Else
                Values(Count) = CDbl(Item)
            End If
        Next Item
        If Count <> BAND_COUNT Then BandValues = xlErrValue Else BandValues = 0

End Function
