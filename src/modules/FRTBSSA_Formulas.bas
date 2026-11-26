Attribute VB_Name = "FRTBSSA_Formulas"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: FRTBSSA_Formulas
'------------------------------------------------------------------------------
' PURPOSE
'   Public worksheet functions for single steps of the MAR40 simplified
'   standardised approach, with the MAR40 reference parameters: the time band
'   of a debt position, the general market risk charge of one currency's
'   maturity ladder, the FX shorthand charge and the commodity simplified
'   charge
'
' WHY THIS EXISTS
'   They let a user check one figure by hand in a worksheet cell; RunSSA
'   does not use them
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Each function validates its arguments, reads the MAR40 reference ladder and
'   parameters, and calls the core procedure of its step:
'     - FRTBSSA_TimeBand: the time band of a debt position
'     - FRTBSSA_LadderCharge: the general market risk charge of one currency
'     - FRTBSSA_FXCharge: the FX and gold charge by the shorthand method
'     - FRTBSSA_CommodityCharge: the simplified commodity charge
'
' ERROR POLICY
'   Never raises: an argument outside its domain returns #VALUE! (wrong type
'   or shape) or #NUM! (number outside its domain)
'
' DEPENDENCIES
'   CORE_Config, CORE_Parameters, CORE_InterestRate, CORE_ForeignExchange and
'   CORE_Util
'
'   Reference: BCBS MAR40.26-40.28 (Tables 4 and 5), MAR40.60-40.61 and
'   MAR40.72-40.73; calculation contracts CALC-03, CALC-06 and CALC-05
'
' NOTES
'   Public module: the four functions are listed in PUBLIC_API.txt and are
'   visible to worksheets. The module owns no state
'
'   Compatibility: Excel VBA; no references beyond the defaults
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'
'------------------------------------------------------------------------------
'
'                             WORKSHEET FUNCTIONS
'
'------------------------------------------------------------------------------
'

Public Function FRTBSSA_TimeBand( _
    ByVal ResidualMaturity As Variant, _
    ByVal Coupon As Variant, _
    Optional ByVal Floating As Boolean = False) _
    As Variant
'
'------------------------------------------------------------------------------
'                               RETURN TIME BAND
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the maturity-method time band (1 to 15) of a debt position
'
' WHY THIS EXISTS
'   A user can check the band of one position in a cell without running
'   RunSSA
'
' INPUTS
'   ResidualMaturity
'     Years to final maturity (fixed) or to the next repricing date
'     (floating); >= 0
'
'   Coupon
'     Annual coupon, decimal (0.03 = 3%); ignored when Floating is True
'
'   Floating
'     True for a floating-rate position (slotted with the coupon >= 3%
'     bands)
'
' RETURNS
'   Band number; #VALUE! for a non-numeric argument, #NUM! for a negative
'   maturity
'
' BEHAVIOR
'   Validates the arguments, loads the MAR40 reference ladder and parameters,
'   and slots the position:
'     - a floating-rate position uses the coupon >= 3% bands
'     - a fixed-rate position uses them when its coupon is at or above
'       PI_COUPON_THRESHOLD
'
' ERROR POLICY
'   Never raises: a non-numeric argument returns #VALUE! and a negative
'   maturity returns #NUM!
'
' DEPENDENCIES
'   IsNumberValue, ReferenceLadder, ReferenceParameters and TimeBand
'
'   Reference: MAR40.25-40.26, Table 4
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
    Dim Ladder()        As Double   'MAR40 ladder
    Dim Parameters()    As Double   'MAR40 parameters

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a non-numeric maturity, or a non-numeric coupon on a fixed-rate position
        If Not IsNumberValue(ResidualMaturity) Or (Not Floating And Not IsNumberValue(Coupon)) Then
            FRTBSSA_TimeBand = CVErr(xlErrValue)
            Exit Function
        End If
    'Reject a negative residual maturity
        If CDbl(ResidualMaturity) < 0 Then
            FRTBSSA_TimeBand = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' LOAD REFERENCE PARAMETERS
'------------------------------------------------------------------------------
    'Load the MAR40 ladder
        ReferenceLadder Ladder
    'Load the MAR40 parameters
        ReferenceParameters Parameters

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Slot a floating-rate position by the coupon >= 3% bands, and a fixed-rate one by its coupon
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
'------------------------------------------------------------------------------
'                             RETURN LADDER CHARGE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the general market risk charge of one currency from its 15
'   weighted long and 15 weighted short positions
'
' WHY THIS EXISTS
'   A user can check the charge of one currency's ladder in a cell without
'   running RunSSA
'
' INPUTS
'   WeightedLong
'     15 cells (or array) of weighted long positions, band 1 to 15, each
'     >= 0; blank = 0
'
'   WeightedShort
'     15 cells of weighted short positions as positive amounts, each >= 0;
'     blank = 0
'
' RETURNS
'   Vertical + within-zone + between-zone disallowances + net position, in
'   the unit of the arguments; #VALUE! for another shape or text, #NUM! for
'   a negative position
'
' BEHAVIOR
'   Reads the longs and the shorts into 15-band arrays, loads the MAR40
'   reference ladder and parameters, applies the disallowances of the
'   maturity method and returns the general market risk charge
'
' ERROR POLICY
'   Never raises: a problem found while reading the positions returns the
'   worksheet error of its code
'
' DEPENDENCIES
'   BandValues, ReferenceLadder, ReferenceParameters and LadderCharges
'
'   Reference: MAR40.27-40.28, Table 5
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
    Dim LongValues()    As Double   'Weighted longs
    Dim ShortValues()   As Double   'Weighted shorts
    Dim Ladder()        As Double   'MAR40 ladder
    Dim Parameters()    As Double   'MAR40 parameters
    Dim Components()    As Double   'LadderCharges result
    Dim Problem         As Long     'Worksheet error code, 0 for none

'------------------------------------------------------------------------------
' READ POSITIONS
'------------------------------------------------------------------------------
    'Read the weighted longs, then the weighted shorts unless the longs failed
        Problem = BandValues(WeightedLong, LongValues)
        If Problem = 0 Then Problem = BandValues(WeightedShort, ShortValues)
    'Return the worksheet error of the first problem
        If Problem <> 0 Then
            FRTBSSA_LadderCharge = CVErr(Problem)
            Exit Function
        End If

'------------------------------------------------------------------------------
' LOAD REFERENCE PARAMETERS
'------------------------------------------------------------------------------
    'Load the MAR40 ladder
        ReferenceLadder Ladder
    'Load the MAR40 parameters
        ReferenceParameters Parameters

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Apply the vertical and horizontal disallowances of the maturity method
        LadderCharges LongValues, ShortValues, Ladder, Parameters, Components
    'Return the general market risk charge
        FRTBSSA_LadderCharge = Components(ZC_GENERAL)

End Function

Public Function FRTBSSA_FXCharge( _
    ByVal SumNetLong As Variant, _
    ByVal SumNetShort As Variant, _
    Optional ByVal GoldPosition As Variant = 0) _
    As Variant
'
'------------------------------------------------------------------------------
'                               RETURN FX CHARGE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the FX and gold requirement under the shorthand method
'
' WHY THIS EXISTS
'   A user can check the FX requirement from the sums of net positions in a
'   cell without running RunSSA
'
' INPUTS
'   SumNetLong
'     Sum of net long currency positions, >= 0
'
'   SumNetShort
'     Sum of net short currency positions, positive, >= 0
'
'   GoldPosition
'     Net gold position, any sign
'
' RETURNS
'   8% x (max(SumNetLong, SumNetShort) + |GoldPosition|); #VALUE! for text,
'   #NUM! for a negative sum
'
' BEHAVIOR
'   Validates the three arguments, loads the MAR40 reference parameters and
'   returns the charge on the overall net open position
'
' ERROR POLICY
'   Never raises: text returns #VALUE! and a negative sum returns #NUM!
'
' DEPENDENCIES
'   IsNumberValue, ReferenceParameters and NetOpenPositionCharge
'
'   Reference: MAR40.60-40.61, Table 9
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
    Dim Parameters()    As Double   'MAR40 parameters
    Dim OverallNet      As Double   'Overall net open position

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an argument that is not a number
        If Not IsNumberValue(SumNetLong) Or Not IsNumberValue(SumNetShort) _
            Or Not IsNumberValue(GoldPosition) Then
            FRTBSSA_FXCharge = CVErr(xlErrValue)
            Exit Function
        End If
    'Reject a negative sum of net positions
        If CDbl(SumNetLong) < 0 Or CDbl(SumNetShort) < 0 Then
            FRTBSSA_FXCharge = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' LOAD REFERENCE PARAMETERS
'------------------------------------------------------------------------------
    'Load the MAR40 parameters
        ReferenceParameters Parameters

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the charge on the overall net open position
        FRTBSSA_FXCharge = NetOpenPositionCharge(CDbl(SumNetLong), CDbl(SumNetShort), _
            CDbl(GoldPosition), Parameters(PI_FX_CHARGE), OverallNet)

End Function

Public Function FRTBSSA_CommodityCharge( _
    ByVal NetPosition As Variant, _
    ByVal GrossPosition As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                           RETURN COMMODITY CHARGE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the simplified-approach charge of one commodity
'
' WHY THIS EXISTS
'   A user can check the charge of one commodity from its net and gross
'   positions in a cell without running RunSSA
'
' INPUTS
'   NetPosition
'     Net position at spot, any sign
'
'   GrossPosition
'     Sum of the absolute positions at spot, >= |NetPosition|
'
' RETURNS
'   15% x |NetPosition| + 3% x GrossPosition; #VALUE! for text, #NUM! when
'   GrossPosition < |NetPosition|
'
' BEHAVIOR
'   Validates the two arguments, loads the MAR40 reference parameters and
'   adds the directional charge on the net position to the charge on the
'   gross position
'
' ERROR POLICY
'   Never raises: text returns #VALUE! and a gross position below the
'   absolute net position returns #NUM!
'
' DEPENDENCIES
'   IsNumberValue and ReferenceParameters
'
'   Reference: MAR40.72-40.73
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
    Dim Parameters()    As Double   'MAR40 parameters

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an argument that is not a number
        If Not IsNumberValue(NetPosition) Or Not IsNumberValue(GrossPosition) Then
            FRTBSSA_CommodityCharge = CVErr(xlErrValue)
            Exit Function
        End If
    'Reject a gross position below the absolute net position
        If CDbl(GrossPosition) < Abs(CDbl(NetPosition)) Then
            FRTBSSA_CommodityCharge = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' LOAD REFERENCE PARAMETERS
'------------------------------------------------------------------------------
    'Load the MAR40 parameters
        ReferenceParameters Parameters

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the directional charge plus the charge on the gross position
        FRTBSSA_CommodityCharge = Parameters(PI_CO_DIRECTIONAL) * Abs(CDbl(NetPosition)) _
            + Parameters(PI_CO_GROSS) * CDbl(GrossPosition)

End Function

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function BandValues( _
    ByVal Source As Variant, _
    ByRef Values() As Double) _
    As Long
'
'------------------------------------------------------------------------------
'                               READ BAND VALUES
'------------------------------------------------------------------------------
' PURPOSE
'   Reads 15 band positions from a range or an array
'
' WHY THIS EXISTS
'   The ladder function takes its two arguments as cells or arrays and
'   needs one place that checks their shape and values
'
' INPUTS
'   Source
'     A range of 15 cells (one row or one column) or an array of 15
'
'   Values
'     Receives the 15 values, band 1 to BAND_COUNT
'
' RETURNS
'   0 and Values(1 To BAND_COUNT) on success; xlErrValue for another shape or
'   a non-numeric cell; xlErrNum for a negative value
'
' BEHAVIOR
'   Takes the values of a range, or the array as given, and reads them in
'   order:
'     - an empty cell counts as 0
'     - a non-numeric item, or an item after the 15th, stops the read
'     - a negative item stops the read
'     - a count other than BAND_COUNT is a shape error
'
' ERROR POLICY
'   Never raises: a problem is returned as the worksheet error code
'
' DEPENDENCIES
'   IsNumberValue; BAND_COUNT, xlErrValue and xlErrNum
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
    Dim Items   As Variant  'Values of the source
    Dim Item    As Variant  'One value
    Dim Count   As Long     'Values read

'------------------------------------------------------------------------------
' READ SOURCE
'------------------------------------------------------------------------------
    'Take the values of a range, or the array as given
        If TypeName(Source) = "Range" Then
            Items = Source.Value
        Else
            Items = Source
        End If
    'Reject a source that is not an array
        If Not IsArray(Items) Then
            BandValues = xlErrValue
            Exit Function
        End If
    'Size the result to the 15 bands
        ReDim Values(1 To BAND_COUNT)

'------------------------------------------------------------------------------
' READ VALUES
'------------------------------------------------------------------------------
    'Read each item and reject a surplus item, text or a negative value
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

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Report a shape error unless exactly 15 items were read
        If Count <> BAND_COUNT Then
            BandValues = xlErrValue
        Else
            BandValues = 0
        End If

End Function
