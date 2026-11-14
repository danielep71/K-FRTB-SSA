Attribute VB_Name = "CORE_InterestRate"
'==============================================================================
' MODULE: CORE_InterestRate
'------------------------------------------------------------------------------
' PURPOSE
'   Interest rate risk under the MAR40 simplified standardised approach:
'   validate each position row, slot it in the maturity ladder of its
'   currency, weight it, net specific risk per issue, and apply the
'   vertical and horizontal disallowances of the maturity method per
'   currency. Returns the IRCalc, IR_Ladder and IR_Zones tables as arrays.
'
' PUBLIC SURFACE
'   None outside this VBA project (Option Private Module).
'     CalculateInterestRate, TimeBand, LadderCharges.
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
'   BCBS MAR40.4-40.8 (specific risk, Table 1), MAR40.24-40.28 (maturity
'   method, Tables 4 and 5).
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
    'Columns of the IRCalc output, counted from column C.
        Private Const OC_ID          As Long = 1     'PositionID
        Private Const OC_STATUS      As Long = 2     'Status
        Private Const OC_CCY         As Long = 3     'Currency code
        Private Const OC_FX          As Long = 4     'FX rate
        Private Const OC_SIGNED      As Long = 5     'Signed value, reporting currency
        Private Const OC_SLOT_DATE   As Long = 6     'Slotting date
        Private Const OC_MATURITY    As Long = 7     'Residual maturity, years
        Private Const OC_COUPON_TBL  As Long = 8     'High or Low coupon table
        Private Const OC_BAND        As Long = 9     'Time band
        Private Const OC_ZONE        As Long = 10    'Zone
        Private Const OC_WEIGHT      As Long = 11    'Risk weight
        Private Const OC_WEIGHTED    As Long = 12    'Weighted position
        Private Const OC_CATEGORY    As Long = 13    'SR category
        Private Const OC_SR_BAND     As Long = 14    'SR maturity band
        Private Const OC_SR_KEY      As Long = 15    'SR key
        Private Const OC_SR_RATE     As Long = 16    'SR rate
        Private Const OC_ISSUE       As Long = 17    'Issue key
        Private Const OC_ISSUE_NET   As Long = 18    'Issue net value
        Private Const OC_ISSUE_FIRST As Long = 19    'First row of issue
        Private Const OC_SR_CHARGE   As Long = 20    'Specific risk charge
        Private Const OC_MIXED       As Long = 21    'MIXED flag
        Private Const OC_ERROR_NO    As Long = 22    'Running number of rejected rows
    'Specific-risk maturity bands (MAR40.6 Table 1).
        Private Const SRB_SHORT      As String = "<=6m"      'Six months or less
        Private Const SRB_MEDIUM     As String = "6m-24m"    'Over 6 to 24 months
        Private Const SRB_LONG       As String = ">24m"      'Over 24 months
        Private Const SRB_ANY        As String = "Any"       'Any maturity


'------------------------------------------------------------------------------
'                                  CALCULATION
'------------------------------------------------------------------------------

Public Sub CalculateInterestRate( _
    ByRef InputRows() As Variant, _
    ByVal AsOfDate As Double, _
    ByRef Parameters() As Double, _
    ByRef FxCodes() As String, _
    ByRef FxRates() As Double, _
    ByVal GoldCode As String, _
    ByRef SrKeys() As String, _
    ByRef SrCharges() As Double, _
    ByRef Ladder() As Double, _
    ByRef LadderLabels() As String, _
    ByRef RowOutput() As Variant, _
    ByRef LadderOutput() As Variant, _
    ByRef ZoneOutput() As Variant, _
    ByRef Totals() As Double)
'
'==============================================================================
'                             CalculateInterestRate
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate interest rate specific and general market risk for all
'   positions and return every intermediate.
'
' INPUTS
'   InputRows           IRPositions cells (1 To n, 1 To IR_INPUT_WIDTH); a row
'                  whose PositionID is blank is ignored.
'   AsOfDate       Reporting date, date serial.
'   Parameters     Vector (1 To PI_COUNT) from CORE_Parameters.
'   FxCodes        FX table codes (1 To m), upper case, "" for an empty slot.
'   FxRates        Units of reporting currency per unit, > 0.
'   GoldCode       Gold code, upper case; gold is not an IR currency.
'   SrKeys         Specific-risk keys, upper case; SrCharges their decimals.
'   Ladder         (1 To BAND_COUNT, 1 To LD_WIDTH) from BuildLadder.
'   LadderLabels   (1 To BAND_COUNT, 1 To 2): band labels, coupon >= and <
'                  threshold; "" when the band does not exist.
'
' RETURNS
'   RowOutput      (1 To n, 1 To IR_CALC_WIDTH): IRCalc columns C:X.
'   LadderOutput   (1 To m * BAND_COUNT, 1 To LADDER_WIDTH): IR_Ladder C:O,
'                  one block of bands per FX table slot.
'   ZoneOutput     (1 To m + 1, 1 To ZONES_WIDTH): IR_Zones C:AF, one row
'                  per FX table slot and a total row.
'   Totals         (1 To TOT_COUNT): TOT_SPECIFIC, TOT_GENERAL (reporting
'                  currency), TOT_USED, TOT_REJECTED, TOT_MIXED.
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when InputRows is not IR_INPUT_WIDTH wide.
'
' REFERENCE
'   MAR40.6 (specific risk on the net position per issue), MAR40.24 (one
'   ladder per currency), MAR40.25 (fixed by residual maturity, floating by
'   next repricing), MAR40.26-40.28.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim RowCount      As Long        'Input rows
    Dim SlotCount     As Long        'FX table slots
    Dim Ids()         As String      'Normalised PositionIDs, for duplicates
    Dim IdCount       As Long        'Distinct PositionIDs
    Dim IdUses()      As Long        'InputRows per distinct PositionID
    Dim RowSlot()     As Long        'FX slot of each OK row, 0 otherwise
    Dim Issues()      As String      'Distinct issue keys
    Dim IssueCount    As Long        'Distinct issue keys in use
    Dim IssueNet()    As Double      'Net signed value per issue
    Dim IssueFirst()  As Long        'First row of each issue
    Dim IssueMixed()  As Boolean     'Issue rows with different SR key or currency
    Dim WeightedLong()  As Double    'Per slot and band, weighted long
    Dim WeightedShort() As Double    'Per slot and band, weighted short
    Dim PlainLong()   As Double      'Per slot and band, long before weighting
    Dim PlainShort()  As Double      'Per slot and band, short before weighting
    Dim SlotUsed()    As Long        'OK rows per slot
    Dim SlotSpecific() As Double     'Specific risk charge per slot
    Dim BandLong()    As Double      'One slot's weighted longs (1 To BAND_COUNT)
    Dim BandShort()   As Double      'One slot's weighted shorts
    Dim Components()  As Double      'LadderCharges result
    Dim Status        As String      'Row status
    Dim CurrencyCode  As String      'Row currency code
    Dim Category      As String      'Row SR category
    Dim SrBand        As String      'Row SR maturity band
    Dim SrKey         As String      'Row SR key
    Dim Grade         As String      'Row rating grade, trimmed
    Dim IsFloating    As Boolean     'Rate type Floating
    Dim SlotDate      As Double      'Slotting date serial
    Dim Residual      As Double      'Residual maturity, years
    Dim Signed        As Double      'Signed value, reporting currency
    Dim Band          As Long        'Time band
    Dim Position      As Long        'Position in a lookup
    Dim ErrorNumber   As Long        'Running number of rejected rows
    Dim r             As Long        'Input row
    Dim s             As Long        'FX slot
    Dim b             As Long        'Time band
    Dim c             As Long        'Output column

        If UBound(InputRows, 2) < IR_INPUT_WIDTH Then
            Err.Raise ERR_ARGUMENT, "CORE_InterestRate.CalculateInterestRate", _
                "InputRows must be " & IR_INPUT_WIDTH & " columns wide."
        End If
        RowCount = UBound(InputRows, 1)
        SlotCount = UBound(FxCodes)
        ReDim RowOutput(1 To RowCount, 1 To IR_CALC_WIDTH)
        ReDim RowSlot(1 To RowCount)
        ReDim Totals(1 To TOT_COUNT)

        'Count each PositionID, ignoring case, to find duplicates.
        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, IR_ID)), True)
        Next r
        ReDim IdUses(1 To IdCount + 1)
        For r = 1 To RowCount
            Position = KeyPosition(Ids, IdCount, CodeText(InputRows(r, IR_ID)), False)
            If Position > 0 Then IdUses(Position) = IdUses(Position) + 1
        Next r

        '--- Row pass: status, slotting and weighting --------------------------
        For r = 1 To RowCount
            If Len(PlainText(InputRows(r, IR_ID))) > 0 Then
                CurrencyCode = CodeText(InputRows(r, IR_CCY))
                Category = CodeText(InputRows(r, IR_CATEGORY))
                Grade = PlainText(InputRows(r, IR_GRADE))
                IsFloating = (CodeText(InputRows(r, IR_RATE_TYPE)) = "FLOATING")
                SrBand = SpecificRiskBand(InputRows(r, IR_MATURITY), AsOfDate, Parameters(PI_DAYS_PER_YEAR))
                SrKey = SpecificRiskKey(Category, Grade, SrBand, SrKeys)
                Status = RowStatus(InputRows, r, CurrencyCode, Category, SrKey, IsFloating, AsOfDate, _
                                   FxCodes, GoldCode, _
                                   IdUses(KeyPosition(Ids, IdCount, CodeText(InputRows(r, IR_ID)), False)))

                RowOutput(r, OC_ID) = PlainText(InputRows(r, IR_ID))
                RowOutput(r, OC_STATUS) = Status
                RowOutput(r, OC_CCY) = CurrencyCode
                Position = FindCode(FxCodes, CurrencyCode)
                If Position >= 1 Then RowOutput(r, OC_FX) = FxRates(Position)
                RowOutput(r, OC_SIGNED) = 0#
                RowOutput(r, OC_WEIGHTED) = 0#
                RowOutput(r, OC_CATEGORY) = Category
                RowOutput(r, OC_SR_BAND) = SrBand
                RowOutput(r, OC_SR_KEY) = SrKey
                RowOutput(r, OC_SR_RATE) = 0#
                RowOutput(r, OC_ISSUE_FIRST) = False
                RowOutput(r, OC_SR_CHARGE) = 0#

                If Status = ST_OK Then
                    Signed = DirectionSign(InputRows(r, IR_DIRECTION)) * NumberValue(InputRows(r, IR_VALUE)) _
                             * FxRates(Position)
                    If IsFloating Then
                        SlotDate = NumberValue(InputRows(r, IR_REPRICING))
                    Else
                        SlotDate = NumberValue(InputRows(r, IR_MATURITY))
                    End If
                    Residual = (SlotDate - AsOfDate) / Parameters(PI_DAYS_PER_YEAR)
                    Band = TimeBand(Residual, IsHighCoupon(InputRows, r, IsFloating, Parameters), Ladder)
                    RowSlot(r) = Position
                    RowOutput(r, OC_SIGNED) = Signed
                    RowOutput(r, OC_SLOT_DATE) = CDate(SlotDate)
                    RowOutput(r, OC_MATURITY) = Residual
                    If IsHighCoupon(InputRows, r, IsFloating, Parameters) Then
                        RowOutput(r, OC_COUPON_TBL) = "High"
                    Else
                        RowOutput(r, OC_COUPON_TBL) = "Low"
                    End If
                    RowOutput(r, OC_BAND) = Band
                    RowOutput(r, OC_ZONE) = Ladder(Band, LD_ZONE)
                    RowOutput(r, OC_WEIGHT) = Ladder(Band, LD_WEIGHT)
                    RowOutput(r, OC_WEIGHTED) = Signed * Ladder(Band, LD_WEIGHT)
                    If SrKey <> SR_NONE Then
                        RowOutput(r, OC_SR_RATE) = SrCharges(FindCode(SrKeys, SrKey))
                    End If
                    If Category <> SR_NONE Then
                        If Len(PlainText(InputRows(r, IR_ISSUE))) = 0 Then
                            RowOutput(r, OC_ISSUE) = PlainText(InputRows(r, IR_ID))
                        Else
                            RowOutput(r, OC_ISSUE) = CodeText(InputRows(r, IR_ISSUE))
                        End If
                    End If
                    Totals(TOT_USED) = Totals(TOT_USED) + 1
                ElseIf IsInvalidStatus(Status) Then
                    ErrorNumber = ErrorNumber + 1
                    RowOutput(r, OC_ERROR_NO) = ErrorNumber
                    Totals(TOT_REJECTED) = Totals(TOT_REJECTED) + 1
                End If
            End If
        Next r

        '--- Issue pass: net per issue, charge on the first row ----------------
        'There are at most RowCount distinct issues.
        ReDim IssueNet(1 To RowCount)
        ReDim IssueFirst(1 To RowCount)
        ReDim IssueMixed(1 To RowCount)
        For r = 1 To RowCount
            If Len(CStr(RowOutput(r, OC_ISSUE))) > 0 Then
                Position = KeyPosition(Issues, IssueCount, CStr(RowOutput(r, OC_ISSUE)), True)
                If IssueFirst(Position) = 0 Then
                    IssueFirst(Position) = r
                Else
                    If UCase$(CStr(RowOutput(r, OC_SR_KEY))) <> _
                       UCase$(CStr(RowOutput(IssueFirst(Position), OC_SR_KEY))) _
                       Or RowOutput(r, OC_CCY) <> RowOutput(IssueFirst(Position), OC_CCY) Then
                        IssueMixed(Position) = True
                    End If
                End If
                IssueNet(Position) = IssueNet(Position) + RowOutput(r, OC_SIGNED)
            End If
        Next r
        ReDim SlotSpecific(1 To SlotCount)
        For r = 1 To RowCount
            If Len(CStr(RowOutput(r, OC_ISSUE))) > 0 Then
                Position = KeyPosition(Issues, IssueCount, CStr(RowOutput(r, OC_ISSUE)), False)
                RowOutput(r, OC_ISSUE_NET) = IssueNet(Position)
                If IssueFirst(Position) = r Then
                    RowOutput(r, OC_ISSUE_FIRST) = True
                    RowOutput(r, OC_SR_CHARGE) = Abs(IssueNet(Position)) * RowOutput(r, OC_SR_RATE)
                    SlotSpecific(RowSlot(r)) = SlotSpecific(RowSlot(r)) + RowOutput(r, OC_SR_CHARGE)
                    Totals(TOT_SPECIFIC) = Totals(TOT_SPECIFIC) + RowOutput(r, OC_SR_CHARGE)
                End If
                If IssueMixed(Position) Then
                    RowOutput(r, OC_MIXED) = LBL_MIXED
                    Totals(TOT_MIXED) = Totals(TOT_MIXED) + 1
                Else
                    RowOutput(r, OC_MIXED) = vbNullString
                End If
            End If
        Next r

        '--- Ladder: accumulate per FX slot and band ---------------------------
        ReDim WeightedLong(1 To SlotCount, 1 To BAND_COUNT)
        ReDim WeightedShort(1 To SlotCount, 1 To BAND_COUNT)
        ReDim PlainLong(1 To SlotCount, 1 To BAND_COUNT)
        ReDim PlainShort(1 To SlotCount, 1 To BAND_COUNT)
        ReDim SlotUsed(1 To SlotCount)
        For r = 1 To RowCount
            s = RowSlot(r)
            If s > 0 Then
                b = RowOutput(r, OC_BAND)
                SlotUsed(s) = SlotUsed(s) + 1
                If RowOutput(r, OC_SIGNED) > 0 Then PlainLong(s, b) = PlainLong(s, b) + RowOutput(r, OC_SIGNED)
                If RowOutput(r, OC_SIGNED) < 0 Then PlainShort(s, b) = PlainShort(s, b) - RowOutput(r, OC_SIGNED)
                If RowOutput(r, OC_WEIGHTED) > 0 Then
                    WeightedLong(s, b) = WeightedLong(s, b) + RowOutput(r, OC_WEIGHTED)
                ElseIf RowOutput(r, OC_WEIGHTED) < 0 Then
                    WeightedShort(s, b) = WeightedShort(s, b) - RowOutput(r, OC_WEIGHTED)
                End If
            End If
        Next r

        ReDim LadderOutput(1 To SlotCount * BAND_COUNT, 1 To LADDER_WIDTH)
        ReDim ZoneOutput(1 To SlotCount + 1, 1 To ZONES_WIDTH)
        ReDim BandLong(1 To BAND_COUNT)
        ReDim BandShort(1 To BAND_COUNT)
        For s = 1 To SlotCount
            CurrencyCode = FxCodes(s)
            If CurrencyCode = GoldCode Then CurrencyCode = vbNullString
            For b = 1 To BAND_COUNT
                r = (s - 1) * BAND_COUNT + b
                LadderOutput(r, 1) = CurrencyCode
                LadderOutput(r, 2) = b
                LadderOutput(r, 3) = Ladder(b, LD_ZONE)
                If Len(LadderLabels(b, 1)) = 0 Then
                    LadderOutput(r, 4) = LBL_NOT_APPLY
                Else
                    LadderOutput(r, 4) = LadderLabels(b, 1)
                End If
                LadderOutput(r, 5) = LadderLabels(b, 2)
                LadderOutput(r, 6) = Ladder(b, LD_WEIGHT)
                If Len(CurrencyCode) > 0 Then
                    BandLong(b) = WeightedLong(s, b)
                    BandShort(b) = WeightedShort(s, b)
                    LadderOutput(r, 7) = PlainLong(s, b)
                    LadderOutput(r, 8) = PlainShort(s, b)
                Else
                    BandLong(b) = 0#
                    BandShort(b) = 0#
                    LadderOutput(r, 7) = 0#
                    LadderOutput(r, 8) = 0#
                End If
                LadderOutput(r, 9) = BandLong(b)
                LadderOutput(r, 10) = BandShort(b)
                LadderOutput(r, 11) = SmallerOf(BandLong(b), BandShort(b))
                LadderOutput(r, 12) = Parameters(PI_VERTICAL) * LadderOutput(r, 11)
                LadderOutput(r, 13) = BandLong(b) - BandShort(b)
            Next b

            '--- Zones of this currency (IR_Zones row s) ------------------------
            LadderCharges BandLong, BandShort, Ladder, Parameters, Components
            ZoneOutput(s, 1) = CurrencyCode
            If Len(CurrencyCode) > 0 Then
                ZoneOutput(s, 2) = SlotUsed(s)
                ZoneOutput(s, 29) = SlotSpecific(s)
            Else
                ZoneOutput(s, 2) = 0
                ZoneOutput(s, 29) = 0#
            End If
            For c = 1 To ZC_COUNT
                ZoneOutput(s, 2 + c) = Components(c)
            Next c
            ZoneOutput(s, 30) = Components(ZC_GENERAL) + ZoneOutput(s, 29)
            Totals(TOT_GENERAL) = Totals(TOT_GENERAL) + Components(ZC_GENERAL)
        Next s

        '--- Total row: sums, except the signed nets that do not add up -------
        ZoneOutput(SlotCount + 1, 1) = LBL_TOTAL_ZONES
        For c = 2 To ZONES_WIDTH
            Select Case c
                Case 8, 13, 18, 20, 21, 23, 24
                    ZoneOutput(SlotCount + 1, c) = Empty
                Case Else
                    ZoneOutput(SlotCount + 1, c) = 0#
                    For s = 1 To SlotCount
                        ZoneOutput(SlotCount + 1, c) = ZoneOutput(SlotCount + 1, c) + ZoneOutput(s, c)
                    Next s
            End Select
        Next c

End Sub


Public Function TimeBand( _
    ByVal ResidualMaturity As Double, _
    ByVal HighCoupon As Boolean, _
    ByRef Ladder() As Double) _
    As Long
'
'==============================================================================
'                                   TimeBand
'------------------------------------------------------------------------------
' PURPOSE
'   Return the maturity-method time band of a position.
'
' INPUTS
'   ResidualMaturity  Years to final maturity (fixed rate) or to the next
'                     repricing date (floating rate); any value, a negative
'                     one falls in band 1.
'   HighCoupon        True for coupon >= threshold or floating rate.
'   Ladder            (1 To BAND_COUNT, 1 To LD_WIDTH); upper bounds are
'                     inclusive, OPEN_BOUND for none.
'
' RETURNS
'   Band 1 to BAND_COUNT: 1 + the number of upper bounds below the maturity.
'
' REFERENCE
'   MAR40.26 Table 4 ("1 month or less", "1 to 3 months", ...).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Column      As Long      'Bound column used
    Dim Band        As Long      'Result
    Dim b           As Long      'Band

        If HighCoupon Then Column = LD_UPPER_HIGH Else Column = LD_UPPER_LOW
        Band = 1
        For b = 1 To BAND_COUNT
            If Ladder(b, Column) < ResidualMaturity Then Band = Band + 1
        Next b
        If Band > BAND_COUNT Then Band = BAND_COUNT
        TimeBand = Band

End Function


Public Sub LadderCharges( _
    ByRef WeightedLong() As Double, _
    ByRef WeightedShort() As Double, _
    ByRef Ladder() As Double, _
    ByRef Parameters() As Double, _
    ByRef Components() As Double)
'
'==============================================================================
'                                 LadderCharges
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the vertical and horizontal disallowances to one currency's
'   weighted positions and return every step of the maturity method.
'
' INPUTS
'   WeightedLong   (1 To BAND_COUNT) weighted long positions, >= 0,
'                  reporting currency.
'   WeightedShort  (1 To BAND_COUNT) weighted short positions, >= 0.
'   Ladder         (1 To BAND_COUNT, 1 To LD_WIDTH); LD_ZONE is used.
'   Parameters     Vector (1 To PI_COUNT); PI_VERTICAL, PI_ZONE1 to 3,
'                  PI_ADJACENT and PI_ZONE13 are used.
'
' RETURNS
'   Components(1 To ZC_COUNT), reporting currency: ZC_VERTICAL; per zone z
'   at ZC_ZONE_FIRST + 5 * (z - 1): net long, net short, matched, charge,
'   net; the zone 1-2, 2-3 and 1-3 offsets with the residuals after each;
'   ZC_BETWEEN; ZC_NET; ZC_GENERAL = the general market risk charge.
'
' REFERENCE
'   MAR40.27 (10% of the matched weighted position per band), MAR40.28 and
'   Table 5 (40% / 30% / 30% within zones, 40% between adjacent zones,
'   100% between zones 1 and 3, offsets in the order 1-2, 2-3, 1-3; 100%
'   of the residual net position).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim ZoneLong(1 To ZONE_COUNT)  As Double    'Net long band positions per zone
    Dim ZoneShort(1 To ZONE_COUNT) As Double    'Net short band positions per zone
    Dim ZoneNet(1 To ZONE_COUNT)   As Double    'Net position per zone
    Dim WithinRate  As Double      'Within-zone disallowance of the zone
    Dim BlockStart  As Long        'First component of the zone block
    Dim BandNet     As Double      'Net weighted position of a band
    Dim Matched     As Double      'An offset
    Dim Zone1       As Double      'Zone 1 net, updated by the offsets
    Dim Zone2       As Double      'Zone 2 net, updated by the offsets
    Dim Zone3       As Double      'Zone 3 net, updated by the offsets
    Dim b           As Long        'Band
    Dim z           As Long        'Zone

        ReDim Components(1 To ZC_COUNT)
        For b = 1 To BAND_COUNT
            Components(ZC_VERTICAL) = Components(ZC_VERTICAL) + Parameters(PI_VERTICAL) _
                * SmallerOf(WeightedLong(b), WeightedShort(b))
            BandNet = WeightedLong(b) - WeightedShort(b)
            z = CLng(Ladder(b, LD_ZONE))
            If BandNet > 0 Then
                ZoneLong(z) = ZoneLong(z) + BandNet
            ElseIf BandNet < 0 Then
                ZoneShort(z) = ZoneShort(z) - BandNet
            End If
        Next b

        For z = 1 To ZONE_COUNT
            Select Case z
                Case 1: WithinRate = Parameters(PI_ZONE1)
                Case 2: WithinRate = Parameters(PI_ZONE2)
                Case Else: WithinRate = Parameters(PI_ZONE3)
            End Select
            BlockStart = ZC_ZONE_FIRST + 5 * (z - 1)
            ZoneNet(z) = ZoneLong(z) - ZoneShort(z)
            Components(BlockStart) = ZoneLong(z)
            Components(BlockStart + 1) = ZoneShort(z)
            Components(BlockStart + 2) = SmallerOf(ZoneLong(z), ZoneShort(z))
            Components(BlockStart + 3) = WithinRate * Components(BlockStart + 2)
            Components(BlockStart + 4) = ZoneNet(z)
        Next z

        'Offsets between zones, in the order of MAR40.28.
        Zone1 = ZoneNet(1)
        Zone2 = ZoneNet(2)
        Zone3 = ZoneNet(3)
        Matched = Offset(Zone1, Zone2)
        Components(ZC_MATCHED12) = Matched
        Zone1 = Zone1 - Sgn(Zone1) * Matched
        Zone2 = Zone2 - Sgn(Zone2) * Matched
        Components(ZC_Z1_AFTER12) = Zone1
        Components(ZC_Z2_AFTER12) = Zone2
        Matched = Offset(Zone2, Zone3)
        Components(ZC_MATCHED23) = Matched
        Zone2 = Zone2 - Sgn(Zone2) * Matched
        Zone3 = Zone3 - Sgn(Zone3) * Matched
        Components(ZC_Z2_AFTER23) = Zone2
        Components(ZC_Z3_AFTER23) = Zone3
        Components(ZC_MATCHED13) = Offset(Zone1, Zone3)
        Components(ZC_BETWEEN) = Parameters(PI_ADJACENT) _
            * (Components(ZC_MATCHED12) + Components(ZC_MATCHED23)) _
            + Parameters(PI_ZONE13) * Components(ZC_MATCHED13)
        Components(ZC_NET) = Abs(ZoneNet(1) + ZoneNet(2) + ZoneNet(3))
        Components(ZC_GENERAL) = Components(ZC_VERTICAL) _
            + Components(ZC_ZONE_FIRST + 3) + Components(ZC_ZONE_FIRST + 8) _
            + Components(ZC_ZONE_FIRST + 13) + Components(ZC_BETWEEN) + Components(ZC_NET)

End Sub


'------------------------------------------------------------------------------
'                                    HELPERS
'------------------------------------------------------------------------------

Private Function RowStatus( _
    ByRef InputRows() As Variant, _
    ByVal r As Long, _
    ByVal CurrencyCode As String, _
    ByVal Category As String, _
    ByVal SrKey As String, _
    ByVal IsFloating As Boolean, _
    ByVal AsOfDate As Double, _
    ByRef FxCodes() As String, _
    ByVal GoldCode As String, _
    ByVal IdUses As Long) _
    As String
'
'==============================================================================
'                                   RowStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Return the status of one interest rate row: "OK" or the first reason
'   it is rejected, in a fixed order.
'
' INPUTS
'   InputRows, r      Input cells and the row index.
'   CurrencyCode     Normalised currency code.
'   Category     Normalised SR category.
'   SrKey        SR key from SpecificRiskKey.
'   IsFloating   True when Rate type is Floating.
'   AsOfDate     Reporting date serial.
'   FxCodes      FX table codes; GoldCode the gold code.
'   IdUses       InputRows sharing this PositionID (1 = unique).
'
' RETURNS
'   "OK" or "INVALID: reason".
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Maturity    As Variant   'Final maturity cell
    Dim Repricing   As Variant   'Next repricing cell
    Dim RateType    As String    'Normalised rate type

        Maturity = InputRows(r, IR_MATURITY)
        Repricing = InputRows(r, IR_REPRICING)
        RateType = CodeText(InputRows(r, IR_RATE_TYPE))
        If IdUses > 1 Then
            RowStatus = ST_INVALID & "duplicate PositionID"
        ElseIf FindCode(FxCodes, CurrencyCode) < 1 Then
            RowStatus = ST_INVALID & "currency not in FX_Table"
        ElseIf CurrencyCode = GoldCode Then
            RowStatus = ST_INVALID & "gold is FX risk; enter it on FXPositions"
        ElseIf DirectionSign(InputRows(r, IR_DIRECTION)) = 0 Then
            RowStatus = ST_INVALID & "Direction must be Long or Short"
        ElseIf Not IsNumberValue(InputRows(r, IR_VALUE)) Then
            RowStatus = ST_INVALID & "Position value must be a number"
        ElseIf NumberValue(InputRows(r, IR_VALUE)) < 0 Then
            RowStatus = ST_INVALID & "Position value must be >= 0 (Direction gives the sign)"
        ElseIf RateType <> "FIXED" And RateType <> "FLOATING" Then
            RowStatus = ST_INVALID & "Rate type must be Fixed or Floating"
        ElseIf Not IsNumberValue(Maturity) Then
            RowStatus = ST_INVALID & "Final maturity date missing"
        ElseIf NumberValue(Maturity) < AsOfDate Then
            RowStatus = ST_INVALID & "final maturity before AsOfDate"
        ElseIf Not IsFloating And Not IsNumberValue(InputRows(r, IR_COUPON)) Then
            RowStatus = ST_INVALID & "Coupon required for a fixed-rate position"
        ElseIf IsFloating And Not IsNumberValue(Repricing) Then
            RowStatus = ST_INVALID & "Next repricing date required for a floating-rate position"
        ElseIf IsFloating And (NumberValue(Repricing) < AsOfDate _
                               Or NumberValue(Repricing) > NumberValue(Maturity)) Then
            RowStatus = ST_INVALID & "repricing date outside AsOfDate to final maturity"
        ElseIf InStr(1, SR_CATEGORIES, "|" & Category & "|", vbBinaryCompare) = 0 _
               Or Len(Category) = 0 Then
            RowStatus = ST_INVALID & "SR category must be GOV, QUAL, OTHER or NONE"
        ElseIf SrKey = SR_NO_MATCH Then
            RowStatus = ST_INVALID & "no MAR40.6 charge for this category and rating grade"
        Else
            RowStatus = ST_OK
        End If

End Function


Private Function SpecificRiskBand( _
    ByVal Maturity As Variant, _
    ByVal AsOfDate As Double, _
    ByVal DaysPerYear As Double) _
    As String
'
'==============================================================================
'                               SpecificRiskBand
'------------------------------------------------------------------------------
' PURPOSE
'   Return the residual-maturity band of MAR40.6 Table 1.
'
' INPUTS
'   Maturity     Final maturity cell; a blank cell reads as 0 (as in Excel).
'   AsOfDate     Reporting date serial.
'   DaysPerYear  Days per year, > 0.
'
' RETURNS
'   "<=6m", "6m-24m", ">24m", or "" when Maturity is text.
'
' REFERENCE
'   MAR40.6 Table 1 (6 months or less; over 6 and up to 24 months; over 24).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Years       As Double    'Residual maturity to final maturity

        If Not IsNumberValue(Maturity) And Not IsEmpty(Maturity) Then
            SpecificRiskBand = vbNullString
            Exit Function
        End If
        Years = (NumberValue(Maturity) - AsOfDate) / DaysPerYear
        If Years <= 0.5 Then
            SpecificRiskBand = SRB_SHORT
        ElseIf Years <= 2 Then
            SpecificRiskBand = SRB_MEDIUM
        Else
            SpecificRiskBand = SRB_LONG
        End If

End Function


Private Function SpecificRiskKey( _
    ByVal Category As String, _
    ByVal Grade As String, _
    ByVal SrBand As String, _
    ByRef SrKeys() As String) _
    As String
'
'==============================================================================
'                                SpecificRiskKey
'------------------------------------------------------------------------------
' PURPOSE
'   Return the key of the Table 1 row that applies to a position: first the
'   maturity-specific key, then the "Any" key of the category and grade.
'
' INPUTS
'   Category   Normalised SR category.
'   Grade      Rating grade, trimmed.
'   SrBand     Band from SpecificRiskBand.
'   SrKeys     Table keys, upper case.
'
' RETURNS
'   "NONE" for category NONE; the matching key as built (category upper
'   case, grade as entered); or "NO MATCH".
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Candidate   As String    'Key tried

        If Category = SR_NONE Then
            SpecificRiskKey = SR_NONE
            Exit Function
        End If
        Candidate = Category & "|" & Grade & "|" & SrBand
        If FindCode(SrKeys, Candidate) >= 1 Then
            SpecificRiskKey = Candidate
            Exit Function
        End If
        Candidate = Category & "|" & Grade & "|" & SRB_ANY
        If FindCode(SrKeys, Candidate) >= 1 Then
            SpecificRiskKey = Candidate
        Else
            SpecificRiskKey = SR_NO_MATCH
        End If

End Function


Private Function IsHighCoupon( _
    ByRef InputRows() As Variant, _
    ByVal r As Long, _
    ByVal IsFloating As Boolean, _
    ByRef Parameters() As Double) _
    As Boolean
'
'==============================================================================
'                                 IsHighCoupon
'------------------------------------------------------------------------------
' PURPOSE
'   Tell which column of Table 4 slots a valid row.
'
' INPUTS
'   InputRows, r      Input cells and the row index (status OK).
'   IsFloating   True for a floating-rate row.
'   Parameters   Vector; PI_COUPON_THRESHOLD is used.
'
' RETURNS
'   True for a floating-rate row or a coupon >= threshold.
'
' REFERENCE
'   MAR40.26 Table 4.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        If IsFloating Then
            IsHighCoupon = True
        Else
            IsHighCoupon = (NumberValue(InputRows(r, IR_COUPON)) >= Parameters(PI_COUPON_THRESHOLD))
        End If

End Function


Private Function Offset( _
    ByVal FirstNet As Double, _
    ByVal SecondNet As Double) _
    As Double
'
'==============================================================================
'                                    Offset
'------------------------------------------------------------------------------
' PURPOSE
'   Return the amount two zone nets of opposite sign can offset.
'
' INPUTS
'   FirstNet, SecondNet   Signed zone nets, reporting currency.
'
' RETURNS
'   min(|FirstNet|, |SecondNet|) when the signs differ, else 0.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        If FirstNet * SecondNet < 0 Then
            Offset = SmallerOf(Abs(FirstNet), Abs(SecondNet))
        Else
            Offset = 0#
        End If

End Function


Private Function SmallerOf( _
    ByVal ValueA As Double, _
    ByVal ValueB As Double) _
    As Double
'
'==============================================================================
'                                SmallerOf
'------------------------------------------------------------------------------
' PURPOSE
'   Smaller of two numbers, without the Excel object model.
'
' RETURNS
'   min(ValueA, ValueB).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        If ValueA < ValueB Then SmallerOf = ValueA Else SmallerOf = ValueB

End Function
