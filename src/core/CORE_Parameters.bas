Attribute VB_Name = "CORE_Parameters"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Parameters
'------------------------------------------------------------------------------
' PURPOSE
'   Turn the raw Params values (as read from the cells) into the validated
'   parameter vector and tables the risk-class procedures use, and report every
'   unusable value as a Checks finding. Also provide the MAR40 reference
'   parameter set and ladder used by the SSA_ worksheet functions and by the
'   "constants changed" check
'
' WHY THIS EXISTS
'   The risk-class procedures need parameters that are checked once, against
'   their domains, before any calculation; a cell value that cannot be used is
'   reported as a finding and counted, and the caller stops the run when a count
'   is not zero. The same component holds the MAR40 reference values, so that
'   the worksheet functions and the check for changed constants read one source
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Builds and validates, from raw cell values:
'     - the scalar parameters: BuildParameters, BuildRegime and
'       TransitionalFactor
'     - the maturity ladder: BuildLadder
'     - the lookup tables: BuildSpecificRiskTable, BuildFxTable and
'       BuildOptionCharges
'
'   Provides the MAR40 reference values and counts the constants that differ
'   from them: ReferenceParameters, ReferenceLadder and ChangedParameterCount
'
' ERROR POLICY
'   Cell content never raises: an unusable value becomes a finding and is
'   counted in the procedure's return value; the caller stops the run when any
'   count is non-zero. ERR_ARGUMENT is raised for a raw array of the wrong
'   shape, which is a layout fault, not an input fault
'
' DEPENDENCIES
'   CORE_Config, CORE_Util
'
'   Reference: BCBS MAR40.1-40.2 (scaling, RWA), 40.6 Table 1, 40.26 Table 4,
'   40.27-40.28 Table 5, 40.43-40.47, 40.60-40.61, 40.72-40.73
'
' NOTES
'   Private module (Option Private Module): none of its members is visible
'   outside this VBA project. The module owns no state: every procedure depends
'   only on its arguments
'
'   Compatibility: Excel VBA; no references beyond the defaults. The module
'   never touches the Excel object model
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const PARAMS_SHEET      As String = "Params"    'Sheet named in findings
    Private Const VALUE_TOLERANCE   As Double = 0.000000001 'Equality of a constant to MAR40


'
'------------------------------------------------------------------------------
'
'                               PARAMETER VECTOR
'
'------------------------------------------------------------------------------
'

Public Function BuildParameters( _
    ByRef RawNumeric() As Variant, _
    ByVal RawAsOfDate As Variant, _
    ByVal RawReportingCcy As Variant, _
    ByVal RawGoldCode As Variant, _
    ByRef Parameters() As Double, _
    ByRef AsOfDate As Double, _
    ByRef ReportingCcy As String, _
    ByRef GoldCode As String, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                          VALIDATE SCALAR PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the scalar parameters and return them typed
'
' WHY THIS EXISTS
'   The parameters reach the workbook as raw cell values; the risk-class
'   procedures need them as numbers in a vector, each one checked against its
'   domain, with every unusable value reported as a finding
'
' INPUTS
'   RawNumeric
'     Cell values (1 To PI_COUNT), in the order of PRM_NUMERIC_CODES; read, not
'     modified
'
'   RawAsOfDate
'     Cell value of AsOfDate; must be a date (CORE_Util.IsDateSerial)
'
'   RawReportingCcy
'     Cell value of ReportingCcy; must not be blank
'
'   RawGoldCode
'     Cell value of GoldCode; must not be blank
'
'   Parameters
'     Output: the parameter vector (1 To PI_VECTOR) as decimals, years and
'     factors
'
'   AsOfDate
'     Output: the reporting date as a date serial
'
'   ReportingCcy
'     Output: the reporting currency, upper case
'
'   GoldCode
'     Output: the gold code, upper case
'
'   Findings
'     Findings list, appended to (CORE_Util.AddFinding)
'
'   FindingCount
'     Findings in use; incremented
'
' RETURNS
'   Number of unusable values. Output parameters: Parameters(1 To PI_VECTOR)
'   as decimals / years / factors; AsOfDate as a date serial; ReportingCcy and
'   GoldCode upper case
'
' BEHAVIOR
'   Checks the shape of RawNumeric, then in order:
'     - reads the as-of date, the reporting currency and the gold code
'     - reads each numeric parameter and checks it against its domain letter in
'       PRM_DOMAINS
'     - returns the number of unusable values
'
'   Domains: P: > 0; N: >= 0; R: 0 <= rate <= 1
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when RawNumeric is not (1 To PI_COUNT). An unusable
'   value does not raise: it becomes a finding and is counted
'
' DEPENDENCIES
'   CORE_Util: IsDateSerial, NumberValue, CodeText, ListItem, IsNumberValue;
'   AddError of this module; CORE_Config: ERR_ARGUMENT, PI_COUNT, PI_VECTOR,
'   PRM_NUMERIC_CODES, PRM_DOMAINS, PRM_ASOF, PRM_REPCCY, PRM_GOLD
'
'   Reference: PRM_DOMAINS (P: > 0; N: >= 0; R: 0 <= rate <= 1)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrorCount  As Long     'Unusable values found
    Dim Code        As String   'Parameter code
    Dim Domain      As String   'Domain letter of the parameter
    Dim Value       As Double   'Numeric value
    Dim i           As Long     'Parameter index

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a raw array of the wrong shape, a layout fault and not an input fault
        If LBound(RawNumeric) <> 1 Or UBound(RawNumeric) <> PI_COUNT Then
            Err.Raise ERR_ARGUMENT, "CORE_Parameters.BuildParameters", _
                "RawNumeric must be (1 To " & PI_COUNT & ")."
        End If

'------------------------------------------------------------------------------
' PREPARE OUTPUT
'------------------------------------------------------------------------------
    'Size the vector for the parameters and the regime flag
        ReDim Parameters(1 To PI_VECTOR)

'------------------------------------------------------------------------------
' READ TEXT AND DATE SCALARS
'------------------------------------------------------------------------------
    'Read the reporting date, which must be a date
        If IsDateSerial(RawAsOfDate) Then
            AsOfDate = NumberValue(RawAsOfDate)
        Else
            AddError Findings, FindingCount, ErrorCount, PRM_ASOF, "must be a date"
        End If
    'Read the reporting currency, which must not be blank
        ReportingCcy = CodeText(RawReportingCcy)
        If Len(ReportingCcy) = 0 Then
            AddError Findings, FindingCount, ErrorCount, PRM_REPCCY, "must not be blank"
        End If
    'Read the gold code, which must not be blank
        GoldCode = CodeText(RawGoldCode)
        If Len(GoldCode) = 0 Then
            AddError Findings, FindingCount, ErrorCount, PRM_GOLD, "must not be blank"
        End If

'------------------------------------------------------------------------------
' CHECK NUMERIC PARAMETERS
'------------------------------------------------------------------------------
    'Read each numeric parameter and check it against the domain letter of its code
        For i = 1 To PI_COUNT
            Code = ListItem(PRM_NUMERIC_CODES, i)
            Domain = Mid$(PRM_DOMAINS, i, 1)
            If Not IsNumberValue(RawNumeric(i)) Then
                AddError Findings, FindingCount, ErrorCount, Code, "must be a number"
            Else
                Value = NumberValue(RawNumeric(i))
                Parameters(i) = Value
                Select Case Domain
                    Case "P"
                        If Value <= 0 Then
                            AddError Findings, FindingCount, ErrorCount, Code, _
                                "must be greater than 0"
                        End If
                    Case "N"
                        If Value < 0 Then
                            AddError Findings, FindingCount, ErrorCount, Code, _
                                "must be 0 or more"
                        End If
                    Case Else
                        If Value < 0 Or Value > 1 Then
                            AddError Findings, FindingCount, ErrorCount, Code, _
                                "must be a decimal rate from 0 to 1"
                        End If
                End Select
            End If
        Next i

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of unusable values
        BuildParameters = ErrorCount

End Function

Public Function BuildRegime( _
    ByVal RawRegime As Variant, _
    ByVal RawEndDate As Variant, _
    ByVal RawOwnFunds As Variant, _
    ByRef Regime As String, _
    ByRef EndDate As Double, _
    ByRef OwnFunds As Double, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                               VALIDATE REGIME
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the regime, the last date of the transitional multiplier and the
'   total own funds
'
' WHY THIS EXISTS
'   The regime selects Basel or CRR treatment; under the CRR the transitional
'   end date and the total own funds must be usable, and under Basel they are
'   still read so that a change of them is counted by ChangedParameterCount
'
' INPUTS
'   RawRegime
'     Cell value of Regime; Basel or CRR, without regard to case
'
'   RawEndDate
'     Cell value of TransitionalEndDate; read when it is a date, whatever the
'     regime, so that a change of it is counted by ChangedParameterCount; it
'     must be a date under the CRR
'
'   RawOwnFunds
'     Cell value of TotalOwnFunds; read when it is a number, whatever the
'     regime; it must be a number greater than 0 under the CRR
'
'   Regime
'     Output: REGIME_BASEL or REGIME_CRR
'
'   EndDate
'     Output: the last date of the multiplier as a date serial
'
'   OwnFunds
'     Output: the total own funds as a number
'
'   Findings
'     Findings list, appended to (CORE_Util.AddFinding)
'
'   FindingCount
'     Findings in use; incremented
'
' RETURNS
'   Number of unusable values. Output parameters: Regime as REGIME_BASEL or
'   REGIME_CRR (REGIME_BASEL when it is unusable); EndDate as a date serial, 0
'   when RawEndDate is not a date; OwnFunds as a number, 0 when RawOwnFunds is
'   not a number
'
' BEHAVIOR
'   Sets the outputs to the Basel regime, 0 and 0, then in order:
'     - reads the regime and reports a value that is neither Basel nor CRR
'     - reads the end date, and reports it when it is not a date under the CRR
'     - reads the own funds, and reports them when they are not greater than 0
'       under the CRR
'     - returns the number of unusable values
'
' ERROR POLICY
'   Never raises: an unusable value becomes a finding and is counted
'
' DEPENDENCIES
'   CORE_Util: CodeText, IsDateSerial, NumberValue, IsNumberValue; AddError of
'   this module; CORE_Config: REGIME_BASEL, REGIME_CRR, PRM_REGIME,
'   PRM_TRANSITION_END, PRM_OWN_FUNDS
'
'   Reference: Article 325(2), Article 351 and Article 495s(2) of Regulation (EU)
'   No 575/2013, as amended by Regulation (EU) 2024/1623 and Delegated
'   Regulation (EU) 2026/1221
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
    Dim ErrorCount  As Long 'Unusable values found

'------------------------------------------------------------------------------
' INITIALIZE OUTPUT
'------------------------------------------------------------------------------
    'Default to the Basel regime, with no end date and no own funds
        Regime = REGIME_BASEL
        EndDate = 0#
        OwnFunds = 0#

'------------------------------------------------------------------------------
' READ REGIME
'------------------------------------------------------------------------------
    'Accept Basel or CRR and report any other value
        Select Case CodeText(RawRegime)
            Case REGIME_BASEL
                Regime = REGIME_BASEL
            Case REGIME_CRR
                Regime = REGIME_CRR
            Case Else
                AddError Findings, FindingCount, ErrorCount, PRM_REGIME, "must be Basel or CRR"
        End Select

'------------------------------------------------------------------------------
' READ END DATE
'------------------------------------------------------------------------------
    'Read the end date when it is a date, and require it under the CRR
        If IsDateSerial(RawEndDate) Then
            EndDate = NumberValue(RawEndDate)
        ElseIf Regime = REGIME_CRR Then
            AddError Findings, FindingCount, ErrorCount, PRM_TRANSITION_END, "must be a date"
        End If

'------------------------------------------------------------------------------
' READ OWN FUNDS
'------------------------------------------------------------------------------
    'Read the own funds when they are a number
        If IsNumberValue(RawOwnFunds) Then OwnFunds = NumberValue(RawOwnFunds)
    'Require own funds greater than 0 under the CRR
        If Regime = REGIME_CRR And OwnFunds <= 0 Then
            AddError Findings, FindingCount, ErrorCount, PRM_OWN_FUNDS, "must be a number greater than 0"
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of unusable values
        BuildRegime = ErrorCount

End Function

Public Function TransitionalFactor( _
    ByVal Regime As String, _
    ByVal AsOfDate As Double, _
    ByVal EndDate As Double, _
    ByVal Multiplier As Double) _
    As Double
'
'------------------------------------------------------------------------------
'                          RETURN TRANSITIONAL FACTOR
'------------------------------------------------------------------------------
' PURPOSE
'   Return the factor that the requirement of the simplified standardised
'   approach is multiplied by on a date: the transitional multiplier under the
'   CRR until its last date, and 1 under the Basel regime and after that date
'
' WHY THIS EXISTS
'   The CRR applies a transitional multiplier to the requirement up to a last
'   date; the choice between the multiplier and 1 is made in one place
'
' INPUTS
'   Regime
'     REGIME_BASEL or REGIME_CRR, from BuildRegime
'
'   AsOfDate
'     Reporting date, a date serial
'
'   EndDate
'     Last date of the multiplier, a date serial
'
'   Multiplier
'     The multiplier, a decimal; PI_TRANSITIONAL
'
' RETURNS
'   Multiplier when Regime is REGIME_CRR and AsOfDate is on or before EndDate;
'   1 otherwise
'
' BEHAVIOR
'   Compares the regime and the reporting date with the end date and returns
'   the multiplier or 1
'
' ERROR POLICY
'   Never raises; the arguments are not validated
'
' DEPENDENCIES
'   CORE_Config: REGIME_CRR
'
'   Reference: Article 495s(2) of Regulation (EU) No 575/2013, from Delegated
'   Regulation (EU) 2026/1221
'
' NOTES
'   None
'
' UPDATED
'   2026-11-20
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Use the multiplier under the CRR up to and including the end date, and 1 otherwise
        If Regime = REGIME_CRR And AsOfDate <= EndDate Then
            TransitionalFactor = Multiplier
        Else
            TransitionalFactor = 1#
        End If

End Function

Public Function ChangedParameterCount( _
    ByRef Parameters() As Double, _
    ByVal GoldCode As String, _
    ByVal TransitionEnd As Double) _
    As Long
'
'------------------------------------------------------------------------------
'                           COUNT CHANGED CONSTANTS
'------------------------------------------------------------------------------
' PURPOSE
'   Count the constants that differ from their reference value (a what-if run)
'
' WHY THIS EXISTS
'   A run with changed constants is not a run of the MAR40 reference set; the
'   count lets the caller report that
'
' INPUTS
'   Parameters
'     Vector (1 To PI_COUNT) from BuildParameters; read, not modified
'
'   GoldCode
'     Gold code in use, upper case
'
'   TransitionEnd
'     Last date of the transitional multiplier from BuildRegime, 0 when it is
'     not a date
'
' RETURNS
'   Number of numeric parameters differing from PRM_REFERENCE_VALUES by more
'   than 1E-9, plus 1 when GoldCode is not XAU, plus 1 when TransitionEnd is
'   not PRM_REFERENCE_END_DATE
'
' BEHAVIOR
'   Loads the reference vector, then in order:
'     - counts each numeric parameter that differs by more than VALUE_TOLERANCE
'     - adds 1 when the gold code is not PRM_REFERENCE_GOLD
'     - adds 1 when the transitional end date differs from
'       PRM_REFERENCE_END_DATE by more than VALUE_TOLERANCE
'
' ERROR POLICY
'   No validation of its own: Parameters must hold at least PI_COUNT items, and
'   a shorter vector raises the VBA subscript error
'
' DEPENDENCIES
'   ReferenceParameters of this module; CORE_Config: PI_COUNT,
'   PRM_REFERENCE_GOLD, PRM_REFERENCE_END_DATE
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
    Dim Reference() As Double   'MAR40 values
    Dim Changed     As Long     'Count
    Dim i           As Long     'Parameter index

'------------------------------------------------------------------------------
' LOAD REFERENCE PARAMETERS
'------------------------------------------------------------------------------
    'Load the MAR40 reference vector
        ReferenceParameters Reference

'------------------------------------------------------------------------------
' COUNT CHANGES
'------------------------------------------------------------------------------
    'Count each numeric parameter that differs from its reference value
        For i = 1 To PI_COUNT
            If Abs(Parameters(i) - Reference(i)) > VALUE_TOLERANCE Then Changed = Changed + 1
        Next i
    'Count a gold code other than the reference one
        If GoldCode <> PRM_REFERENCE_GOLD Then Changed = Changed + 1
    'Count a transitional end date other than the reference one
        If Abs(TransitionEnd - PRM_REFERENCE_END_DATE) > VALUE_TOLERANCE Then Changed = Changed + 1

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of changed constants
        ChangedParameterCount = Changed

End Function

Public Sub ReferenceParameters( _
    ByRef Parameters() As Double)
'
'------------------------------------------------------------------------------
'                         RETURN REFERENCE PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Return the MAR40 reference parameter vector
'
' WHY THIS EXISTS
'   The SSA_ worksheet functions and the check for changed constants need the
'   MAR40 values without reading the Params cells
'
' INPUTS
'   Parameters
'     Output: the reference vector (1 To PI_VECTOR)
'
' RETURNS
'   Nothing. Parameters(1 To PI_VECTOR): the first PI_COUNT items from
'   PRM_REFERENCE_VALUES, and 0 (the Basel regime) in PI_REGIME_FLAG
'
' BEHAVIOR
'   Parses the reference list into the first PI_COUNT items, then extends the
'   vector to PI_VECTOR, which leaves the regime flag at 0
'
' ERROR POLICY
'   No error handling of its own: an error raised by ParseList passes to the
'   caller
'
' DEPENDENCIES
'   CORE_Util: ParseList; CORE_Config: PRM_REFERENCE_VALUES, PI_COUNT,
'   PI_VECTOR
'
'   Reference: MAR40.1-40.2, 40.27-40.28, 40.43, 40.47, 40.61, 40.72-40.73
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' LOAD REFERENCE VALUES
'------------------------------------------------------------------------------
    'Parse the reference values into the first PI_COUNT items
        ParseList PRM_REFERENCE_VALUES, PI_COUNT, 0#, Parameters

'------------------------------------------------------------------------------
' ADD REGIME FLAG
'------------------------------------------------------------------------------
    'Extend the vector to the regime flag item, keeping the parsed values
        ReDim Preserve Parameters(1 To PI_VECTOR)

End Sub

'
'------------------------------------------------------------------------------
'
'                               MATURITY LADDER
'
'------------------------------------------------------------------------------
'

Public Function BuildLadder( _
    ByRef RawLadder() As Variant, _
    ByRef Ladder() As Double, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                               VALIDATE LADDER
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the maturity-ladder table of Params and return it as numbers
'
' WHY THIS EXISTS
'   The ladder reaches the workbook as raw cells; the debt procedures need the
'   zones, the bounds and the risk weights as numbers that are known to be in
'   order
'
' INPUTS
'   RawLadder
'     Cell values (1 To BAND_COUNT, 1 To LD_WIDTH): zone, upper bound for
'     coupon >= threshold, upper bound for coupon below it (years, blank =
'     open-ended or absent), risk weight
'
'   Ladder
'     Output: the ladder (1 To BAND_COUNT, 1 To LD_WIDTH) as numbers
'
'   Findings
'     Findings list, appended to
'
'   FindingCount
'     Findings in use; incremented
'
' RETURNS
'   Number of unusable values. Ladder(1 To BAND_COUNT, 1 To LD_WIDTH) as
'   numbers, OPEN_BOUND for a blank bound
'
' BEHAVIOR
'   Checks the shape of RawLadder, then in order:
'     - checks the zone and the risk weight of each band
'     - checks each upper-bound column, where a blank bound is open and a
'       finite bound must exceed the previous finite bound of the column
'     - returns the number of unusable values
'
'   Zones must be 1 to 3 and non-decreasing; bounds must increase from band to
'   band; weights must be 0 or more
'
' ERROR POLICY
'   Raises ERR_ARGUMENT for a RawLadder of another shape. An unusable value does
'   not raise: it becomes a finding and is counted
'
' DEPENDENCIES
'   CORE_Util: IsNumberValue, NumberValue, PlainText; AddError of this module;
'   CORE_Config: ERR_ARGUMENT, BAND_COUNT, LD_WIDTH, LD_ZONE, LD_WEIGHT,
'   LD_UPPER_HIGH, LD_UPPER_LOW, OPEN_BOUND
'
'   Reference: MAR40.26 Table 4
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
    Dim ErrorCount  As Long     'Unusable values found
    Dim Previous    As Double   'Previous finite bound in the column
    Dim Value       As Variant  'Cell value
    Dim Item        As String   'Finding item, "Ladder band n"
    Dim b           As Long     'Band
    Dim c           As Long     'Bound column

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a raw ladder of another shape
        If UBound(RawLadder, 1) <> BAND_COUNT Or UBound(RawLadder, 2) <> LD_WIDTH Then
            Err.Raise ERR_ARGUMENT, "CORE_Parameters.BuildLadder", _
                "RawLadder must be (1 To " & BAND_COUNT & ", 1 To " & LD_WIDTH & ")."
        End If

'------------------------------------------------------------------------------
' PREPARE OUTPUT
'------------------------------------------------------------------------------
    'Size the ladder
        ReDim Ladder(1 To BAND_COUNT, 1 To LD_WIDTH)

'------------------------------------------------------------------------------
' CHECK ZONES AND WEIGHTS
'------------------------------------------------------------------------------
    'Check the zone and the risk weight of each band
        For b = 1 To BAND_COUNT
            Item = "Ladder band " & b
            Value = RawLadder(b, LD_ZONE)
            If Not IsNumberValue(Value) Then
                AddError Findings, FindingCount, ErrorCount, Item, "zone must be 1, 2 or 3"
            ElseIf NumberValue(Value) <> 1 And NumberValue(Value) <> 2 And NumberValue(Value) <> 3 Then
                AddError Findings, FindingCount, ErrorCount, Item, "zone must be 1, 2 or 3"
            Else
                Ladder(b, LD_ZONE) = NumberValue(Value)
                If b > 1 Then
                    If Ladder(b, LD_ZONE) < Ladder(b - 1, LD_ZONE) Then
                        AddError Findings, FindingCount, ErrorCount, Item, _
                            "zones must not decrease from band to band"
                    End If
                End If
            End If
            Value = RawLadder(b, LD_WEIGHT)
            If Not IsNumberValue(Value) Then
                AddError Findings, FindingCount, ErrorCount, Item, "risk weight must be a number"
            ElseIf NumberValue(Value) < 0 Then
                AddError Findings, FindingCount, ErrorCount, Item, "risk weight must be 0 or more"
            Else
                Ladder(b, LD_WEIGHT) = NumberValue(Value)
            End If
        Next b

'------------------------------------------------------------------------------
' CHECK UPPER BOUNDS
'------------------------------------------------------------------------------
    'Check each bound column: a blank means open and finite bounds must increase
        For c = LD_UPPER_HIGH To LD_UPPER_LOW
            Previous = -1#
            For b = 1 To BAND_COUNT
                Value = RawLadder(b, c)
                If IsEmpty(Value) Or PlainText(Value) = vbNullString Then
                    Ladder(b, c) = OPEN_BOUND
                ElseIf Not IsNumberValue(Value) Then
                    AddError Findings, FindingCount, ErrorCount, "Ladder band " & b, _
                        "upper bound must be a number of years or blank"
                ElseIf NumberValue(Value) <= Previous Then
                    AddError Findings, FindingCount, ErrorCount, "Ladder band " & b, _
                        "upper bounds must increase from band to band"
                Else
                    Ladder(b, c) = NumberValue(Value)
                    Previous = Ladder(b, c)
                End If
            Next b
        Next c

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of unusable values
        BuildLadder = ErrorCount

End Function

Public Sub ReferenceLadder( _
    ByRef Ladder() As Double)
'
'------------------------------------------------------------------------------
'                           RETURN REFERENCE LADDER
'------------------------------------------------------------------------------
' PURPOSE
'   Return the MAR40 Table 4 ladder
'
' WHY THIS EXISTS
'   The SSA_ worksheet functions and the debt procedures need the MAR40 ladder
'   without reading the Params cells
'
' INPUTS
'   Ladder
'     Output: the ladder (1 To BAND_COUNT, 1 To LD_WIDTH)
'
' RETURNS
'   Nothing. Ladder(1 To BAND_COUNT, 1 To LD_WIDTH), as BuildLadder returns it
'
' BEHAVIOR
'   Parses the zone, the two upper-bound and the risk-weight lists in turn and
'   copies each into its column, then sets the first band's bounds to one month
'
' ERROR POLICY
'   No error handling of its own: an error raised by ParseList passes to the
'   caller
'
' DEPENDENCIES
'   CORE_Util: ParseList; CORE_Config: BAND_COUNT, LD_WIDTH, LD_ZONE,
'   LD_UPPER_HIGH, LD_UPPER_LOW, LD_WEIGHT, LADDER_ZONES, LADDER_UPPER_HIGH,
'   LADDER_UPPER_LOW, LADDER_WEIGHTS, OPEN_BOUND
'
'   Reference: MAR40.26 Table 4
'
' NOTES
'   The upper bound of the first band is set to 1 / 12 after the lists are
'   copied: one month is 1/12 of a year exactly, and a decimal literal falls
'   short of it
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Column()    As Double   'One parsed list
    Dim b           As Long     'Band

'------------------------------------------------------------------------------
' PREPARE OUTPUT
'------------------------------------------------------------------------------
    'Size the ladder
        ReDim Ladder(1 To BAND_COUNT, 1 To LD_WIDTH)

'------------------------------------------------------------------------------
' LOAD LADDER COLUMNS
'------------------------------------------------------------------------------
    'Copy the zones
        ParseList LADDER_ZONES, BAND_COUNT, 0#, Column
        For b = 1 To BAND_COUNT
            Ladder(b, LD_ZONE) = Column(b)
        Next b
    'Copy the upper bounds for coupon >= threshold, open-ended where absent
        ParseList LADDER_UPPER_HIGH, BAND_COUNT, OPEN_BOUND, Column
        For b = 1 To BAND_COUNT
            Ladder(b, LD_UPPER_HIGH) = Column(b)
        Next b
    'Copy the upper bounds for coupon below the threshold, open-ended where absent
        ParseList LADDER_UPPER_LOW, BAND_COUNT, OPEN_BOUND, Column
        For b = 1 To BAND_COUNT
            Ladder(b, LD_UPPER_LOW) = Column(b)
        Next b
    'Copy the risk weights
        ParseList LADDER_WEIGHTS, BAND_COUNT, 0#, Column
        For b = 1 To BAND_COUNT
            Ladder(b, LD_WEIGHT) = Column(b)
        Next b

'------------------------------------------------------------------------------
' SET FIRST BAND
'------------------------------------------------------------------------------
    'One month is 1/12 of a year exactly; a decimal literal falls short of it
        Ladder(1, LD_UPPER_HIGH) = 1# / 12#
        Ladder(1, LD_UPPER_LOW) = 1# / 12#

End Sub

'
'------------------------------------------------------------------------------
'
'                                LOOKUP TABLES
'
'------------------------------------------------------------------------------
'

Public Function BuildSpecificRiskTable( _
    ByRef RawTable() As Variant, _
    ByRef Keys() As String, _
    ByRef Charges() As Double, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                         VALIDATE SPECIFIC RISK TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the specific-risk table (key "CATEGORY|grade|maturity band",
'   charge) and return it typed
'
' WHY THIS EXISTS
'   The specific-risk charges reach the workbook as raw cells; the debt
'   procedures look them up by key, so the keys must be unique and the charges
'   must be decimal rates
'
' INPUTS
'   RawTable
'     Cell values (1 To n, 1 To 2): key, charge
'
'   Keys
'     Output: the keys (1 To n), upper case, "" for a blank row
'
'   Charges
'     Output: the charges (1 To n) as decimals
'
'   Findings
'     Findings list, appended to
'
'   FindingCount
'     Findings in use; incremented
'
' RETURNS
'   Number of unusable values. Keys(1 To n) upper case ("" for a blank row) and
'   Charges(1 To n) as decimals
'
' BEHAVIOR
'   For each row with a key, in order:
'     - checks that the charge is a number from 0 to 1
'     - checks that the key is not listed in an earlier row
'
' ERROR POLICY
'   Never raises: an unusable value becomes a finding and is counted
'
' DEPENDENCIES
'   CORE_Util: CodeText, IsNumberValue, NumberValue; AddError of this module
'
'   Reference: MAR40.6 Table 1; charges must be decimals from 0 to 1, keys
'   unique
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
    Dim ErrorCount  As Long 'Unusable values found
    Dim RowCount    As Long 'Rows of the table
    Dim r           As Long 'Table row
    Dim j           As Long 'Earlier row

'------------------------------------------------------------------------------
' PREPARE OUTPUT
'------------------------------------------------------------------------------
    'Size the keys and the charges to the rows of the table
        RowCount = UBound(RawTable, 1)
        ReDim Keys(1 To RowCount)
        ReDim Charges(1 To RowCount)

'------------------------------------------------------------------------------
' CHECK ROWS
'------------------------------------------------------------------------------
    'Check the charge and the uniqueness of the key of each row that has a key
        For r = 1 To RowCount
            Keys(r) = CodeText(RawTable(r, 1))
            If Len(Keys(r)) > 0 Then
                If Not IsNumberValue(RawTable(r, 2)) Then
                    AddError Findings, FindingCount, ErrorCount, Keys(r), "charge must be a number"
                ElseIf NumberValue(RawTable(r, 2)) < 0 Or NumberValue(RawTable(r, 2)) > 1 Then
                    AddError Findings, FindingCount, ErrorCount, Keys(r), _
                        "charge must be a decimal rate from 0 to 1"
                Else
                    Charges(r) = NumberValue(RawTable(r, 2))
                End If
                For j = 1 To r - 1
                    If Keys(j) = Keys(r) Then
                        AddError Findings, FindingCount, ErrorCount, Keys(r), _
                            "specific-risk key listed more than once"
                        Exit For
                    End If
                Next j
            End If
        Next r

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of unusable values
        BuildSpecificRiskTable = ErrorCount

End Function

Public Function BuildFxTable( _
    ByRef RawTable() As Variant, _
    ByVal ReportingCcy As String, _
    ByRef Codes() As String, _
    ByRef Rates() As Double, _
    ByRef ReportingErrors As Long, _
    ByRef DuplicateErrors As Long, _
    ByRef RateErrors As Long, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                              VALIDATE FX TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the FX table (currency, units of reporting currency per unit) and
'   return it typed
'
' WHY THIS EXISTS
'   The FX rates reach the workbook as raw cells; the reporting currency must
'   be listed at rate 1, a currency must not be listed twice, and every rate
'   must be a number greater than 0. The three kinds of fault are counted
'   separately, one count for each rule
'
' INPUTS
'   RawTable
'     Cell values (1 To n, 1 To 2): currency code, rate
'
'   ReportingCcy
'     Reporting currency, upper case
'
'   Codes
'     Output: the currency codes (1 To n), upper case, "" for a blank row
'
'   Rates
'     Output: the rates (1 To n), units of reporting currency per unit
'
'   ReportingErrors
'     Output: 0 or 1; 1 when the reporting currency is missing or not at rate 1
'     (rule C05)
'
'   DuplicateErrors
'     Output: codes listed again (rule C06)
'
'   RateErrors
'     Output: rates missing, not numbers or not > 0 (rule C07)
'
'   Findings
'     Findings list, appended to
'
'   FindingCount
'     Findings in use; incremented
'
' RETURNS
'   Total unusable values. Codes(1 To n) upper case ("" for a blank row),
'   Rates(1 To n); ReportingErrors (0 or 1: reporting currency missing or not at
'   rate 1, rule C05), DuplicateErrors (codes listed again, C06), RateErrors
'   (rates missing, not numbers or not > 0, C07)
'
' BEHAVIOR
'   Sets the three counts to 0, then in order:
'     - for each row with a code, checks the rate and looks for an earlier row
'       with the same code
'     - looks for the reporting currency in the codes and checks that its rate
'       is 1
'     - returns the sum of the three counts
'
' ERROR POLICY
'   Never raises: an unusable value becomes a finding and is counted
'
' DEPENDENCIES
'   CORE_Util: CodeText, IsNumberValue, NumberValue, FindCode, AddFinding;
'   CORE_Config: ST_INVALID
'
' NOTES
'   The findings name the sheet PARAMS_SHEET
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RowCount    As Long 'Rows of the table
    Dim Position    As Long 'Row of the reporting currency
    Dim r           As Long 'Table row
    Dim j           As Long 'Earlier row

'------------------------------------------------------------------------------
' PREPARE OUTPUT
'------------------------------------------------------------------------------
    'Size the codes and the rates to the rows of the table
        RowCount = UBound(RawTable, 1)
        ReDim Codes(1 To RowCount)
        ReDim Rates(1 To RowCount)
    'Start the three counts at zero
        ReportingErrors = 0
        DuplicateErrors = 0
        RateErrors = 0

'------------------------------------------------------------------------------
' CHECK ROWS
'------------------------------------------------------------------------------
    'Check the rate and the uniqueness of the code of each row that has a code
        For r = 1 To RowCount
            Codes(r) = CodeText(RawTable(r, 1))
            If Len(Codes(r)) > 0 Then
                If Not IsNumberValue(RawTable(r, 2)) Then
                    RateErrors = RateErrors + 1
                    AddFinding Findings, FindingCount, PARAMS_SHEET, Codes(r), "", _
                        ST_INVALID & "FX rate must be a number"
                ElseIf NumberValue(RawTable(r, 2)) <= 0 Then
                    RateErrors = RateErrors + 1
                    AddFinding Findings, FindingCount, PARAMS_SHEET, Codes(r), "", _
                        ST_INVALID & "FX rate must be greater than 0"
                Else
                    Rates(r) = NumberValue(RawTable(r, 2))
                End If
                For j = 1 To r - 1
                    If Codes(j) = Codes(r) Then
                        DuplicateErrors = DuplicateErrors + 1
                        AddFinding Findings, FindingCount, PARAMS_SHEET, Codes(r), "", _
                            ST_INVALID & "currency listed more than once in FX_Table"
                        Exit For
                    End If
                Next j
            End If
        Next r

'------------------------------------------------------------------------------
' CHECK REPORTING CURRENCY
'------------------------------------------------------------------------------
    'Require the reporting currency to be listed at rate 1
        Position = FindCode(Codes, ReportingCcy)
        If Position < 1 Then
            ReportingErrors = 1
        ElseIf Rates(Position) <> 1 Then
            ReportingErrors = 1
        End If
    'Report the reporting currency once, whatever the fault
        If ReportingErrors > 0 Then
            AddFinding Findings, FindingCount, PARAMS_SHEET, ReportingCcy, "", _
                ST_INVALID & "reporting currency must be listed in FX_Table with rate 1"
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the total of the three counts
        BuildFxTable = ReportingErrors + DuplicateErrors + RateErrors

End Function

Public Function BuildOptionCharges( _
    ByRef RawCharges() As Variant, _
    ByRef Charges() As Double, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                           VALIDATE OPTION CHARGES
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the non-delta option charges entered per risk class
'
' WHY THIS EXISTS
'   The option charges are entered in one cell per risk class and are added
'   before scaling (MAR40.2); a blank cell means no charge, and any other value must
'   be a number of 0 or more
'
' INPUTS
'   RawCharges
'     Cell values (1 To RC_COUNT), in RC_ order; blank = 0
'
'   Charges
'     Output: the charges (1 To RC_COUNT) in reporting currency
'
'   Findings
'     Findings list, appended to
'
'   FindingCount
'     Findings in use; incremented
'
' RETURNS
'   Number of unusable values. Charges(1 To RC_COUNT) in reporting currency
'
' BEHAVIOR
'   For each risk class, in order: a blank cell gives 0, a value that is not a
'   number or is negative is reported and counted, and any other value is kept
'
' ERROR POLICY
'   Never raises: an unusable value becomes a finding and is counted
'
' DEPENDENCIES
'   CORE_Util: ListItem, IsNumberValue, NumberValue, AddFinding; CORE_Config:
'   RC_COUNT, RC_CODES, NM_OPTIONS_PREFIX, SH_OPTIONS, ST_INVALID
'
'   Reference: MAR40.2 (option charges added before scaling), MAR40.74-40.84
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
    Dim ErrorCount  As Long     'Unusable values found
    Dim Item        As String   'Finding item, workbook name
    Dim i           As Long     'Risk class

'------------------------------------------------------------------------------
' PREPARE OUTPUT
'------------------------------------------------------------------------------
    'Size the charges to the risk classes
        ReDim Charges(1 To RC_COUNT)

'------------------------------------------------------------------------------
' CHECK CHARGES
'------------------------------------------------------------------------------
    'Check the charge of each risk class, where a blank cell is 0
        For i = 1 To RC_COUNT
            Item = NM_OPTIONS_PREFIX & ListItem(RC_CODES, i)
            If IsEmpty(RawCharges(i)) Then
                Charges(i) = 0#
            ElseIf Not IsNumberValue(RawCharges(i)) Then
                ErrorCount = ErrorCount + 1
                AddFinding Findings, FindingCount, SH_OPTIONS, Item, "", ST_INVALID & "must be a number"
            ElseIf NumberValue(RawCharges(i)) < 0 Then
                ErrorCount = ErrorCount + 1
                AddFinding Findings, FindingCount, SH_OPTIONS, Item, "", ST_INVALID & "must be 0 or more"
            Else
                Charges(i) = NumberValue(RawCharges(i))
            End If
        Next i

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of unusable values
        BuildOptionCharges = ErrorCount

End Function

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub AddError( _
    ByRef Findings() As String, _
    ByRef FindingCount As Long, _
    ByRef ErrorCount As Long, _
    ByVal Item As String, _
    ByVal Message As String)
'
'------------------------------------------------------------------------------
'                            RECORD UNUSABLE VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Record one unusable Params value: append the finding and count it
'
' WHY THIS EXISTS
'   Every validation in this component reports an unusable Params value the
'   same way, as a finding on the Params sheet, and adds it to the count of its
'   caller
'
' INPUTS
'   Findings
'     Findings list; appended to
'
'   FindingCount
'     Findings in use; incremented
'
'   ErrorCount
'     Running count of the caller; incremented
'
'   Item
'     Parameter code or table row
'
'   Message
'     What is wrong, without the INVALID prefix
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds 1 to ErrorCount, then appends a finding for the Params sheet whose
'   status is ST_INVALID followed by the message
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Util: AddFinding; CORE_Config: ST_INVALID
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COUNT AND REPORT
'------------------------------------------------------------------------------
    'Count the unusable value
        ErrorCount = ErrorCount + 1
    'Append the finding with the INVALID status prefix
        AddFinding Findings, FindingCount, PARAMS_SHEET, Item, "", ST_INVALID & Message

End Sub
