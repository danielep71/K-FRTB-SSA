Attribute VB_Name = "CORE_Parameters"
'==============================================================================
' MODULE: CORE_Parameters
'------------------------------------------------------------------------------
' PURPOSE
'   Turn the raw Params values (as read from the cells) into the validated
'   parameter vector and tables the risk-class procedures use, and report
'   every unusable value as a Checks finding. Also provide the MAR40
'   reference parameter set and ladder used by the SSA_ worksheet functions
'   and by the "constants changed" check.
'
' PUBLIC SURFACE
'   None outside this VBA project (Option Private Module).
'     BuildParameters, BuildRegime, TransitionalFactor, BuildLadder,
'     BuildSpecificRiskTable, BuildFxTable,
'     BuildOptionCharges, ChangedParameterCount, ReferenceParameters,
'     ReferenceLadder.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util.
'
' STATE OWNERSHIP
'   None. Every procedure depends only on its arguments.
'
' ERROR POLICY
'   Cell content never raises: an unusable value becomes a finding and is
'   counted in the procedure's return value; the caller stops the run when
'   any count is non-zero. ERR_ARGUMENT is raised for a raw array of the
'   wrong shape, which is a layout fault, not an input fault.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults. Never touches the Excel
'   object model.
'
' REFERENCE
'   BCBS MAR40.1-40.2 (scaling, RWA), 40.6 Table 1, 40.26 Table 4,
'   40.27-40.28 Table 5, 40.43-40.47, 40.60-40.61, 40.72-40.73.
'
' UPDATED
'   2026-11-15
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
        Private Const PARAMS_SHEET   As String = "Params"     'Sheet named in findings
        Private Const VALUE_TOLERANCE As Double = 0.000000001  'Equality of a constant to MAR40


'------------------------------------------------------------------------------
'                               PARAMETER VECTOR
'------------------------------------------------------------------------------

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
'==============================================================================
'                                BuildParameters
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the scalar parameters and return them typed.
'
' INPUTS
'   RawNumeric       Cell values (1 To PI_COUNT), in the order of
'                    PRM_NUMERIC_CODES.
'   RawAsOfDate      Cell value of AsOfDate; must be a date
'                    (CORE_Util.IsDateSerial).
'   RawReportingCcy  Cell value of ReportingCcy; must not be blank.
'   RawGoldCode      Cell value of GoldCode; must not be blank.
'   Findings         Findings list, appended to (CORE_Util.AddFinding).
'   FindingCount     Findings in use; incremented.
'
' RETURNS
'   Number of unusable values. Output parameters: Parameters(1 To PI_COUNT)
'   as decimals / years / factors; AsOfDate as a date serial; ReportingCcy
'   and GoldCode upper case.
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when RawNumeric is not (1 To PI_COUNT).
'
' REFERENCE
'   Domains: PRM_DOMAINS (P: > 0; N: >= 0; R: 0 <= rate <= 1).
'
' UPDATED
'   2026-11-15
'==============================================================================
'
    Dim ErrorCount  As Long      'Unusable values found
    Dim Code        As String    'Parameter code
    Dim Domain      As String    'Domain letter of the parameter
    Dim Value       As Double    'Numeric value
    Dim i           As Long      'Parameter index

        If LBound(RawNumeric) <> 1 Or UBound(RawNumeric) <> PI_COUNT Then
            Err.Raise ERR_ARGUMENT, "CORE_Parameters.BuildParameters", _
                "RawNumeric must be (1 To " & PI_COUNT & ")."
        End If
        ReDim Parameters(1 To PI_COUNT)

        'Scalars of text and date type.
        If IsDateSerial(RawAsOfDate) Then
            AsOfDate = NumberValue(RawAsOfDate)
        Else
            AddError Findings, FindingCount, ErrorCount, PRM_ASOF, "must be a date"
        End If
        ReportingCcy = CodeText(RawReportingCcy)
        If Len(ReportingCcy) = 0 Then
            AddError Findings, FindingCount, ErrorCount, PRM_REPCCY, "must not be blank"
        End If
        GoldCode = CodeText(RawGoldCode)
        If Len(GoldCode) = 0 Then
            AddError Findings, FindingCount, ErrorCount, PRM_GOLD, "must not be blank"
        End If

        'Numeric parameters, each against its domain.
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
        BuildParameters = ErrorCount

End Function


Public Function BuildRegime( _
    ByVal RawRegime As Variant, _
    ByVal RawEndDate As Variant, _
    ByRef Regime As String, _
    ByRef EndDate As Double, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'==============================================================================
'                                  BuildRegime
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the regime and the last date of the transitional multiplier.
'
' INPUTS
'   RawRegime    Cell value of Regime; Basel or CRR, without regard to case.
'   RawEndDate   Cell value of TransitionalEndDate; must be a date when the
'                regime is the CRR, and is not read otherwise.
'   Findings     Findings list, appended to (CORE_Util.AddFinding).
'   FindingCount Findings in use; incremented.
'
' RETURNS
'   Number of unusable values. Output parameters: Regime as REGIME_BASEL or
'   REGIME_CRR (REGIME_BASEL when it is unusable); EndDate as a date serial,
'   0 when it is not read or unusable.
'
' REFERENCE
'   Article 325(2) and Article 495s(2) of Regulation (EU) No 575/2013, as
'   amended by Regulation (EU) 2024/1623 and Delegated Regulation (EU)
'   2026/1221.
'
' UPDATED
'   2026-11-20
'==============================================================================
'
    Dim ErrorCount  As Long      'Unusable values found

        Regime = REGIME_BASEL
        EndDate = 0#
        Select Case CodeText(RawRegime)
            Case REGIME_BASEL
                Regime = REGIME_BASEL
            Case REGIME_CRR
                Regime = REGIME_CRR
            Case Else
                AddError Findings, FindingCount, ErrorCount, PRM_REGIME, "must be Basel or CRR"
        End Select
        If Regime = REGIME_CRR Then
            If IsDateSerial(RawEndDate) Then
                EndDate = NumberValue(RawEndDate)
            Else
                AddError Findings, FindingCount, ErrorCount, PRM_TRANSITION_END, "must be a date"
            End If
        End If
        BuildRegime = ErrorCount

End Function


Public Function TransitionalFactor( _
    ByVal Regime As String, _
    ByVal AsOfDate As Double, _
    ByVal EndDate As Double, _
    ByVal Multiplier As Double) _
    As Double
'
'==============================================================================
'                              TransitionalFactor
'------------------------------------------------------------------------------
' PURPOSE
'   The factor that the requirement of the simplified standardised approach
'   is multiplied by on a date: the transitional multiplier under the CRR
'   until its last date, and 1 under the Basel regime and after that date.
'
' INPUTS
'   Regime       REGIME_BASEL or REGIME_CRR, from BuildRegime.
'   AsOfDate     Reporting date, a date serial.
'   EndDate      Last date of the multiplier, a date serial.
'   Multiplier   The multiplier, a decimal; PI_TRANSITIONAL.
'
' RETURNS
'   Multiplier when Regime is REGIME_CRR and AsOfDate is on or before
'   EndDate; 1 otherwise.
'
' REFERENCE
'   Article 495s(2) of Regulation (EU) No 575/2013, from Delegated
'   Regulation (EU) 2026/1221.
'
' UPDATED
'   2026-11-20
'==============================================================================
'
        If Regime = REGIME_CRR And AsOfDate <= EndDate Then
            TransitionalFactor = Multiplier
        Else
            TransitionalFactor = 1#
        End If

End Function


Public Function ChangedParameterCount( _
    ByRef Parameters() As Double, _
    ByVal GoldCode As String) _
    As Long
'
'==============================================================================
'                             ChangedParameterCount
'------------------------------------------------------------------------------
' PURPOSE
'   Count the constants that differ from their MAR40 value (a what-if run).
'
' INPUTS
'   Parameters   Vector (1 To PI_COUNT) from BuildParameters.
'   GoldCode     Gold code in use, upper case.
'
' RETURNS
'   Number of numeric parameters differing from PRM_REFERENCE_VALUES by more
'   than 1E-9, plus 1 when GoldCode is not XAU.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Reference() As Double    'MAR40 values
    Dim Changed     As Long      'Count
    Dim i           As Long      'Parameter index

        ReferenceParameters Reference
        For i = 1 To PI_COUNT
            If Abs(Parameters(i) - Reference(i)) > VALUE_TOLERANCE Then Changed = Changed + 1
        Next i
        If GoldCode <> PRM_REFERENCE_GOLD Then Changed = Changed + 1
        ChangedParameterCount = Changed

End Function


Public Sub ReferenceParameters( _
    ByRef Parameters() As Double)
'
'==============================================================================
'                              ReferenceParameters
'------------------------------------------------------------------------------
' PURPOSE
'   Return the MAR40 reference parameter vector.
'
' RETURNS
'   Parameters(1 To PI_COUNT), from PRM_REFERENCE_VALUES.
'
' REFERENCE
'   MAR40.1-40.2, 40.27-40.28, 40.43, 40.47, 40.61, 40.72-40.73.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        ParseList PRM_REFERENCE_VALUES, PI_COUNT, 0#, Parameters

End Sub


'------------------------------------------------------------------------------
'                                MATURITY LADDER
'------------------------------------------------------------------------------

Public Function BuildLadder( _
    ByRef RawLadder() As Variant, _
    ByRef Ladder() As Double, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'==============================================================================
'                                  BuildLadder
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the maturity-ladder table of Params and return it as numbers.
'
' INPUTS
'   RawLadder     Cell values (1 To BAND_COUNT, 1 To LD_WIDTH): zone, upper
'                 bound for coupon >= threshold, upper bound for coupon below
'                 it (years, blank = open-ended or absent), risk weight.
'   Findings      Findings list, appended to.
'   FindingCount  Findings in use; incremented.
'
' RETURNS
'   Number of unusable values. Ladder(1 To BAND_COUNT, 1 To LD_WIDTH) as
'   numbers, OPEN_BOUND for a blank bound.
'
' ERROR POLICY
'   Raises ERR_ARGUMENT for a RawLadder of another shape.
'
' REFERENCE
'   MAR40.26 Table 4. Zones must be 1 to 3 and non-decreasing; bounds must
'   increase from band to band; weights must be 0 or more.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim ErrorCount  As Long      'Unusable values found
    Dim Previous    As Double    'Previous finite bound in the column
    Dim Value       As Variant   'Cell value
    Dim Item        As String    'Finding item, "Ladder band n"
    Dim b           As Long      'Band
    Dim c           As Long      'Bound column

        If UBound(RawLadder, 1) <> BAND_COUNT Or UBound(RawLadder, 2) <> LD_WIDTH Then
            Err.Raise ERR_ARGUMENT, "CORE_Parameters.BuildLadder", _
                "RawLadder must be (1 To " & BAND_COUNT & ", 1 To " & LD_WIDTH & ")."
        End If
        ReDim Ladder(1 To BAND_COUNT, 1 To LD_WIDTH)
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

        'Upper bounds: blank means open; finite bounds must increase.
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
        BuildLadder = ErrorCount

End Function


Public Sub ReferenceLadder( _
    ByRef Ladder() As Double)
'
'==============================================================================
'                                ReferenceLadder
'------------------------------------------------------------------------------
' PURPOSE
'   Return the MAR40 Table 4 ladder.
'
' RETURNS
'   Ladder(1 To BAND_COUNT, 1 To LD_WIDTH), as BuildLadder returns it.
'
' REFERENCE
'   MAR40.26 Table 4.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Column()    As Double    'One parsed list
    Dim b           As Long      'Band

        ReDim Ladder(1 To BAND_COUNT, 1 To LD_WIDTH)
        ParseList LADDER_ZONES, BAND_COUNT, 0#, Column
        For b = 1 To BAND_COUNT
            Ladder(b, LD_ZONE) = Column(b)
        Next b
        ParseList LADDER_UPPER_HIGH, BAND_COUNT, OPEN_BOUND, Column
        For b = 1 To BAND_COUNT
            Ladder(b, LD_UPPER_HIGH) = Column(b)
        Next b
        ParseList LADDER_UPPER_LOW, BAND_COUNT, OPEN_BOUND, Column
        For b = 1 To BAND_COUNT
            Ladder(b, LD_UPPER_LOW) = Column(b)
        Next b
        ParseList LADDER_WEIGHTS, BAND_COUNT, 0#, Column
        For b = 1 To BAND_COUNT
            Ladder(b, LD_WEIGHT) = Column(b)
        Next b
        'One month is 1/12 of a year exactly; a decimal literal falls short of it.
        Ladder(1, LD_UPPER_HIGH) = 1# / 12#
        Ladder(1, LD_UPPER_LOW) = 1# / 12#

End Sub


'------------------------------------------------------------------------------
'                                 LOOKUP TABLES
'------------------------------------------------------------------------------

Public Function BuildSpecificRiskTable( _
    ByRef RawTable() As Variant, _
    ByRef Keys() As String, _
    ByRef Charges() As Double, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'==============================================================================
'                            BuildSpecificRiskTable
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the specific-risk table (key "CATEGORY|grade|maturity band",
'   charge) and return it typed.
'
' INPUTS
'   RawTable      Cell values (1 To n, 1 To 2): key, charge.
'   Findings      Findings list, appended to.
'   FindingCount  Findings in use; incremented.
'
' RETURNS
'   Number of unusable values. Keys(1 To n) upper case ("" for a blank
'   row) and Charges(1 To n) as decimals.
'
' REFERENCE
'   MAR40.6 Table 1. Charges must be decimals from 0 to 1; keys unique.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim ErrorCount  As Long      'Unusable values found
    Dim RowCount    As Long      'InputRows of the table
    Dim r           As Long      'Table row
    Dim j           As Long      'Earlier row

        RowCount = UBound(RawTable, 1)
        ReDim Keys(1 To RowCount)
        ReDim Charges(1 To RowCount)
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
'==============================================================================
'                                 BuildFxTable
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the FX table (currency, units of reporting currency per unit)
'   and return it typed.
'
' INPUTS
'   RawTable      Cell values (1 To n, 1 To 2): currency code, rate.
'   ReportingCcy  Reporting currency, upper case.
'   Findings      Findings list, appended to.
'   FindingCount  Findings in use; incremented.
'
' RETURNS
'   Total unusable values. Codes(1 To n) upper case ("" for a blank row),
'   Rates(1 To n); ReportingErrors (0 or 1: reporting currency missing or
'   not at rate 1, rule C05), DuplicateErrors (codes listed again, C06),
'   RateErrors (rates missing, not numbers or not > 0, C07).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim RowCount    As Long      'InputRows of the table
    Dim Position    As Long      'Row of the reporting currency
    Dim r           As Long      'Table row
    Dim j           As Long      'Earlier row

        RowCount = UBound(RawTable, 1)
        ReDim Codes(1 To RowCount)
        ReDim Rates(1 To RowCount)
        ReportingErrors = 0
        DuplicateErrors = 0
        RateErrors = 0
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
        Position = FindCode(Codes, ReportingCcy)
        If Position < 1 Then
            ReportingErrors = 1
        ElseIf Rates(Position) <> 1 Then
            ReportingErrors = 1
        End If
        If ReportingErrors > 0 Then
            AddFinding Findings, FindingCount, PARAMS_SHEET, ReportingCcy, "", _
                ST_INVALID & "reporting currency must be listed in FX_Table with rate 1"
        End If
        BuildFxTable = ReportingErrors + DuplicateErrors + RateErrors

End Function


Public Function BuildOptionCharges( _
    ByRef RawCharges() As Variant, _
    ByRef Charges() As Double, _
    ByRef Findings() As String, _
    ByRef FindingCount As Long) _
    As Long
'
'==============================================================================
'                              BuildOptionCharges
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the non-delta option charges entered per risk class.
'
' INPUTS
'   RawCharges    Cell values (1 To RC_COUNT), in RC_ order; blank = 0.
'   Findings      Findings list, appended to.
'   FindingCount  Findings in use; incremented.
'
' RETURNS
'   Number of unusable values. Charges(1 To RC_COUNT) in reporting currency.
'
' REFERENCE
'   MAR40.2 (option charges added before scaling), MAR40.74-40.84.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim ErrorCount  As Long      'Unusable values found
    Dim Item        As String    'Finding item, workbook name
    Dim i           As Long      'Risk class

        ReDim Charges(1 To RC_COUNT)
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
        BuildOptionCharges = ErrorCount

End Function


'------------------------------------------------------------------------------
'                                    HELPERS
'------------------------------------------------------------------------------

Private Sub AddError( _
    ByRef Findings() As String, _
    ByRef FindingCount As Long, _
    ByRef ErrorCount As Long, _
    ByVal Item As String, _
    ByVal Message As String)
'
'==============================================================================
'                                   AddError
'------------------------------------------------------------------------------
' PURPOSE
'   Record one unusable Params value: append the finding and count it.
'
' INPUTS
'   Findings, FindingCount   Findings list and its count; appended to.
'   ErrorCount               Running count of the caller; incremented.
'   Item                     Parameter code or table row.
'   Message                  What is wrong, without the INVALID prefix.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        ErrorCount = ErrorCount + 1
        AddFinding Findings, FindingCount, PARAMS_SHEET, Item, "", ST_INVALID & Message

End Sub
