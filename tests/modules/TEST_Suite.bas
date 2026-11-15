Attribute VB_Name = "TEST_Suite"
'==============================================================================
' MODULE: TEST_Suite
'------------------------------------------------------------------------------
' PURPOSE
'   Unit tests of the K-FRTB-SSA core and facade on hand-checked cases, and
'   tests of the whole workbook on the shipped sample data: a regression of
'   its results, runs that fail and inputs changed after a run.
'
' PUBLIC SURFACE
'   Test_RunAll (core, facade and workbook), Test_RunCore (no workbook
'   access). Never part of the production import set.
'
' DEPENDENCIES
'   CORE_Config, CORE_Util, CORE_Parameters, CORE_InterestRate,
'   CORE_Equity, CORE_Commodity, CORE_ForeignExchange, FRTBSSA_Formulas, and
'   M_Main for the workbook test.
'
' STATE OWNERSHIP
'   mPassed, mFailed and mReport accumulate one run's results. The workbook
'   tests run RunSSA_Silent, which rewrites the output sheets; two of them
'   change a cell or a workbook name for one operation and put it back.
'
' ERROR POLICY
'   An unexpected error inside a test is recorded as a failure of that
'   test; the run continues.
'
' COMPATIBILITY
'   Excel VBA; also runs under LibreOffice's VBA mode through the harness
'   in tools/.
'
' REFERENCE
'   Expected values: hand calculation (stated in each test), MAR40.61
'   Table 9, and tools/prototype/ssa_reference.py for the sample data. The
'   fingerprints of CheckFingerprint were calculated apart from the VBA
'   code, from the rule its header states.
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
    'Require explicit declarations; keep the tests inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const TOLERANCE          As Double = 0.000001    'Absolute tolerance
        Private Const SAMPLE_REQUIREMENT As Double = 48267441#   'Sample data, ssa_reference.py
        Private Const SAMPLE_RWA         As Double = 603343012.5 'Sample data, ssa_reference.py
        Private Const SAMPLE_POSITIONS   As Long = 35            'Sample rows with status OK

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mPassed      As Long      'Assertions passed
        Private mFailed      As Long      'Assertions failed
        Private mReport      As String    'Failure messages


'------------------------------------------------------------------------------
'                                    RUNNERS
'------------------------------------------------------------------------------

Public Function Test_RunAll() As String
'
'==============================================================================
'                                  Test_RunAll
'------------------------------------------------------------------------------
' PURPOSE
'   Run every test, including the workbook regression.
'
' RETURNS
'   "PASS=n; FAIL=m" followed by one line per failure.
'
' STATE OWNERSHIP
'   Rewrites the output sheets through RunSSA_Silent.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        ResetResults
        RunCoreTests
        RunTest "FailedRun"
        RunTest "OutdatedResults"
        RunTest "Workbook"
        Test_RunAll = ResultLine()
        Debug.Print Test_RunAll

End Function


Public Function Test_RunCore() As String
'
'==============================================================================
'                                 Test_RunCore
'------------------------------------------------------------------------------
' PURPOSE
'   Run the core and facade tests only; the workbook is not touched.
'
' RETURNS
'   "PASS=n; FAIL=m" followed by one line per failure.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        ResetResults
        RunCoreTests
        Test_RunCore = ResultLine()
        Debug.Print Test_RunCore

End Function


Private Sub RunCoreTests()
'
'==============================================================================
'                                 RunCoreTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run each core and facade test.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        RunTest "TimeBand"
        RunTest "LadderCharges"
        RunTest "ForeignExchange"
        RunTest "Equity"
        RunTest "Commodity"
        RunTest "InterestRateRows"
        RunTest "IssueNetting"
        RunTest "DateBounds"
        RunTest "Parameters"
        RunTest "Facade"
        RunTest "Fingerprint"

End Sub


Private Sub RunTest( _
    ByVal TestName As String)
'
'==============================================================================
'                                    RunTest
'------------------------------------------------------------------------------
' PURPOSE
'   Run one test by name, recording an unexpected error as a failure.
'
' INPUTS
'   TestName   Name of the test procedure without the "Check" prefix.
'
' ERROR POLICY
'   Contains any error raised by the test and records it.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        On Error GoTo Failed
        Select Case TestName
            Case "TimeBand": CheckTimeBand
            Case "LadderCharges": CheckLadderCharges
            Case "ForeignExchange": CheckForeignExchange
            Case "Equity": CheckEquity
            Case "Commodity": CheckCommodity
            Case "InterestRateRows": CheckInterestRateRows
            Case "IssueNetting": CheckIssueNetting
            Case "DateBounds": CheckDateBounds
            Case "Parameters": CheckParameters
            Case "Facade": CheckFacade
            Case "Fingerprint": CheckFingerprint
            Case "FailedRun": CheckFailedRun
            Case "OutdatedResults": CheckOutdatedResults
            Case "Workbook": CheckWorkbook
        End Select
        Exit Sub

Failed:
        Record False, TestName & ": unexpected error " & Err.Number & " " & Err.Description

End Sub


'------------------------------------------------------------------------------
'                                     TESTS
'------------------------------------------------------------------------------

Private Sub CheckTimeBand()
'
'==============================================================================
'                                 CheckTimeBand
'------------------------------------------------------------------------------
' PURPOSE
'   Band boundaries of MAR40.26 Table 4 (upper bounds inclusive).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Ladder()    As Double    'MAR40 ladder

        ReferenceLadder Ladder
        AssertNear "TimeBand 0y", TimeBand(0#, True, Ladder), 1
        AssertNear "TimeBand 1 month", TimeBand(1# / 12#, True, Ladder), 1
        AssertNear "TimeBand 1 month + 1 day", TimeBand(1# / 12# + 1# / 365#, True, Ladder), 2
        AssertNear "TimeBand 1y high", TimeBand(1#, True, Ladder), 4
        AssertNear "TimeBand 1.5y high", TimeBand(1.5, True, Ladder), 5
        AssertNear "TimeBand 20y high", TimeBand(20#, True, Ladder), 12
        AssertNear "TimeBand 25y high", TimeBand(25#, True, Ladder), 13
        AssertNear "TimeBand 1.95y low", TimeBand(1.95, False, Ladder), 6
        AssertNear "TimeBand 12y low", TimeBand(12#, False, Ladder), 13
        AssertNear "TimeBand 25y low", TimeBand(25#, False, Ladder), 15

End Sub


Private Sub CheckLadderCharges()
'
'==============================================================================
'                              CheckLadderCharges
'------------------------------------------------------------------------------
' PURPOSE
'   General market risk of hand-calculated ladders.
'
' REFERENCE
'   Case A: band 4 long 100 / short 60, band 6 short 50, band 10 long 80.
'   Vertical 10% x 60 = 6; zones +40, -50, +80; offset 1-2 = 40, 2-3 = 10,
'   1-3 = 0; between 40% x 50 = 20; net |70| = 70; total 96.
'   Case B: band 1 long 10, band 2 short 4: within zone 1, 40% x 4 = 1.6;
'   net 6; total 7.6.
'   Case C: band 3 long 30, band 12 short 50: offset 1-3 = 30 at 100%;
'   net 20; total 50.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Ladder()     As Double    'MAR40 ladder
    Dim Parameters() As Double    'MAR40 parameters
    Dim Longs()      As Double    'Weighted longs
    Dim Shorts()     As Double    'Weighted shorts
    Dim Components() As Double    'Result

        ReferenceLadder Ladder
        ReferenceParameters Parameters
        ReDim Longs(1 To BAND_COUNT)
        ReDim Shorts(1 To BAND_COUNT)
        Longs(4) = 100: Shorts(4) = 60: Shorts(6) = 50: Longs(10) = 80
        LadderCharges Longs, Shorts, Ladder, Parameters, Components
        AssertNear "Ladder A vertical", Components(ZC_VERTICAL), 6
        AssertNear "Ladder A between", Components(ZC_BETWEEN), 20
        AssertNear "Ladder A net", Components(ZC_NET), 70
        AssertNear "Ladder A total", Components(ZC_GENERAL), 96

        ReDim Longs(1 To BAND_COUNT)
        ReDim Shorts(1 To BAND_COUNT)
        Longs(1) = 10: Shorts(2) = 4
        LadderCharges Longs, Shorts, Ladder, Parameters, Components
        AssertNear "Ladder B zone 1 charge", Components(ZC_ZONE_FIRST + 3), 1.6
        AssertNear "Ladder B total", Components(ZC_GENERAL), 7.6

        ReDim Longs(1 To BAND_COUNT)
        ReDim Shorts(1 To BAND_COUNT)
        Longs(3) = 30: Shorts(12) = 50
        LadderCharges Longs, Shorts, Ladder, Parameters, Components
        AssertNear "Ladder C offset 1-3", Components(ZC_MATCHED13), 30
        AssertNear "Ladder C total", Components(ZC_GENERAL), 50

End Sub


Private Sub CheckForeignExchange()
'
'==============================================================================
'                             CheckForeignExchange
'------------------------------------------------------------------------------
' PURPOSE
'   MAR40.61 Table 9 at rate 1, plus an excluded and a duplicate row.
'
' REFERENCE
'   Table 9: longs 50 + 100 + 150 = 300, shorts 20 + 180 = 200, gold 35;
'   charge 8% x (300 + 35) = 26.8.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim InputRows()        As Variant   'FXPositions cells
    Dim FxCodes()     As String    'FX table codes
    Dim FxRates()     As Double    'FX table rates
    Dim Parameters()  As Double    'MAR40 parameters
    Dim RowOutput()   As Variant   'FXCalc
    Dim Currencies()  As Variant   'Buckets FX block
    Dim Totals()      As Double    'Totals

        ReferenceParameters Parameters
        FxTable "EUR|CAD|SEK|GBP|NOK|USD|XAU", FxCodes, FxRates
        InputRows = RowsFrom(FX_INPUT_WIDTH, Array( _
            Array("F1", "", "CAD", "Long", 50), _
            Array("F2", "", "SEK", "Long", 100), _
            Array("F3", "", "gbp", "Long", 150), _
            Array("F4", "", "NOK", "Short", 20), _
            Array("F5", "", "USD", "Short", 180), _
            Array("F6", "", "XAU", "Short", 35), _
            Array("F7", "", "EUR", "Long", 999), _
            Array("F8", "", "USD", "Long", 1), _
            Array("f8", "", "USD", "Long", 1)))
        CalculateForeignExchange InputRows, Parameters, FxCodes, FxRates, "EUR", "XAU", _
                                 RowOutput, Currencies, Totals
        AssertNear "FX sum of longs", Totals(TOT_LONG), 300
        AssertNear "FX sum of shorts", Totals(TOT_SHORT), 200
        AssertNear "FX gold", Totals(TOT_GOLD), 35
        AssertNear "FX overall net open position", Totals(TOT_ONOP), 335
        AssertNear "FX charge (Table 9)", Totals(TOT_GENERAL), 26.8
        AssertNear "FX rows used", Totals(TOT_USED), 6
        AssertNear "FX rows excluded", Totals(TOT_EXCLUDED), 1
        AssertNear "FX rows rejected", Totals(TOT_REJECTED), 2
        AssertText "FX reporting-currency status", CStr(RowOutput(7, 2)), _
                   ST_EXCLUDED & "reporting currency carries no FX risk"
        AssertText "FX duplicate status", CStr(RowOutput(9, 2)), ST_INVALID & "duplicate PositionID"

End Sub


Private Sub CheckEquity()
'
'==============================================================================
'                                  CheckEquity
'------------------------------------------------------------------------------
' PURPOSE
'   Netting of the same stock (case-insensitive) and index treatment.
'
' REFERENCE
'   Stock X: +100 - 40 = 60, specific 8% x 60 = 4.8; index +50, 2% = 1;
'   general 8% x |60 + 50| = 8.8 (MAR40.43, 40.46, 40.47).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim InputRows()        As Variant   'EQPositions cells
    Dim FxCodes()     As String    'FX table codes
    Dim FxRates()     As Double    'FX table rates
    Dim Parameters()  As Double    'MAR40 parameters
    Dim RowOutput()   As Variant   'EQCalc
    Dim Markets()     As Variant   'Buckets equity block
    Dim Totals()      As Double    'Totals

        ReferenceParameters Parameters
        FxTable "EUR|XAU", FxCodes, FxRates
        InputRows = RowsFrom(EQ_INPUT_WIDTH, Array( _
            Array("E1", "", "IT", "Single name", "X", "EUR", "Long", 100), _
            Array("E2", "", "it", "single name", "x", "EUR", "Short", 40), _
            Array("E3", "", "IT", "Index", "IDX", "EUR", "Long", 50), _
            Array("E4", "", "IT", "Basket", "B", "EUR", "Long", 50), _
            Array("E5", "", "IT", "Single name", "Y", "XAU", "Long", 50)))
        CalculateEquity InputRows, Parameters, FxCodes, FxRates, "XAU", RowOutput, Markets, Totals
        AssertNear "EQ specific", Totals(TOT_SPECIFIC), 4.8
        AssertNear "EQ index", Totals(TOT_OTHER), 1
        AssertNear "EQ general", Totals(TOT_GENERAL), 8.8
        AssertNear "EQ markets", Totals(TOT_BUCKETS), 1
        AssertNear "EQ rejected", Totals(TOT_REJECTED), 2
        AssertText "EQ type status", CStr(RowOutput(4, 2)), ST_INVALID & "Type must be Single name or Index"
        AssertText "EQ gold status", CStr(RowOutput(5, 2)), ST_INVALID & "currency not in FX_Table"

End Sub


Private Sub CheckCommodity()
'
'==============================================================================
'                                CheckCommodity
'------------------------------------------------------------------------------
' PURPOSE
'   Simplified approach on one commodity entered in two cases.
'
' REFERENCE
'   Oil +100, -30 at spot: 15% x 70 = 10.5; 3% x 130 = 3.9 (MAR40.72-40.73).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim InputRows()        As Variant   'COPositions cells
    Dim FxCodes()     As String    'FX table codes
    Dim FxRates()     As Double    'FX table rates
    Dim Parameters()  As Double    'MAR40 parameters
    Dim RowOutput()   As Variant   'COCalc
    Dim Commodities() As Variant   'Buckets commodity block
    Dim Totals()      As Double    'Totals

        ReferenceParameters Parameters
        FxTable "EUR|XAU", FxCodes, FxRates
        InputRows = RowsFrom(CO_INPUT_WIDTH, Array( _
            Array("C1", "", "Oil", "Long", 10, "bbl", 10, "EUR"), _
            Array("C2", "", "OIL ", "Short", 3, "bbl", 10, "EUR"), _
            Array("C3", "", "Oil", "Long", "ten", "bbl", 10, "EUR")))
        CalculateCommodity InputRows, Parameters, FxCodes, FxRates, "XAU", RowOutput, Commodities, Totals
        AssertNear "CO directional", Totals(TOT_GENERAL), 10.5
        AssertNear "CO gross", Totals(TOT_OTHER), 3.9
        AssertNear "CO commodities", Totals(TOT_BUCKETS), 1
        AssertText "CO quantity status", CStr(RowOutput(3, 2)), ST_INVALID & "Quantity must be a number"

End Sub


Private Sub CheckInterestRateRows()
'
'==============================================================================
'                             CheckInterestRateRows
'------------------------------------------------------------------------------
' PURPOSE
'   Row validation, slotting and specific risk of interest rate rows.
'
' REFERENCE
'   Row 1: QUAL A+ to BBB-, 5 years 1 day to maturity, 4% coupon: band 9
'   (5-7 years), weight 3.25%, weighted 32,500; specific 1.6% x 1,000,000 =
'   16,000 (MAR40.6, 40.26).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim InputRows()        As Variant   'IRPositions cells
    Dim FxCodes()     As String    'FX table codes
    Dim FxRates()     As Double    'FX table rates
    Dim Parameters()  As Double    'MAR40 parameters
    Dim Ladder()      As Double    'MAR40 ladder
    Dim Labels()      As String    'Ladder labels
    Dim SrKeys()      As String    'Specific-risk keys
    Dim SrCharges()   As Double    'Specific-risk charges
    Dim RowOutput()   As Variant   'IRCalc
    Dim LadderOutput() As Variant  'IR_Ladder
    Dim ZoneOutput()  As Variant   'IR_Zones
    Dim Totals()      As Double    'Totals
    Dim AsOfDate      As Double    'Reporting date

        ReferenceParameters Parameters
        ReferenceLadder Ladder
        ReDim Labels(1 To BAND_COUNT, 1 To 2)
        FxTable "EUR|XAU", FxCodes, FxRates
        ReDim SrKeys(1 To 2)
        ReDim SrCharges(1 To 2)
        SrKeys(1) = "QUAL|A+ TO BBB-|>24M": SrCharges(1) = 0.016
        SrKeys(2) = "GOV|AAA TO AA-|ANY": SrCharges(2) = 0
        AsOfDate = CDbl(DateSerial(2026, 9, 30))
        InputRows = RowsFrom(IR_INPUT_WIDTH, Array( _
            Array("R1", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.04, DateSerial(2031, 10, 1), _
                  Empty, "QUAL", "A+ to BBB-", "ISS1"), _
            Array("R2", "", "Bond", "EUR", "Long", 1000, "Fixed", Empty, DateSerial(2031, 10, 1), _
                  Empty, "GOV", "AAA to AA-", ""), _
            Array("R3", "", "FRN", "EUR", "Long", 1000, "Floating", Empty, DateSerial(2028, 1, 1), _
                  DateSerial(2029, 1, 1), "GOV", "AAA to AA-", ""), _
            Array("R4", "", "Bond", "EUR", "Long", 1000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "OTHER", "A+ to BBB-", ""), _
            Array("R5", "", "Bond", "XAU", "Long", 1000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "GOV", "AAA to AA-", "")))
        CalculateInterestRate InputRows, AsOfDate, Parameters, FxCodes, FxRates, "XAU", SrKeys, _
                              SrCharges, Ladder, Labels, RowOutput, LadderOutput, ZoneOutput, Totals
        AssertText "IR row 1 status", CStr(RowOutput(1, 2)), ST_OK
        AssertNear "IR row 1 band", RowOutput(1, 9), 9
        AssertNear "IR row 1 weighted", RowOutput(1, 12), 32500
        AssertNear "IR row 1 specific", RowOutput(1, 20), 16000
        AssertText "IR row 2 status", CStr(RowOutput(2, 2)), _
                   ST_INVALID & "Coupon required for a fixed-rate position"
        AssertText "IR row 3 status", CStr(RowOutput(3, 2)), _
                   ST_INVALID & "repricing date outside AsOfDate to final maturity"
        AssertText "IR row 4 status", CStr(RowOutput(4, 2)), _
                   ST_INVALID & "no MAR40.6 charge for this category and rating grade"
        AssertText "IR row 5 status", CStr(RowOutput(5, 2)), _
                   ST_INVALID & "gold is FX risk; enter it on FXPositions"
        AssertNear "IR specific total", Totals(TOT_SPECIFIC), 16000
        AssertNear "IR general total (net 32,500)", Totals(TOT_GENERAL), 32500
        AssertNear "IR rejected", Totals(TOT_REJECTED), 4

End Sub


Private Sub CheckIssueNetting()
'
'==============================================================================
'                               CheckIssueNetting
'------------------------------------------------------------------------------
' PURPOSE
'   Specific-risk netting: rows net only through an entered Issue ID, and a
'   row without one never nets, even when its PositionID equals the Issue ID
'   of another row.
'
' REFERENCE
'   Rows 1 and 2: PositionID R1 without an Issue ID, and Issue ID R1; long
'   and short 1,000,000 at 1.6%: 16,000 each, not a net of zero. Rows 3 to
'   5: Issue ID ISS9 in two cases, net 600,000 - 200,000 + 1,000 = 401,000,
'   charged on the first row at its 1.6% = 6,416; row 5 has another SR key,
'   so the three rows are flagged MIXED (MAR40.6).
'
' UPDATED
'   2026-11-15
'==============================================================================
'
    Dim InputRows()   As Variant   'IRPositions cells
    Dim FxCodes()     As String    'FX table codes
    Dim FxRates()     As Double    'FX table rates
    Dim Parameters()  As Double    'MAR40 parameters
    Dim Ladder()      As Double    'MAR40 ladder
    Dim Labels()      As String    'Ladder labels
    Dim SrKeys()      As String    'Specific-risk keys
    Dim SrCharges()   As Double    'Specific-risk charges
    Dim RowOutput()   As Variant   'IRCalc
    Dim LadderOutput() As Variant  'IR_Ladder
    Dim ZoneOutput()  As Variant   'IR_Zones
    Dim Totals()      As Double    'Totals
    Dim AsOfDate      As Double    'Reporting date

        ReferenceParameters Parameters
        ReferenceLadder Ladder
        ReDim Labels(1 To BAND_COUNT, 1 To 2)
        FxTable "EUR|XAU", FxCodes, FxRates
        ReDim SrKeys(1 To 2)
        ReDim SrCharges(1 To 2)
        SrKeys(1) = "QUAL|A+ TO BBB-|>24M": SrCharges(1) = 0.016
        SrKeys(2) = "GOV|AAA TO AA-|ANY": SrCharges(2) = 0
        AsOfDate = CDbl(DateSerial(2026, 9, 30))
        InputRows = RowsFrom(IR_INPUT_WIDTH, Array( _
            Array("R1", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.04, DateSerial(2031, 10, 1), _
                  Empty, "QUAL", "A+ to BBB-", ""), _
            Array("R2", "", "Bond", "EUR", "Short", 1000000, "Fixed", 0.04, DateSerial(2031, 10, 1), _
                  Empty, "QUAL", "A+ to BBB-", "R1"), _
            Array("R3", "", "Bond", "EUR", "Long", 600000, "Fixed", 0.04, DateSerial(2031, 10, 1), _
                  Empty, "QUAL", "A+ to BBB-", "iss9"), _
            Array("R4", "", "Bond", "EUR", "Short", 200000, "Fixed", 0.04, DateSerial(2031, 10, 1), _
                  Empty, "QUAL", "A+ to BBB-", "ISS9"), _
            Array("R5", "", "Bond", "EUR", "Long", 1000, "Fixed", 0.04, DateSerial(2031, 10, 1), _
                  Empty, "GOV", "AAA to AA-", "Iss9")))
        CalculateInterestRate InputRows, AsOfDate, Parameters, FxCodes, FxRates, "XAU", SrKeys, _
                              SrCharges, Ladder, Labels, RowOutput, LadderOutput, ZoneOutput, Totals
        AssertText "Issue key without an Issue ID", CStr(RowOutput(1, 17)), "POS:R1"
        AssertText "Issue key of an entered Issue ID", CStr(RowOutput(2, 17)), "R1"
        AssertNear "Issue net, row without an Issue ID", RowOutput(1, 18), 1000000
        AssertNear "Issue net, Issue ID equal to that PositionID", RowOutput(2, 18), -1000000
        AssertNear "Specific risk, row without an Issue ID", RowOutput(1, 20), 16000
        AssertNear "Specific risk, Issue ID equal to that PositionID", RowOutput(2, 20), 16000
        AssertText "Issue key ignores case", CStr(RowOutput(3, 17)), "ISS9"
        AssertNear "Issue net, same Issue ID in two cases", RowOutput(3, 18), 401000
        AssertNear "Specific risk on the first row of the issue", RowOutput(3, 20), 6416
        AssertNear "No specific risk on a later row of the issue", RowOutput(4, 20), 0
        AssertNear "Rows of an issue that mixes SR keys", Totals(TOT_MIXED), 3
        AssertNear "IR specific total, issue netting", Totals(TOT_SPECIFIC), 38416
        AssertNear "IR rows used, issue netting", Totals(TOT_USED), 5

End Sub


Private Sub CheckDateBounds()
'
'==============================================================================
'                                CheckDateBounds
'------------------------------------------------------------------------------
' PURPOSE
'   A number that is not the serial of a date is not a date: such a final
'   maturity or next repricing date rejects its row and the run goes on,
'   and such a reporting date is an unusable parameter.
'
' REFERENCE
'   Row 1: maturity typed as the number 20311231, above the serial of 31
'   December 9999 (2,958,465). Row 2: floating, repricing date typed as
'   20290101. Row 3: maturity on 31 December 9999, the last date. Row 4:
'   an ordinary row.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
    Dim InputRows()   As Variant   'IRPositions cells
    Dim FxCodes()     As String    'FX table codes
    Dim FxRates()     As Double    'FX table rates
    Dim Parameters()  As Double    'MAR40 parameters
    Dim Ladder()      As Double    'MAR40 ladder
    Dim Labels()      As String    'Ladder labels
    Dim SrKeys()      As String    'Specific-risk keys
    Dim SrCharges()   As Double    'Specific-risk charges
    Dim RowOutput()   As Variant   'IRCalc
    Dim LadderOutput() As Variant  'IR_Ladder
    Dim ZoneOutput()  As Variant   'IR_Zones
    Dim Totals()      As Double    'Totals
    Dim AsOfDate      As Double    'Reporting date
    Dim Raw()         As Variant   'Raw numeric parameters
    Dim Validated()   As Double    'Validated parameters
    Dim Findings()    As String    'Findings
    Dim FindingCount  As Long      'Findings in use
    Dim ReportingCcy  As String    'Reporting currency
    Dim GoldCode      As String    'Gold code
    Dim i             As Long      'Parameter

        ReferenceParameters Parameters
        ReferenceLadder Ladder
        ReDim Labels(1 To BAND_COUNT, 1 To 2)
        FxTable "EUR|XAU", FxCodes, FxRates
        ReDim SrKeys(1 To 1)
        ReDim SrCharges(1 To 1)
        SrKeys(1) = "GOV|AAA TO AA-|ANY": SrCharges(1) = 0
        AsOfDate = CDbl(DateSerial(2026, 9, 30))
        InputRows = RowsFrom(IR_INPUT_WIDTH, Array( _
            Array("R1", "", "Bond", "EUR", "Long", 1000, "Fixed", 0.04, 20311231, _
                  Empty, "GOV", "AAA to AA-", ""), _
            Array("R2", "", "FRN", "EUR", "Long", 1000, "Floating", Empty, DateSerial(2031, 10, 1), _
                  20290101, "GOV", "AAA to AA-", ""), _
            Array("R3", "", "Bond", "EUR", "Long", 1000, "Fixed", 0.04, DATE_SERIAL_MAX, _
                  Empty, "GOV", "AAA to AA-", ""), _
            Array("R4", "", "Bond", "EUR", "Long", 1000, "Fixed", 0.04, DateSerial(2031, 10, 1), _
                  Empty, "GOV", "AAA to AA-", "")))
        CalculateInterestRate InputRows, AsOfDate, Parameters, FxCodes, FxRates, "XAU", SrKeys, _
                              SrCharges, Ladder, Labels, RowOutput, LadderOutput, ZoneOutput, Totals
        AssertText "Maturity typed as a number is not a date", CStr(RowOutput(1, 2)), _
                   ST_INVALID & "Final maturity date is not a date"
        AssertText "Repricing date typed as a number is not a date", CStr(RowOutput(2, 2)), _
                   ST_INVALID & "Next repricing date is not a date"
        AssertText "Maturity on the last date is a date", CStr(RowOutput(3, 2)), ST_OK
        AssertText "A row after a rejected date is used", CStr(RowOutput(4, 2)), ST_OK
        AssertNear "IR rows used, date bounds", Totals(TOT_USED), 2
        AssertNear "IR rows rejected, date bounds", Totals(TOT_REJECTED), 2

        ReDim Raw(1 To PI_COUNT)
        For i = 1 To PI_COUNT
            Raw(i) = Parameters(i)
        Next i
        AssertNear "Reporting date typed as a number is unusable", _
            BuildParameters(Raw, 20260930, "EUR", "XAU", Validated, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount), 1
        AssertNear "Reporting date of zero is unusable", _
            BuildParameters(Raw, 0, "EUR", "XAU", Validated, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount), 1
        AssertNear "Reporting date on the last date is usable", _
            BuildParameters(Raw, DATE_SERIAL_MAX, "EUR", "XAU", Validated, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount), 0

End Sub


Private Sub CheckParameters()
'
'==============================================================================
'                                CheckParameters
'------------------------------------------------------------------------------
' PURPOSE
'   Unusable parameters are counted; a changed constant is detected; the
'   FX table must list the reporting currency at rate 1.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Reference()   As Double    'MAR40 parameters
    Dim Raw()         As Variant   'Raw numeric parameters
    Dim RawFx()       As Variant   'Raw FX table
    Dim Parameters()  As Double    'Validated parameters
    Dim FxCodes()     As String    'FX codes
    Dim FxRates()     As Double    'FX rates
    Dim Findings()    As String    'Findings
    Dim FindingCount  As Long      'Findings in use
    Dim AsOfDate      As Double    'Reporting date
    Dim ReportingCcy  As String    'Reporting currency
    Dim GoldCode      As String    'Gold code
    Dim ReportingErrors As Long    'C05
    Dim DuplicateErrors As Long    'C06
    Dim RateErrors    As Long      'C07
    Dim i             As Long      'Parameter

        ReferenceParameters Reference
        ReDim Raw(1 To PI_COUNT)
        For i = 1 To PI_COUNT
            Raw(i) = Reference(i)
        Next i
        AssertNear "Parameters: reference set is valid", _
            BuildParameters(Raw, DateSerial(2026, 9, 30), "EUR", "XAU", Parameters, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount), 0
        AssertNear "Parameters: reference set unchanged", ChangedParameterCount(Parameters, GoldCode), 0
        Raw(PI_DAYS_PER_YEAR) = 360
        Raw(PI_VERTICAL) = "ten"
        Raw(PI_ZONE1) = 1.5
        AssertNear "Parameters: two unusable values", _
            BuildParameters(Raw, DateSerial(2026, 9, 30), "EUR", "XAU", Parameters, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount), 2

        ReDim RawFx(1 To 3, 1 To 2)
        RawFx(1, 1) = "USD": RawFx(1, 2) = 0.9
        RawFx(2, 1) = "usd": RawFx(2, 2) = 0.9
        RawFx(3, 1) = "EUR": RawFx(3, 2) = 1.1
        BuildFxTable RawFx, "EUR", FxCodes, FxRates, ReportingErrors, DuplicateErrors, RateErrors, _
                     Findings, FindingCount
        AssertNear "FX table: reporting currency not at 1", ReportingErrors, 1
        AssertNear "FX table: duplicate code", DuplicateErrors, 1
        AssertNear "FX table: no rate error", RateErrors, 0

End Sub


Private Sub CheckFacade()
'
'==============================================================================
'                                  CheckFacade
'------------------------------------------------------------------------------
' PURPOSE
'   The worksheet functions return the core results and worksheet errors
'   outside their domain.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Longs(1 To 15)  As Variant   'Weighted longs, ladder case A
    Dim Shorts(1 To 15) As Variant   'Weighted shorts, ladder case A
    Dim b               As Long      'Band

        For b = 1 To BAND_COUNT
            Longs(b) = 0
            Shorts(b) = 0
        Next b
        Longs(4) = 100: Shorts(4) = 60: Shorts(6) = 50: Longs(10) = 80
        AssertNear "FRTBSSA_TimeBand 6 months, 5%", FRTBSSA_TimeBand(0.5, 0.05), 3
        AssertNear "FRTBSSA_TimeBand 25 years, 1%", FRTBSSA_TimeBand(25, 0.01), 15
        AssertNear "FRTBSSA_TimeBand floating 2.5 years", FRTBSSA_TimeBand(2.5, Empty, True), 6
        AssertText "FRTBSSA_TimeBand negative maturity", CStr(IsError(FRTBSSA_TimeBand(-1, 0.05))), "True"
        AssertNear "FRTBSSA_LadderCharge case A", FRTBSSA_LadderCharge(Longs, Shorts), 96
        AssertNear "FRTBSSA_FXCharge Table 9", FRTBSSA_FXCharge(300, 200, -35), 26.8
        AssertNear "FRTBSSA_CommodityCharge", FRTBSSA_CommodityCharge(70, 130), 14.4
        AssertText "FRTBSSA_CommodityCharge gross below net", CStr(IsError(FRTBSSA_CommodityCharge(70, 10))), "True"

End Sub


Private Sub CheckFingerprint()
'
'==============================================================================
'                               CheckFingerprint
'------------------------------------------------------------------------------
' PURPOSE
'   The fingerprint of a vector is the one its rule gives, and a changed
'   value, or two values that change places, give another one.
'
' REFERENCE
'   Each value enters as its key and code 30: "N" and Str$ of a number,
'   "S", the length, ":" and the text, "_" for an empty value, "B1" for
'   True; an array as "A" and its rows, then its values by row. Two hashes
'   start at 131 and 137 and take each code unit as h = (h * base + code)
'   modulo 2^31 - 1 (base 131) and 2^31 - 19 (base 137).
'
' UPDATED
'   2026-11-15
'==============================================================================
'
    Dim Values(1 To 6)       As Variant   'Vector: scalars, a table and a column
    Dim Table(1 To 2, 1 To 2) As Variant  'Table; its last cell stays empty
    Dim Column(1 To 3)       As Variant   'Array of one dimension

        Values(1) = 1.5
        Values(2) = "EUR"
        Values(4) = True
        Table(1, 1) = 1
        Table(1, 2) = "a"
        Table(2, 1) = -2
        Values(5) = Table
        Column(1) = 2.25
        Column(2) = "b"
        Column(3) = 3
        Values(6) = Column
        AssertText "Fingerprint of a known vector", Fingerprint(Values), "F1-68BD702560607F96"
        AssertText "Fingerprint, the same vector again", Fingerprint(Values), "F1-68BD702560607F96"
        Column(2) = "B"
        Values(6) = Column
        AssertText "Fingerprint, a text in another case", Fingerprint(Values), "F1-7833FF864C712519"
        Column(2) = "b"
        Values(6) = Column
        Table(1, 2) = -2
        Table(2, 1) = "a"
        Values(5) = Table
        AssertText "Fingerprint, two values change places", Fingerprint(Values), "F1-4BF78E8A19460933"

End Sub


Private Sub CheckFailedRun()
'
'==============================================================================
'                                CheckFailedRun
'------------------------------------------------------------------------------
' PURPOSE
'   A run that raises an error, while it writes or before, leaves no result
'   of that run or of the one before, and a status line that says FAILED.
'
' STATE OWNERSHIP
'   Rewrites the output sheets. Changes the first rule identifier of Checks
'   for one run, and the range of the name Ladder_RW for another, and puts
'   both back.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
    Dim Target      As Range     'Cell changed to make a run fail
    Dim Saved       As Variant   'Its value before the test
    Dim Ladder      As Object    'Workbook name changed to make a run fail
    Dim SavedRange  As String    'What the name referred to
    Dim ErrNumber   As Long      'Error number of the failed run

        'While writing: the rules of Checks are written after every result
        'table, and the first rule is not found under its header.
        AssertText "Run before the failure while writing", Left$(RunSSA_Silent(), 9), "RESULT=OK"
        Set Target = ThisWorkbook.Names(NM_CHECKS_HEADER).RefersToRange.Offset(1, 0)
        Saved = Target.Value
        Target.Value = "X01"
        ErrNumber = RunError()
        Target.Value = Saved
        AssertNear "Failure while writing: error number", ErrNumber, ERR_LAYOUT
        AssertNoResult "Failure while writing"

        'Before writing: a column of the maturity ladder is one row short.
        AssertText "Run before the failure while reading", Left$(RunSSA_Silent(), 9), "RESULT=OK"
        Set Ladder = ThisWorkbook.Names(NM_LADDER_RW)
        SavedRange = Ladder.RefersTo
        Set Target = Ladder.RefersToRange
        Ladder.RefersTo = "='" & Target.Worksheet.Name & "'!" & _
                          Target.Resize(Target.Rows.Count - 1, 1).Address
        ErrNumber = RunError()
        Ladder.RefersTo = SavedRange
        AssertNear "Failure while reading: error number", ErrNumber, ERR_LAYOUT
        AssertNoResult "Failure while reading"
        AssertText "Run after the failures", Left$(RunSSA_Silent(), 9), "RESULT=OK"

End Sub


Private Sub CheckOutdatedResults()
'
'==============================================================================
'                             CheckOutdatedResults
'------------------------------------------------------------------------------
' PURPOSE
'   After a run, a changed parameter or position makes the check mark the
'   results OUTDATED; the value put back makes it take the mark away.
'
' STATE OWNERSHIP
'   Rewrites the output sheets. Changes the IR scaling factor and the value
'   of the first interest rate position for one check each, and puts both
'   back.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
    Dim Target      As Range     'Input cell changed
    Dim Saved       As Variant   'Its value before the test

        AssertText "Run before the changes", Left$(RunSSA_Silent(), 9), "RESULT=OK"
        AssertText "Check after the run", CheckOutcome(), "RESULT=CURRENT; cleanup=PASS"
        AssertText "Status after the run", Left$(StatusLine(), 4), "OK: "

        'A parameter.
        Set Target = ThisWorkbook.Names("ScalingIR").RefersToRange
        Saved = Target.Value
        Target.Value = Saved + 0.1
        AssertText "Check after a parameter changed", CheckOutcome(), "RESULT=OUTDATED; cleanup=PASS"
        AssertText "Status after a parameter changed", Left$(StatusLine(), 14), "OUTDATED: OK: "
        AssertText "Check again, still changed", CheckOutcome(), "RESULT=OUTDATED; cleanup=PASS"
        AssertText "Status is marked once", Left$(StatusLine(), 14), "OUTDATED: OK: "
        Target.Value = Saved
        AssertText "Check after the parameter is put back", CheckOutcome(), "RESULT=CURRENT; cleanup=PASS"
        AssertText "Status after the parameter is put back", Left$(StatusLine(), 4), "OK: "

        'A position, through the entry point of the event procedures.
        Set Target = ThisWorkbook.Names(NM_IR_POS_HEADER).RefersToRange.Offset(1, IR_VALUE - 1)
        Saved = Target.Value
        Target.Value = Saved + 1
        CheckStatus_OnEvent
        AssertText "Status after a position changed", Left$(StatusLine(), 14), "OUTDATED: OK: "
        Target.Value = Saved
        CheckStatus_OnEvent
        AssertText "Status after the position is put back", Left$(StatusLine(), 4), "OK: "

        'No result to check.
        ClearOutputs_Silent
        AssertText "Check without results", CheckOutcome(), "RESULT=NOT_CALCULATED; cleanup=PASS"

End Sub


Private Sub CheckWorkbook()
'
'==============================================================================
'                                 CheckWorkbook
'------------------------------------------------------------------------------
' PURPOSE
'   Run the workbook on the shipped sample data and compare the totals with
'   the independent re-calculation; every TestCatalogue test must PASS.
'
' STATE OWNERSHIP
'   Rewrites the output sheets.
'
' REFERENCE
'   tools/prototype/ssa_reference.py (requirement 48,267,441; RWA
'   603,343,012.5; 35 positions used).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim ResultText  As String    'RunSSA_Silent result line
    Dim Results     As Variant   'TestCatalogue Result column
    Dim Failures    As Long      'FAIL count
    Dim Item        As Variant   'One result

        ResultText = RunSSA_Silent()
        AssertText "Workbook run outcome", Left$(ResultText, 9), "RESULT=OK"
        AssertNear "Workbook requirement", ThisWorkbook.Names("SSA_Requirement").RefersToRange.Value, _
                   SAMPLE_REQUIREMENT
        AssertNear "Workbook RWA", ThisWorkbook.Names("SSA_RWA").RefersToRange.Value, SAMPLE_RWA
        AssertNear "Workbook positions used", Val(Split(Split(ResultText, "positions_used=")(1), ";")(0)), _
                   SAMPLE_POSITIONS
        Application.Calculate
        Results = ThisWorkbook.Names("Tests_Result").RefersToRange.Value
        For Each Item In Results
            If CStr(Item) <> "PASS" Then Failures = Failures + 1
        Next Item
        AssertNear "TestCatalogue tests not PASS", Failures, 0

End Sub


'------------------------------------------------------------------------------
'                                    HELPERS
'------------------------------------------------------------------------------

Private Function RunError() As Long
'
'==============================================================================
'                                   RunError
'------------------------------------------------------------------------------
' PURPOSE
'   Run RunSSA_Silent where it is expected to raise an error.
'
' RETURNS
'   The error number; 0 when the run raised none.
'
' ERROR POLICY
'   Contains the error of the run, so that the caller can put back what it
'   changed.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
    Dim ResultText  As String    'Result line, not used

        On Error GoTo Failed
        ResultText = RunSSA_Silent()
        RunError = 0
        Exit Function

Failed:
        RunError = Err.Number

End Function


Private Function CheckOutcome() As String
'
'==============================================================================
'                                 CheckOutcome
'------------------------------------------------------------------------------
' PURPOSE
'   Run CheckStatus_Silent between a change of an input and its undoing.
'
' RETURNS
'   The result line, or "ERROR", the number and the description.
'
' ERROR POLICY
'   Contains the error of the check, so that the caller can put back what
'   it changed.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
        On Error GoTo Failed
        CheckOutcome = CheckStatus_Silent()
        Exit Function

Failed:
        CheckOutcome = "ERROR " & Err.Number & " " & Err.Description

End Function


Private Function StatusLine() As String
'
'==============================================================================
'                                  StatusLine
'------------------------------------------------------------------------------
' PURPOSE
'   Return the status line of Results.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
        StatusLine = PlainText(ThisWorkbook.Names(NM_RESULTS_STATUS).RefersToRange.Value)

End Function


Private Sub AssertNoResult( _
    ByVal Label As String)
'
'==============================================================================
'                                AssertNoResult
'------------------------------------------------------------------------------
' PURPOSE
'   Assert that the sheets show a failed run and no result: the status line,
'   the requirement, the first calculation row, the error count of Checks,
'   the fingerprint and what the check returns.
'
' INPUTS
'   Label   Start of each assertion's label.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
        AssertText Label & ": status line", Left$(StatusLine(), 8), "FAILED: "
        AssertText Label & ": requirement", _
                   PlainText(ThisWorkbook.Names("SSA_Requirement").RefersToRange.Value), vbNullString
        AssertText Label & ": first calculation row", _
                   PlainText(ThisWorkbook.Names(NM_IR_CALC_HEADER).RefersToRange.Offset(1, 0).Value), _
                   vbNullString
        AssertText Label & ": error count of Checks", _
                   PlainText(ThisWorkbook.Names(NM_CHECKS_ERRORS).RefersToRange.Value), vbNullString
        AssertText Label & ": fingerprint", _
                   PlainText(ThisWorkbook.Names(NM_FINGERPRINT).RefersToRange.Value), vbNullString
        AssertText Label & ": check", CheckOutcome(), "RESULT=FAILED; cleanup=PASS"

End Sub


Private Function RowsFrom( _
    ByVal Width As Long, _
    ByVal Records As Variant) _
    As Variant
'
'==============================================================================
'                                   RowsFrom
'------------------------------------------------------------------------------
' PURPOSE
'   Build an input table, as Range.Value would return it, from an array of
'   record arrays; missing trailing fields are Empty.
'
' INPUTS
'   Width     Columns of the table.
'   Records   Array of Array(...) records, one per row.
'
' RETURNS
'   Variant (1 To n, 1 To Width).
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Table()     As Variant   'Result
    Dim Record      As Variant   'One record
    Dim r           As Long      'Row
    Dim c           As Long      'Field

        ReDim Table(1 To UBound(Records) - LBound(Records) + 1, 1 To Width)
        For r = 1 To UBound(Table, 1)
            Record = Records(LBound(Records) + r - 1)
            For c = 1 To Width
                If c <= UBound(Record) - LBound(Record) + 1 Then
                    Table(r, c) = Record(LBound(Record) + c - 1)
                End If
            Next c
        Next r
        RowsFrom = Table

End Function


Private Sub FxTable( _
    ByVal CodeList As String, _
    ByRef Codes() As String, _
    ByRef Rates() As Double)
'
'==============================================================================
'                                    FxTable
'------------------------------------------------------------------------------
' PURPOSE
'   Build an FX table with every rate 1.
'
' INPUTS
'   CodeList   Currency codes separated by "|".
'
' RETURNS
'   Codes and Rates, base 1.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Items()     As String    'Codes
    Dim i           As Long      'Code

        Items = Split(CodeList, "|")
        ReDim Codes(1 To UBound(Items) + 1)
        ReDim Rates(1 To UBound(Items) + 1)
        For i = 1 To UBound(Items) + 1
            Codes(i) = Items(i - 1)
            Rates(i) = 1#
        Next i

End Sub


Private Sub AssertNear( _
    ByVal Label As String, _
    ByVal Actual As Variant, _
    ByVal Expected As Double)
'
'==============================================================================
'                                  AssertNear
'------------------------------------------------------------------------------
' PURPOSE
'   Record whether a number equals the expected value within TOLERANCE
'   (relative for values above 1).
'
' INPUTS
'   Label      What is checked.
'   Actual     Value produced; a non-number fails.
'   Expected   Expected value.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Allowed     As Double    'Tolerance for this value

        Allowed = TOLERANCE
        If Abs(Expected) > 1 Then Allowed = TOLERANCE * Abs(Expected)
        If Not IsNumberValue(Actual) Then
            Record False, Label & ": expected " & Expected & ", got a non-number"
        ElseIf Abs(CDbl(Actual) - Expected) > Allowed Then
            Record False, Label & ": expected " & Expected & ", got " & CDbl(Actual)
        Else
            Record True, Label
        End If

End Sub


Private Sub AssertText( _
    ByVal Label As String, _
    ByVal Actual As String, _
    ByVal Expected As String)
'
'==============================================================================
'                                  AssertText
'------------------------------------------------------------------------------
' PURPOSE
'   Record whether a text equals the expected text exactly.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        If Actual = Expected Then
            Record True, Label
        Else
            Record False, Label & ": expected """ & Expected & """, got """ & Actual & """"
        End If

End Sub


Private Sub Record( _
    ByVal Passed As Boolean, _
    ByVal Message As String)
'
'==============================================================================
'                                    Record
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and keep the message of a failure.
'
' STATE OWNERSHIP
'   Updates mPassed, mFailed and mReport.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        If Passed Then
            mPassed = mPassed + 1
        Else
            mFailed = mFailed + 1
            mReport = mReport & vbLf & "FAIL " & Message
        End If

End Sub


Private Sub ResetResults()
'
'==============================================================================
'                                 ResetResults
'------------------------------------------------------------------------------
' PURPOSE
'   Start a new run.
'
' STATE OWNERSHIP
'   Resets mPassed, mFailed and mReport.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        mPassed = 0
        mFailed = 0
        mReport = vbNullString

End Sub


Private Function ResultLine() As String
'
'==============================================================================
'                                  ResultLine
'------------------------------------------------------------------------------
' PURPOSE
'   Summarise the run.
'
' RETURNS
'   "PASS=n; FAIL=m" and the failure messages.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        ResultLine = "PASS=" & mPassed & "; FAIL=" & mFailed & mReport

End Function
