Attribute VB_Name = "TEST_Suite"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Suite
'------------------------------------------------------------------------------
' PURPOSE
'   The core suite of the workbook's tests: the K-FRTB-SSA core and facade on
'   hand-checked inputs, as cases and assertions of TEST_Harness. No case reads
'   or writes a sheet; the tests of the whole workbook are in TEST_Workbook.
'
' WHY THIS EXISTS
'   Checks the core procedures and the worksheet functions of the facade
'   against values worked out by hand, on inputs written in each case, so that
'   the calculation is tested without a sheet; the tests that use the sheets
'   are in TEST_Workbook.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunCoreTests is the documented test entry point: it runs the suite and
'   raises when the run does not pass. Test_RunCore runs the same suite and
'   returns its result line where a caller cannot read an error, as
'   tools/devbuild/lo_macro.py. CertifyCore certifies the run and saves its log
'   and its certification log as evidence of an Excel run.
'
'   RunSuite runs these cases between BeginRun and EndRun, in this order:
'     - CaseTimeBand: band boundaries of MAR40.26 table 4
'     - CaseLadderCharges: general market risk of hand-calculated ladders,
'       under the Basel and the CRR regime
'     - CaseForeignExchange: MAR40.61 table 9, an excluded row and a duplicate
'       row
'     - CaseEquity: netting of the same stock and the treatment of an index
'     - CaseCommodity: the simplified approach on one commodity entered in two
'       cases
'     - CaseInterestRateRows: row validation, slotting and specific risk of
'       interest rate rows
'     - CaseSpecificRiskCrr: specific risk under the CRR regime
'     - CaseIssueNetting: specific-risk netting through an entered Issue ID
'     - CaseDateBounds: a date typed as a number, and the last date Excel holds
'     - CaseParameters: the reference parameters, an unusable value and a
'       faulty FX table
'     - CaseRegime: the regime, the transitional factor and the de minimis of
'       Article 351
'     - CaseVolume: the largest volume of a position table the workbook
'       supports
'     - CaseFacade: the worksheet functions and their error values
'     - CaseFingerprint: the fingerprint of a vector
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED, or
'   the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START. An unexpected error inside a case is
'   recorded as a failure of that case, and the run continues. Test_RunCore
'   contains those errors; CertifyCore records them, saves the log, and raises
'   only when the certification did not pass.
'
' DEPENDENCIES
'   TEST_Harness; CORE_Config, CORE_Util, CORE_Parameters, CORE_Engine,
'   CORE_InterestRate, CORE_Equity, CORE_Commodity, CORE_ForeignExchange and
'   FRTBSSA_Formulas.
'
'   Reference: Expected values: hand calculation, stated in each case, and
'   MAR40.61 table 9. The fingerprints of CaseFingerprint were calculated apart
'   from the VBA code, from the rule its header states.
'
' NOTES
'   Option Explicit requires explicit declarations; Option Private Module keeps
'   the entry points out of the workbook's external surface and the tests
'   inside this project. Never part of the production import set.
'
'   State ownership: none of its own. The harness holds the run's state.
'
'   Compatibility: Excel VBA; also runs under LibreOffice's VBA mode through
'   the tools of tools/devbuild/, which is never Excel evidence.
'
'   Usage: in a development workbook, type TEST_Suite.RunCoreTests in the
'   Immediate window. To record an Excel run, type
'
'       TEST_Suite.CertifyCore "<folder>", "<checkout>\tests\fixtures"
'
'   (INSTALLATION.md, Recording an Excel run).
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-17
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Name and size of a complete run, which starts every case and evaluates every assertion once
    Private Const SUITE_NAME          As String = "core"        'Suite name given to BeginRun and BeginCertification
    Private Const EXPECTED_CASES      As Long = 14              'Cases of a complete run
    Private Const EXPECTED_ASSERTIONS As Long = 149             'Assertions of a complete run

    'Tolerances of a derived value (docs/methodology/CALCULATION_CONTRACTS.md, Decisions)
    Private Const TOL_ABSOLUTE        As Double = 0.000001      'Absolute tolerance
    Private Const TOL_RELATIVE        As Double = 0.000000001   'Relative tolerance


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunCoreTests()

'
'------------------------------------------------------------------------------
'                                RUN CORE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Run the core suite. A passing run ends with
'   RESULT=PASS; suite=core; operation=PASS; completeness=COMPLETE;
'   cases=14; assertions=149; failures=0; cleanup=PASS; recovery=NONE.
'
' WHY THIS EXISTS
'   It is the documented entry point of the core suite: typed in the Immediate
'   window of a development workbook, it runs the suite and raises when the run
'   does not pass.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the suite with RunSuite.
'
' ERROR POLICY
'   As TEST_Harness.EndRun: raises when the run does not pass.
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   Test_RunCore runs the same suite for a caller that reads a result and not
'   an error; CertifyCore certifies the run.
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run every case of the core suite
        RunSuite

End Sub

Public Function Test_RunCore() As String

'
'------------------------------------------------------------------------------
'                            RUN SUITE FOR A RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Run the core suite for a caller that reads a result and not an error.
'
' WHY THIS EXISTS
'   A caller such as tools/devbuild/lo_macro.py, which runs the suite in
'   LibreOffice's VBA mode, reads the string a macro returns and cannot read an
'   error.
'
' INPUTS
'   None
'
' RETURNS
'   The RESULT line of the run, as TEST_Harness.LastResult gives it.
'
' BEHAVIOR
'   Runs the suite with RunSuite while errors are suppressed, restores normal
'   error handling and returns the result line of the run.
'
' ERROR POLICY
'   Contains the error of a run that does not pass, or of a refused start: the
'   result line says so.
'
' DEPENDENCIES
'   RunSuite, TEST_Harness.LastResult
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run the suite and contain its error, so that the result line reports the outcome
        On Error Resume Next
        RunSuite
        On Error GoTo 0

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return the result line of the run
        Test_RunCore = TEST_Harness.LastResult()

End Function

Public Sub CertifyCore( _
    ByVal Folder As String, _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                              CERTIFY CORE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Certify the core suite's run as evidence of an Excel run: run it, save its
'   log as core.log, and write core-certification.log.
'
' WHY THIS EXISTS
'   A run recorded as evidence needs its log and its certification log saved
'   even when the run fails, so the error of the run is recorded and not
'   raised.
'
' INPUTS
'   Folder
'     An existing folder for the logs, without a trailing separator.
'
'   Fixtures
'     The full path of tests/fixtures/ in the checkout; not read, and taken so
'     that every certification entry point has the same two arguments.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the certification, runs the suite with errors suppressed and keeps
'   the number and the description of the error that ended it, saves the log as
'   core.log, echoes where it was saved, certifies the run in NORMAL mode and
'   ends the certification by writing core-certification.log.
'
' ERROR POLICY
'   The error of the run is recorded, not raised, so that the log is always
'   saved. Raises ERR_TEST_FAILED after writing the certification log when the
'   certification did not pass. An error of BeginCertification or SaveLog
'   propagates; run TEST_Harness.ResetTests before trying again.
'
' DEPENDENCIES
'   RunSuite; TEST_Harness (BeginCertification, SaveLog, Echo, CertifyRun,
'   EndCertification, ResetTests)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RunErrorNumber      As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription As String   'Its description

'------------------------------------------------------------------------------
' BEGIN CERTIFICATION
'------------------------------------------------------------------------------
    'Begin the certification of the suite
        TEST_Harness.BeginCertification SUITE_NAME

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run the suite with errors suppressed, so that the log is always saved
        On Error Resume Next
        RunSuite
    'Keep the error that ended the run, if any
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' SAVE LOG
'------------------------------------------------------------------------------
    'Save the log of the run
        TEST_Harness.SaveLog Folder & "\core.log"
    'Echo where the log was saved and the error of the run
        TEST_Harness.Echo "SAVED=" & Folder & "\core.log; mode=NORMAL; error=" & _
            CStr(RunErrorNumber) & IIf(RunErrorNumber = 0, vbNullString, " " & RunErrorDescription)

'------------------------------------------------------------------------------
' CERTIFY RUN
'------------------------------------------------------------------------------
    'Certify the run in normal mode with the error of the run
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
    'End the certification and write its log
        TEST_Harness.EndCertification Folder & "\core-certification.log"

End Sub

'
'------------------------------------------------------------------------------
'
'                                    RUNNER
'
'------------------------------------------------------------------------------
'

Private Sub RunSuite()

'
'------------------------------------------------------------------------------
'                                RUN ALL CASES
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun.
'
' WHY THIS EXISTS
'   The three entry points run the same sequence of cases, and a run needs one
'   place that begins it, calls each case once and ends it with the error that
'   escaped, if any.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the run with SUITE_NAME, EXPECTED_CASES and EXPECTED_ASSERTIONS,
'   calls the cases in this order, each alone on its line: CaseTimeBand,
'   CaseLadderCharges, CaseForeignExchange, CaseEquity, CaseCommodity,
'   CaseInterestRateRows, CaseSpecificRiskCrr, CaseIssueNetting,
'   CaseDateBounds, CaseParameters, CaseRegime, CaseVolume, CaseFacade and
'   CaseFingerprint, and ends the run with EndRun.
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it and
'   re-raises it unchanged. BeginRun's refusal propagates directly.
'
' DEPENDENCIES
'   TEST_Harness (BeginRun, EndRun); the Case procedures of this module
'
' NOTES
'   tools/check_excel_run.py and the tests read the cases from the lines of
'   this procedure that hold one Case procedure name each.
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' BEGIN RUN
'------------------------------------------------------------------------------
    'Begin the run and send an error of a case to the runner handler
        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS
        On Error GoTo RunFailed

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Call each case once, in a fixed order
        CaseTimeBand
        CaseLadderCharges
        CaseForeignExchange
        CaseEquity
        CaseCommodity
        CaseInterestRateRows
        CaseSpecificRiskCrr
        CaseIssueNetting
        CaseDateBounds
        CaseParameters
        CaseRegime
        CaseVolume
        CaseFacade
        CaseFingerprint

'------------------------------------------------------------------------------
' END RUN
'------------------------------------------------------------------------------
    'Resume here after the handler, with normal error handling restored
Finish:
    'End the run with the error that escaped, if any
        On Error GoTo 0
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Save the error that escaped a case and resume at the end of the run
RunFailed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume Finish

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseTimeBand()

'
'------------------------------------------------------------------------------
'                          CHECK TIME BAND BOUNDARIES
'------------------------------------------------------------------------------
' PURPOSE
'   Band boundaries of MAR40.26 table 4: an upper bound belongs to its band.
'
' WHY THIS EXISTS
'   The band of a position decides its risk weight, so the boundaries between
'   the bands are fixed by a test.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case time-band, loads the reference ladder and checks the band
'   of the boundary maturities, for a high and for a low coupon.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   ReferenceLadder, TimeBand; TEST_Harness (BeginCase, AssertEqualDouble,
'   RecordCaseError)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ladder()    As Double   'MAR40 ladder

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "time-band"

'------------------------------------------------------------------------------
' LOAD REFERENCE DATA
'------------------------------------------------------------------------------
    'Load the reference ladder of MAR40
        ReferenceLadder Ladder

'------------------------------------------------------------------------------
' CHECK BAND BOUNDARIES
'------------------------------------------------------------------------------
    'Zero and one month belong to band 1: an upper bound belongs to its band
        TEST_Harness.AssertEqualDouble "at-zero", 1, TimeBand(0#, True, Ladder)
        TEST_Harness.AssertEqualDouble "at-one-month", 1, TimeBand(1# / 12#, True, Ladder)
    'One day after one month is in band 2
        TEST_Harness.AssertEqualDouble "one-day-after-one-month", 2, _
            TimeBand(1# / 12# + 1# / 365#, True, Ladder)
    'Maturities with a high coupon
        TEST_Harness.AssertEqualDouble "one-year-high-coupon", 4, TimeBand(1#, True, Ladder)
        TEST_Harness.AssertEqualDouble "one-and-a-half-years-high-coupon", 5, _
            TimeBand(1.5, True, Ladder)
        TEST_Harness.AssertEqualDouble "twenty-years-high-coupon", 12, TimeBand(20#, True, Ladder)
        TEST_Harness.AssertEqualDouble "twenty-five-years-high-coupon", 13, TimeBand(25#, True, Ladder)
    'Maturities with a low coupon
        TEST_Harness.AssertEqualDouble "one-point-nine-five-years-low-coupon", 6, _
            TimeBand(1.95, False, Ladder)
        TEST_Harness.AssertEqualDouble "twelve-years-low-coupon", 13, TimeBand(12#, False, Ladder)
        TEST_Harness.AssertEqualDouble "twenty-five-years-low-coupon", 15, TimeBand(25#, False, Ladder)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseLadderCharges()

'
'------------------------------------------------------------------------------
'                             CHECK LADDER CHARGES
'------------------------------------------------------------------------------
' PURPOSE
'   General market risk of hand-calculated ladders.
'
' WHY THIS EXISTS
'   The general market risk of the interest rate class is built from vertical,
'   horizontal and net components, which hand-calculated ladders check one by
'   one, under the Basel and the CRR regime.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case ladder-charges, loads the reference ladder and parameters,
'   and calculates and checks ladder A, ladder B and ladder C under the Basel
'   regime; then selects the CRR regime and checks ladder C and ladder A again.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   ReferenceLadder, ReferenceParameters, LadderCharges; TEST_Harness
'   (BeginCase, AssertNear, RecordCaseError)
'
'   Reference: Ladder A: band 4 long 100 / short 60, band 6 short 50, band 10
'   long 80. Vertical 10% x 60 = 6; zones +40, -50, +80; offset 1-2 = 40, 2-3 =
'   10, 1-3 = 0; between 40% x 50 = 20; net |70| = 70; total 96. Ladder B: band
'   1 long 10, band 2 short 4: within zone 1, 40% x 4 = 1.6; net 6; total 7.6.
'   Ladder C: band 3 long 30, band 12 short 50: offset 1-3 = 30 at 100%; net
'   20; total 50. Under the CRR the offset is charged at 150%: 45; net 20;
'   total 65. Ladder A has no offset between zones 1 and 3, and its total is 96
'   under both regimes.
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
    Dim Ladder()        As Double   'MAR40 ladder
    Dim Parameters()    As Double   'MAR40 parameters
    Dim Longs()         As Double   'Weighted longs
    Dim Shorts()        As Double   'Weighted shorts
    Dim Components()    As Double   'Result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "ladder-charges"

'------------------------------------------------------------------------------
' LOAD REFERENCE DATA
'------------------------------------------------------------------------------
    'Load the reference parameters and the reference ladder
        ReferenceLadder Ladder
        ReferenceParameters Parameters

'------------------------------------------------------------------------------
' CHECK LADDER A
'------------------------------------------------------------------------------
    'Weight ladder A: band 4 long 100 and short 60, band 6 short 50, band 10 long 80
        ReDim Longs(1 To BAND_COUNT)
        ReDim Shorts(1 To BAND_COUNT)
        Longs(4) = 100: Shorts(4) = 60: Shorts(6) = 50: Longs(10) = 80
        LadderCharges Longs, Shorts, Ladder, Parameters, Components
    'Check the vertical, between-zones, net and total charges
        TEST_Harness.AssertNear "a-vertical", 6, Components(ZC_VERTICAL), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "a-between-zones", 20, _
            Components(ZC_BETWEEN), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "a-net", 70, Components(ZC_NET), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "a-total", 96, Components(ZC_GENERAL), TOL_ABSOLUTE, TOL_RELATIVE

'------------------------------------------------------------------------------
' CHECK LADDER B
'------------------------------------------------------------------------------
    'Weight ladder B: band 1 long 10 and band 2 short 4
        ReDim Longs(1 To BAND_COUNT)
        ReDim Shorts(1 To BAND_COUNT)
        Longs(1) = 10: Shorts(2) = 4
        LadderCharges Longs, Shorts, Ladder, Parameters, Components
    'Check the charge within zone 1 and the total
        TEST_Harness.AssertNear "b-within-zone-1", 1.6, _
            Components(ZC_ZONE_FIRST + 3), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "b-total", 7.6, Components(ZC_GENERAL), TOL_ABSOLUTE, TOL_RELATIVE

'------------------------------------------------------------------------------
' CHECK LADDER C
'------------------------------------------------------------------------------
    'Weight ladder C: band 3 long 30 and band 12 short 50
        ReDim Longs(1 To BAND_COUNT)
        ReDim Shorts(1 To BAND_COUNT)
        Longs(3) = 30: Shorts(12) = 50
        LadderCharges Longs, Shorts, Ladder, Parameters, Components
    'Check the offset between zones 1 and 3 and the total
        TEST_Harness.AssertNear "c-matched-zones-1-and-3", 30, _
            Components(ZC_MATCHED13), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "c-total", 50, Components(ZC_GENERAL), TOL_ABSOLUTE, TOL_RELATIVE

'------------------------------------------------------------------------------
' CHECK CRR REGIME
'------------------------------------------------------------------------------
    'Select the CRR regime, which charges the offset between zones 1 and 3 at 150% (Article 339(9)(f))
        Parameters(PI_REGIME_FLAG) = 1#
        LadderCharges Longs, Shorts, Ladder, Parameters, Components
    'Check the charge between zones and the total under the CRR
        TEST_Harness.AssertNear "c-between-zones-crr", 45, _
            Components(ZC_BETWEEN), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "c-total-crr", 65, Components(ZC_GENERAL), TOL_ABSOLUTE, TOL_RELATIVE
    'Weight ladder A again: it has no such offset, so its total is the same
        ReDim Longs(1 To BAND_COUNT)
        ReDim Shorts(1 To BAND_COUNT)
        Longs(4) = 100: Shorts(4) = 60: Shorts(6) = 50: Longs(10) = 80
        LadderCharges Longs, Shorts, Ladder, Parameters, Components
    'Check the total of ladder A under the CRR
        TEST_Harness.AssertNear "a-total-crr", 96, Components(ZC_GENERAL), TOL_ABSOLUTE, TOL_RELATIVE

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseForeignExchange()

'
'------------------------------------------------------------------------------
'                         CHECK FOREIGN EXCHANGE CLASS
'------------------------------------------------------------------------------
' PURPOSE
'   MAR40.61 table 9 at rate 1, plus an excluded and a duplicate row.
'
' WHY THIS EXISTS
'   Checks the foreign exchange class against the worked example of MAR40.61
'   table 9, with the rows that the class excludes or rejects.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case foreign-exchange, enters the positions of table 9 with a
'   row in the reporting currency and a duplicate PositionID, calculates the
'   class and checks the sums, the net open position, the charge, the row
'   counts and the status of the excluded and the duplicate row.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   ReferenceParameters, FxTable, RowsFrom, CalculateForeignExchange;
'   TEST_Harness (BeginCase, AssertNear, AssertEqualDouble, AssertEqualText,
'   RecordCaseError)
'
'   Reference: Table 9: longs 50 + 100 + 150 = 300, shorts 20 + 180 = 200, gold
'   35; charge 8% x (300 + 35) = 26.8.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InputRows()     As Variant  'FXPositions cells
    Dim FxCodes()       As String   'FX table codes
    Dim FxRates()       As Double   'FX table rates
    Dim Parameters()    As Double   'MAR40 parameters
    Dim RowOutput()     As Variant  'FXCalc
    Dim Currencies()    As Variant  'Buckets FX block
    Dim Totals()        As Double   'Totals

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "foreign-exchange"

'------------------------------------------------------------------------------
' LOAD REFERENCE DATA
'------------------------------------------------------------------------------
    'Load the reference parameters and an FX table with every rate 1
        ReferenceParameters Parameters
        FxTable "EUR|CAD|SEK|GBP|NOK|USD|XAU", FxCodes, FxRates

'------------------------------------------------------------------------------
' BUILD POSITIONS
'------------------------------------------------------------------------------
    'Enter the positions of table 9, a row in the reporting currency and a duplicate PositionID
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

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Calculate the foreign exchange class
        CalculateForeignExchange InputRows, Parameters, FxCodes, FxRates, "EUR", "XAU", _
                                 RowOutput, Currencies, Totals

'------------------------------------------------------------------------------
' CHECK TOTALS
'------------------------------------------------------------------------------
    'Check the sums, the overall net open position and the charge of table 9
        TEST_Harness.AssertNear "sum-of-longs", 300, Totals(TOT_LONG), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "sum-of-shorts", 200, Totals(TOT_SHORT), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "gold", 35, Totals(TOT_GOLD), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "overall-net-open-position", 335, _
            Totals(TOT_ONOP), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "charge-of-table-9", 26.8, _
            Totals(TOT_GENERAL), TOL_ABSOLUTE, TOL_RELATIVE
    'Check the counts of used, excluded and rejected rows
        TEST_Harness.AssertEqualDouble "rows-used", 6, Totals(TOT_USED)
        TEST_Harness.AssertEqualDouble "rows-excluded", 1, Totals(TOT_EXCLUDED)
        TEST_Harness.AssertEqualDouble "rows-rejected", 2, Totals(TOT_REJECTED)

'------------------------------------------------------------------------------
' CHECK ROW STATUS
'------------------------------------------------------------------------------
    'Check the status of the reporting currency row and of the duplicate
        TEST_Harness.AssertEqualText "reporting-currency-status", _
            ST_EXCLUDED & "reporting currency carries no FX risk", CStr(RowOutput(7, 2))
        TEST_Harness.AssertEqualText "duplicate-status", _
            ST_INVALID & "duplicate PositionID", CStr(RowOutput(9, 2))

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseEquity()

'
'------------------------------------------------------------------------------
'                              CHECK EQUITY CLASS
'------------------------------------------------------------------------------
' PURPOSE
'   Netting of the same stock, whatever the case of its codes, and the
'   treatment of an index.
'
' WHY THIS EXISTS
'   Checks that the same stock is netted whatever the case of its codes, that
'   an index has its own charge, and that the CRR regime removes that charge
'   (Article 344).
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case equity, enters one stock twice, an index, a basket and a
'   stock in gold, calculates the class and checks the charges and the rejected
'   rows under the Basel regime; then selects the CRR regime, calculates again
'   and checks the charges.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   ReferenceParameters, FxTable, RowsFrom, CalculateEquity; TEST_Harness
'   (BeginCase, AssertNear, AssertEqualDouble, AssertEqualText,
'   RecordCaseError)
'
'   Reference: Stock X: +100 - 40 = 60, specific 8% x 60 = 4.8; index +50, 2% =
'   1; general 8% x |60 + 50| = 8.8 (MAR40.43, 40.46, 40.47). Under the CRR the
'   index has no charge, 0; the specific risk 4.8 and the general risk 8.8 are
'   the same (Article 344).
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
    Dim InputRows()     As Variant  'EQPositions cells
    Dim FxCodes()       As String   'FX table codes
    Dim FxRates()       As Double   'FX table rates
    Dim Parameters()    As Double   'MAR40 parameters
    Dim RowOutput()     As Variant  'EQCalc
    Dim Markets()       As Variant  'Buckets equity block
    Dim Totals()        As Double   'Totals

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "equity"

'------------------------------------------------------------------------------
' LOAD REFERENCE DATA
'------------------------------------------------------------------------------
    'Load the reference parameters and an FX table with every rate 1
        ReferenceParameters Parameters
        FxTable "EUR|XAU", FxCodes, FxRates

'------------------------------------------------------------------------------
' BUILD POSITIONS
'------------------------------------------------------------------------------
    'Enter stock X long 100 and short 40 in two cases, an index, a basket and a stock in gold
        InputRows = RowsFrom(EQ_INPUT_WIDTH, Array( _
            Array("E1", "", "IT", "Single name", "X", "EUR", "Long", 100), _
            Array("E2", "", "it", "single name", "x", "EUR", "Short", 40), _
            Array("E3", "", "IT", "Index", "IDX", "EUR", "Long", 50), _
            Array("E4", "", "IT", "Basket", "B", "EUR", "Long", 50), _
            Array("E5", "", "IT", "Single name", "Y", "XAU", "Long", 50)))

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Calculate the equity class
        CalculateEquity InputRows, Parameters, FxCodes, FxRates, "XAU", RowOutput, Markets, Totals

'------------------------------------------------------------------------------
' CHECK BASEL RESULTS
'------------------------------------------------------------------------------
    'Check the specific, index and general charges and the number of markets
        TEST_Harness.AssertNear "specific", 4.8, Totals(TOT_SPECIFIC), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "index", 1, Totals(TOT_OTHER), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "general", 8.8, Totals(TOT_GENERAL), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertEqualDouble "markets", 1, Totals(TOT_BUCKETS)
    'Check the rejected rows and the status of the basket and of the stock in gold
        TEST_Harness.AssertEqualDouble "rows-rejected", 2, Totals(TOT_REJECTED)
        TEST_Harness.AssertEqualText "type-status", _
            ST_INVALID & "Type must be Single name or Index", CStr(RowOutput(4, 2))
        TEST_Harness.AssertEqualText "gold-currency-status", _
            ST_INVALID & "currency not in FX_Table", CStr(RowOutput(5, 2))

'------------------------------------------------------------------------------
' CHECK CRR RESULTS
'------------------------------------------------------------------------------
    'Under the CRR the index carries no charge of its own (Article 344)
    'Its general risk stays in the net position of the market
    'Select the CRR regime and calculate again
        Parameters(PI_REGIME_FLAG) = 1#
        CalculateEquity InputRows, Parameters, FxCodes, FxRates, "XAU", RowOutput, Markets, Totals
    'Check the index, specific and general charges under the CRR
        TEST_Harness.AssertNear "index-crr", 0, Totals(TOT_OTHER), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "specific-crr", 4.8, Totals(TOT_SPECIFIC), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "general-crr", 8.8, Totals(TOT_GENERAL), TOL_ABSOLUTE, TOL_RELATIVE

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCommodity()

'
'------------------------------------------------------------------------------
'                            CHECK COMMODITY CLASS
'------------------------------------------------------------------------------
' PURPOSE
'   Simplified approach on one commodity entered in two cases.
'
' WHY THIS EXISTS
'   Checks the simplified approach on one commodity entered in two cases, with
'   a row rejected for its quantity and a row in gold.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case commodity, enters oil in two cases, a row with a text
'   quantity and a row in gold, calculates the class and checks the charges,
'   the number of commodities and the status of the rejected rows.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   ReferenceParameters, FxTable, RowsFrom, CalculateCommodity; TEST_Harness
'   (BeginCase, AssertNear, AssertEqualDouble, AssertEqualText,
'   RecordCaseError)
'
'   Reference: Oil +100, -30 at spot: 15% x 70 = 10.5; 3% x 130 = 3.9
'   (MAR40.72-40.73).
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InputRows()     As Variant  'COPositions cells
    Dim FxCodes()       As String   'FX table codes
    Dim FxRates()       As Double   'FX table rates
    Dim Parameters()    As Double   'MAR40 parameters
    Dim RowOutput()     As Variant  'COCalc
    Dim Commodities()   As Variant  'Buckets commodity block
    Dim Totals()        As Double   'Totals

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "commodity"

'------------------------------------------------------------------------------
' LOAD REFERENCE DATA
'------------------------------------------------------------------------------
    'Load the reference parameters and an FX table with every rate 1
        ReferenceParameters Parameters
        FxTable "EUR|XAU", FxCodes, FxRates

'------------------------------------------------------------------------------
' BUILD POSITIONS
'------------------------------------------------------------------------------
    'Enter oil long and short in two cases, a row with a text quantity and a row in gold
        InputRows = RowsFrom(CO_INPUT_WIDTH, Array( _
            Array("C1", "", "Oil", "Long", 10, "bbl", 10, "EUR"), _
            Array("C2", "", "OIL ", "Short", 3, "bbl", 10, "EUR"), _
            Array("C3", "", "Oil", "Long", "ten", "bbl", 10, "EUR"), _
            Array("C4", "", " xau ", "Long", 1, "oz", 3500, "EUR")))

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Calculate the commodity class
        CalculateCommodity InputRows, Parameters, FxCodes, FxRates, "XAU", RowOutput, Commodities, _
                           Totals

'------------------------------------------------------------------------------
' CHECK RESULTS
'------------------------------------------------------------------------------
    'Check the directional and gross charges and the number of commodities
        TEST_Harness.AssertNear "directional", 10.5, Totals(TOT_GENERAL), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "gross", 3.9, Totals(TOT_OTHER), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertEqualDouble "commodities", 1, Totals(TOT_BUCKETS)
    'Check the status of the row with a text quantity and of the row in gold
        TEST_Harness.AssertEqualText "quantity-status", _
            ST_INVALID & "Quantity must be a number", CStr(RowOutput(3, 2))
        TEST_Harness.AssertEqualText "gold-status", _
            ST_INVALID & "gold is FX risk; enter it on FXPositions", CStr(RowOutput(4, 2))

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseInterestRateRows()

'
'------------------------------------------------------------------------------
'                           CHECK INTEREST RATE ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Row validation, slotting and specific risk of interest rate rows.
'
' WHY THIS EXISTS
'   Checks the validation of position rows, the slotting of a row into a time
'   band and its specific risk charge against one hand-calculated row, with the
'   rows that the validation rejects.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case interest-rate-rows, enters a valid bond row and six rows
'   that the validation rejects, calculates the class and checks the band, the
'   weighted position and the specific risk of the valid row, the status of
'   each rejected row and the totals.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   ReferenceParameters, ReferenceLadder, FxTable, RowsFrom, NumberOf,
'   CalculateInterestRate; TEST_Harness (BeginCase, AssertNear,
'   AssertEqualDouble, AssertEqualText, RecordCaseError)
'
'   Reference: Row 1: QUAL A+ to BBB-, 5 years 1 day to maturity, 4% coupon:
'   band 9 (5-7 years), weight 3.25%, weighted 32,500; specific 1.6% x
'   1,000,000 = 16,000 (MAR40.6, 40.26).
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InputRows()     As Variant  'IRPositions cells
    Dim FxCodes()       As String   'FX table codes
    Dim FxRates()       As Double   'FX table rates
    Dim Parameters()    As Double   'MAR40 parameters
    Dim Ladder()        As Double   'MAR40 ladder
    Dim Labels()        As String   'Ladder labels
    Dim SrKeys()        As String   'Specific-risk keys
    Dim SrCharges()     As Double   'Specific-risk charges
    Dim RowOutput()     As Variant  'IRCalc
    Dim LadderOutput()  As Variant  'IR_Ladder
    Dim ZoneOutput()    As Variant  'IR_Zones
    Dim Totals()        As Double   'Totals
    Dim AsOfDate        As Double   'Reporting date

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "interest-rate-rows"

'------------------------------------------------------------------------------
' LOAD REFERENCE DATA
'------------------------------------------------------------------------------
    'Load the reference data, size the ladder labels and build an FX table with every rate 1
        ReferenceParameters Parameters
        ReferenceLadder Ladder
        ReDim Labels(1 To BAND_COUNT, 1 To 2)
        FxTable "EUR|XAU", FxCodes, FxRates

'------------------------------------------------------------------------------
' BUILD TABLES
'------------------------------------------------------------------------------
    'Size the specific-risk table and enter its two charges
        ReDim SrKeys(1 To 2)
        ReDim SrCharges(1 To 2)
        SrKeys(1) = "QUAL|A+ TO BBB-|>24M": SrCharges(1) = 0.016
        SrKeys(2) = "GOV|AAA TO AA-|ANY": SrCharges(2) = 0
    'Set the reporting date
        AsOfDate = CDbl(DateSerial(2026, 9, 30))

'------------------------------------------------------------------------------
' BUILD POSITIONS
'------------------------------------------------------------------------------
    'Enter one valid bond row and six rows that the validation rejects
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
                  Empty, "GOV", "AAA to AA-", ""), _
            Array("R6", "", "Securitisation", "EUR", "Long", 1000, "Fixed", 0.05, _
                  DateSerial(2031, 10, 1), Empty, "GOV", "AAA to AA-", ""), _
            Array("R7", "", Empty, "EUR", "Long", 1000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "GOV", "AAA to AA-", "")))

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Calculate the interest rate class
        CalculateInterestRate InputRows, AsOfDate, Parameters, FxCodes, FxRates, "XAU", SrKeys, _
                              SrCharges, Ladder, Labels, RowOutput, LadderOutput, ZoneOutput, Totals

'------------------------------------------------------------------------------
' CHECK VALID ROW
'------------------------------------------------------------------------------
    'Check the status, the band, the weighted position and the specific risk of row 1
        TEST_Harness.AssertEqualText "row-1-status", ST_OK, CStr(RowOutput(1, 2))
        TEST_Harness.AssertEqualDouble "row-1-band", 9, NumberOf(RowOutput(1, 9))
        TEST_Harness.AssertNear "row-1-weighted", 32500, _
            NumberOf(RowOutput(1, 12)), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "row-1-specific", 16000, _
            NumberOf(RowOutput(1, 20)), TOL_ABSOLUTE, TOL_RELATIVE

'------------------------------------------------------------------------------
' CHECK REJECTED ROWS
'------------------------------------------------------------------------------
    'Check the status of each rejected row
        TEST_Harness.AssertEqualText "row-2-status", _
            ST_INVALID & "Coupon required for a fixed-rate position", CStr(RowOutput(2, 2))
        TEST_Harness.AssertEqualText "row-3-status", _
            ST_INVALID & "repricing date outside AsOfDate to final maturity", CStr(RowOutput(3, 2))
        TEST_Harness.AssertEqualText "row-4-status", _
            ST_INVALID & "no MAR40.6 charge for this category and rating grade", CStr(RowOutput(4, 2))
        TEST_Harness.AssertEqualText "row-5-status", _
            ST_INVALID & "gold is FX risk; enter it on FXPositions", CStr(RowOutput(5, 2))
        TEST_Harness.AssertEqualText "row-6-status", _
            ST_INVALID & "Instrument must be Bond, FRN, Swap leg, Future leg, FRA leg, " & _
            "Forward leg or Other", CStr(RowOutput(6, 2))
        TEST_Harness.AssertEqualText "row-7-status", _
            ST_INVALID & "Instrument must be Bond, FRN, Swap leg, Future leg, FRA leg, " & _
            "Forward leg or Other", CStr(RowOutput(7, 2))

'------------------------------------------------------------------------------
' CHECK TOTALS
'------------------------------------------------------------------------------
    'Check the specific and general totals and the number of rejected rows
        TEST_Harness.AssertNear "specific-total", 16000, _
            Totals(TOT_SPECIFIC), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "general-total", 32500, Totals(TOT_GENERAL), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertEqualDouble "rows-rejected", 6, Totals(TOT_REJECTED)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseSpecificRiskCrr()

'
'------------------------------------------------------------------------------
'                           CHECK CRR SPECIFIC RISK
'------------------------------------------------------------------------------
' PURPOSE
'   Specific risk under the CRR regime: the category CRR with the risk weight
'   as its grade, the three ranges of the residual term, a weight that the
'   table does not list, and the categories of the other regime.
'
' WHY THIS EXISTS
'   Pins the specific risk charges of the CRR regime to Article 336, and the
'   categories that each regime accepts.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case specific-risk-crr, selects the CRR regime, enters eight
'   rows, calculates the interest rate class and checks the charge of each risk
'   weight, the rows that the CRR rejects and the totals; then selects the
'   Basel regime, calculates two rows and checks their status.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   ReferenceParameters, ReferenceLadder, FxTable, RowsFrom, NumberOf,
'   CalculateInterestRate; TEST_Harness (BeginCase, AssertNear,
'   AssertEqualDouble, AssertEqualText, RecordCaseError)
'
'   Reference: Article 336(1), table 1, and (3): 0% at a weight of 0%; 0.25%,
'   1.00% and 1.60% at 20% and 50%, by the residual term; half of that at 10%;
'   8% at 100%, and, by decision (#79), at 75%; 12% at 150%. Each position is
'   1,000,000; the rows without an Issue ID do not net.
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
    Dim InputRows()     As Variant  'IRPositions cells
    Dim FxCodes()       As String   'FX table codes
    Dim FxRates()       As Double   'FX table rates
    Dim Parameters()    As Double   'MAR40 parameters
    Dim Ladder()        As Double   'MAR40 ladder
    Dim Labels()        As String   'Ladder labels
    Dim SrKeys()        As String   'Specific-risk keys
    Dim SrCharges()     As Double   'Specific-risk charges
    Dim RowOutput()     As Variant  'IRCalc
    Dim LadderOutput()  As Variant  'IR_Ladder
    Dim ZoneOutput()    As Variant  'IR_Zones
    Dim Totals()        As Double   'Totals
    Dim AsOfDate        As Double   'Reporting date

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "specific-risk-crr"

'------------------------------------------------------------------------------
' LOAD REFERENCE DATA
'------------------------------------------------------------------------------
    'Load the reference data, size the ladder labels and build an FX table with every rate 1
        ReferenceParameters Parameters
        ReferenceLadder Ladder
        ReDim Labels(1 To BAND_COUNT, 1 To 2)
        FxTable "EUR|XAU", FxCodes, FxRates

'------------------------------------------------------------------------------
' BUILD TABLES
'------------------------------------------------------------------------------
    'Size the specific-risk table and enter its six charges
        ReDim SrKeys(1 To 6)
        ReDim SrCharges(1 To 6)
        SrKeys(1) = "CRR|RW 0%|ANY": SrCharges(1) = 0
        SrKeys(2) = "CRR|RW 10%|>24M": SrCharges(2) = 0.008
        SrKeys(3) = "CRR|RW 20%|<=6M": SrCharges(3) = 0.0025
        SrKeys(4) = "CRR|RW 75%|ANY": SrCharges(4) = 0.08
        SrKeys(5) = "CRR|RW 150%|ANY": SrCharges(5) = 0.12
        SrKeys(6) = "GOV|AAA TO AA-|ANY": SrCharges(6) = 0
    'Set the reporting date
        AsOfDate = CDbl(DateSerial(2026, 9, 30))
    'Select the CRR regime
        Parameters(PI_REGIME_FLAG) = 1#

'------------------------------------------------------------------------------
' BUILD POSITIONS
'------------------------------------------------------------------------------
    'Enter eight rows, each of 1,000,000, with the categories and weights to check
        InputRows = RowsFrom(IR_INPUT_WIDTH, Array( _
            Array("R1", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "CRR", "RW 10%", ""), _
            Array("R2", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.05, DateSerial(2026, 12, 31), _
                  Empty, "CRR", "RW 20%", ""), _
            Array("R3", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "CRR", "RW 75%", ""), _
            Array("R4", "", "Bond", "EUR", "Short", 1000000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "crr", "RW 150%", ""), _
            Array("R5", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "GOV", "AAA to AA-", ""), _
            Array("R6", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "CRR", "RW 33%", ""), _
            Array("R7", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "NONE", "", ""), _
            Array("R8", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "CRR", "RW 0%", "")))

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Calculate the interest rate class under the CRR
        CalculateInterestRate InputRows, AsOfDate, Parameters, FxCodes, FxRates, "XAU", SrKeys, _
                              SrCharges, Ladder, Labels, RowOutput, LadderOutput, ZoneOutput, Totals

'------------------------------------------------------------------------------
' CHECK CRR CHARGES
'------------------------------------------------------------------------------
    'Check the charge of each risk weight and range of the residual term
        TEST_Harness.AssertEqualText "weight-10-status", ST_OK, CStr(RowOutput(1, 2))
        TEST_Harness.AssertNear "weight-10-over-24-months", 8000, _
            NumberOf(RowOutput(1, 20)), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "weight-20-up-to-6-months", 2500, _
            NumberOf(RowOutput(2, 20)), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "weight-75-as-100", 80000, _
            NumberOf(RowOutput(3, 20)), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "weight-150-short-lower-case-category", 120000, _
            NumberOf(RowOutput(4, 20)), TOL_ABSOLUTE, TOL_RELATIVE
    'Check the rows that the CRR rejects
        TEST_Harness.AssertEqualText "basel-category-under-the-crr", _
            ST_INVALID & "SR category must be CRR or NONE under the CRR", CStr(RowOutput(5, 2))
        TEST_Harness.AssertEqualText "weight-not-in-the-table", _
            ST_INVALID & "no Article 336 charge for this risk weight", CStr(RowOutput(6, 2))
    'Check a category NONE and a weight of 0%
        TEST_Harness.AssertEqualText "category-none-under-the-crr", ST_OK, CStr(RowOutput(7, 2))
        TEST_Harness.AssertNear "weight-0", 0, NumberOf(RowOutput(8, 20)), TOL_ABSOLUTE, TOL_RELATIVE
    'Check the specific total and the number of rejected rows
        TEST_Harness.AssertNear "specific-total-under-the-crr", 210500, _
            Totals(TOT_SPECIFIC), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertEqualDouble "rows-rejected-under-the-crr", 2, Totals(TOT_REJECTED)

'------------------------------------------------------------------------------
' CHECK BASEL CATEGORIES
'------------------------------------------------------------------------------
    'Select the Basel regime
        Parameters(PI_REGIME_FLAG) = 0#
    'Enter a row of the category CRR and a row of the category GOV
        InputRows = RowsFrom(IR_INPUT_WIDTH, Array( _
            Array("R9", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "CRR", "RW 20%", ""), _
            Array("R10", "", "Bond", "EUR", "Long", 1000000, "Fixed", 0.05, DateSerial(2031, 10, 1), _
                  Empty, "GOV", "AAA to AA-", "")))
    'Calculate again under the Basel regime
        CalculateInterestRate InputRows, AsOfDate, Parameters, FxCodes, FxRates, "XAU", SrKeys, _
                              SrCharges, Ladder, Labels, RowOutput, LadderOutput, ZoneOutput, Totals
    'Check that the Basel regime rejects the category CRR and accepts the category GOV
        TEST_Harness.AssertEqualText "crr-category-under-basel", _
            ST_INVALID & "SR category must be GOV, QUAL, OTHER or NONE", CStr(RowOutput(1, 2))
        TEST_Harness.AssertEqualText "basel-category-under-basel", ST_OK, CStr(RowOutput(2, 2))

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseIssueNetting()

'
'------------------------------------------------------------------------------
'                             CHECK ISSUE NETTING
'------------------------------------------------------------------------------
' PURPOSE
'   Specific-risk netting: rows net only through an entered Issue ID, and a row
'   without one never nets, even when its PositionID equals the Issue ID of
'   another row.
'
' WHY THIS EXISTS
'   Pins the rule that rows net only through an entered Issue ID, whatever the
'   case of the Issue ID.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case issue-netting, enters five rows, calculates the interest
'   rate class and checks the key, the net position and the specific risk of
'   the rows without an Issue ID, of the issue ISS9 entered in two cases, the
'   number of rows flagged MIXED and the totals.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   ReferenceParameters, ReferenceLadder, FxTable, RowsFrom, NumberOf,
'   CalculateInterestRate; TEST_Harness (BeginCase, AssertNear,
'   AssertEqualDouble, AssertEqualText, RecordCaseError)
'
'   Reference: Rows 1 and 2: PositionID R1 without an Issue ID, and Issue ID
'   R1; long and short 1,000,000 at 1.6%: 16,000 each, not a net of zero. Rows
'   3 to 5: Issue ID ISS9 in two cases, net 600,000 - 200,000 + 1,000 =
'   401,000, charged on the first row at its 1.6% = 6,416; row 5 has another SR
'   key, so the three rows are flagged MIXED (MAR40.6).
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InputRows()     As Variant  'IRPositions cells
    Dim FxCodes()       As String   'FX table codes
    Dim FxRates()       As Double   'FX table rates
    Dim Parameters()    As Double   'MAR40 parameters
    Dim Ladder()        As Double   'MAR40 ladder
    Dim Labels()        As String   'Ladder labels
    Dim SrKeys()        As String   'Specific-risk keys
    Dim SrCharges()     As Double   'Specific-risk charges
    Dim RowOutput()     As Variant  'IRCalc
    Dim LadderOutput()  As Variant  'IR_Ladder
    Dim ZoneOutput()    As Variant  'IR_Zones
    Dim Totals()        As Double   'Totals
    Dim AsOfDate        As Double   'Reporting date

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "issue-netting"

'------------------------------------------------------------------------------
' LOAD REFERENCE DATA
'------------------------------------------------------------------------------
    'Load the reference data, size the ladder labels and build an FX table with every rate 1
        ReferenceParameters Parameters
        ReferenceLadder Ladder
        ReDim Labels(1 To BAND_COUNT, 1 To 2)
        FxTable "EUR|XAU", FxCodes, FxRates

'------------------------------------------------------------------------------
' BUILD TABLES
'------------------------------------------------------------------------------
    'Size the specific-risk table and enter its two charges
        ReDim SrKeys(1 To 2)
        ReDim SrCharges(1 To 2)
        SrKeys(1) = "QUAL|A+ TO BBB-|>24M": SrCharges(1) = 0.016
        SrKeys(2) = "GOV|AAA TO AA-|ANY": SrCharges(2) = 0
    'Set the reporting date
        AsOfDate = CDbl(DateSerial(2026, 9, 30))

'------------------------------------------------------------------------------
' BUILD POSITIONS
'------------------------------------------------------------------------------
    'Enter a row whose PositionID equals the Issue ID of the next row, and three rows of issue ISS9
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

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Calculate the interest rate class
        CalculateInterestRate InputRows, AsOfDate, Parameters, FxCodes, FxRates, "XAU", SrKeys, _
                              SrCharges, Ladder, Labels, RowOutput, LadderOutput, ZoneOutput, Totals

'------------------------------------------------------------------------------
' CHECK ROWS WITHOUT AN ISSUE ID
'------------------------------------------------------------------------------
    'Check the keys of the first two rows
        TEST_Harness.AssertEqualText "key-without-an-issue-id", "POS:R1", CStr(RowOutput(1, 17))
        TEST_Harness.AssertEqualText "key-of-an-entered-issue-id", "R1", CStr(RowOutput(2, 17))
    'Check that the two rows do not net: each keeps its net position and its specific risk
        TEST_Harness.AssertNear "net-without-an-issue-id", 1000000, _
            NumberOf(RowOutput(1, 18)), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "net-of-issue-id-equal-to-that-position-id", -1000000, _
            NumberOf(RowOutput(2, 18)), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "specific-without-an-issue-id", 16000, _
            NumberOf(RowOutput(1, 20)), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "specific-of-issue-id-equal-to-that-position-id", 16000, _
            NumberOf(RowOutput(2, 20)), TOL_ABSOLUTE, TOL_RELATIVE

'------------------------------------------------------------------------------
' CHECK ISSUE ISS9
'------------------------------------------------------------------------------
    'Check the key, the net position and the specific risk of the issue, charged on its first row
        TEST_Harness.AssertEqualText "key-ignores-case", "ISS9", CStr(RowOutput(3, 17))
        TEST_Harness.AssertNear "net-of-one-issue-id-in-two-cases", 401000, _
            NumberOf(RowOutput(3, 18)), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "specific-on-the-first-row-of-the-issue", 6416, _
            NumberOf(RowOutput(3, 20)), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertNear "no-specific-on-a-later-row-of-the-issue", 0, _
            NumberOf(RowOutput(4, 20)), TOL_ABSOLUTE, TOL_RELATIVE
    'Check the number of rows of an issue that mixes keys
        TEST_Harness.AssertEqualDouble "rows-of-an-issue-that-mixes-keys", 3, Totals(TOT_MIXED)

'------------------------------------------------------------------------------
' CHECK TOTALS
'------------------------------------------------------------------------------
    'Check the specific total and the number of used rows
        TEST_Harness.AssertNear "specific-total", 38416, _
            Totals(TOT_SPECIFIC), TOL_ABSOLUTE, TOL_RELATIVE
        TEST_Harness.AssertEqualDouble "rows-used", 5, Totals(TOT_USED)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseDateBounds()

'
'------------------------------------------------------------------------------
'                              CHECK DATE BOUNDS
'------------------------------------------------------------------------------
' PURPOSE
'   A date typed as a number is not a date, on a row and as the reporting date;
'   the last date Excel holds is one.
'
' WHY THIS EXISTS
'   A date typed as a number such as 20311231 must be rejected and not read as
'   a serial, while the last date Excel holds must be accepted.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case date-bounds, enters four rows with dates typed as numbers,
'   the last date and a valid date, calculates the interest rate class and
'   checks their status; then checks BuildParameters on a reporting date typed
'   as a number, on zero and on the last date.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   ReferenceParameters, ReferenceLadder, FxTable, RowsFrom,
'   CalculateInterestRate, BuildParameters; TEST_Harness (BeginCase,
'   AssertEqualDouble, AssertEqualText, RecordCaseError); DATE_SERIAL_MAX
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InputRows()     As Variant  'IRPositions cells
    Dim FxCodes()       As String   'FX table codes
    Dim FxRates()       As Double   'FX table rates
    Dim Parameters()    As Double   'MAR40 parameters
    Dim Ladder()        As Double   'MAR40 ladder
    Dim Labels()        As String   'Ladder labels
    Dim SrKeys()        As String   'Specific-risk keys
    Dim SrCharges()     As Double   'Specific-risk charges
    Dim RowOutput()     As Variant  'IRCalc
    Dim LadderOutput()  As Variant  'IR_Ladder
    Dim ZoneOutput()    As Variant  'IR_Zones
    Dim Totals()        As Double   'Totals
    Dim AsOfDate        As Double   'Reporting date
    Dim Raw()           As Variant  'Raw numeric parameters
    Dim Validated()     As Double   'Validated parameters
    Dim Findings()      As String   'Findings
    Dim FindingCount    As Long     'Findings in use
    Dim ReportingCcy    As String   'Reporting currency
    Dim GoldCode        As String   'Gold code
    Dim i               As Long     'Parameter

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "date-bounds"

'------------------------------------------------------------------------------
' LOAD REFERENCE DATA
'------------------------------------------------------------------------------
    'Load the reference data, size the ladder labels and build an FX table with every rate 1
        ReferenceParameters Parameters
        ReferenceLadder Ladder
        ReDim Labels(1 To BAND_COUNT, 1 To 2)
        FxTable "EUR|XAU", FxCodes, FxRates

'------------------------------------------------------------------------------
' BUILD TABLES
'------------------------------------------------------------------------------
    'Size the specific-risk table and enter its charge
        ReDim SrKeys(1 To 1)
        ReDim SrCharges(1 To 1)
        SrKeys(1) = "GOV|AAA TO AA-|ANY": SrCharges(1) = 0
    'Set the reporting date
        AsOfDate = CDbl(DateSerial(2026, 9, 30))

'------------------------------------------------------------------------------
' BUILD POSITIONS
'------------------------------------------------------------------------------
    'Enter a maturity and a repricing date typed as numbers, the last date and a valid row
        InputRows = RowsFrom(IR_INPUT_WIDTH, Array( _
            Array("R1", "", "Bond", "EUR", "Long", 1000, "Fixed", 0.04, 20311231, _
                  Empty, "GOV", "AAA to AA-", ""), _
            Array("R2", "", "FRN", "EUR", "Long", 1000, "Floating", Empty, DateSerial(2031, 10, 1), _
                  20290101, "GOV", "AAA to AA-", ""), _
            Array("R3", "", "Bond", "EUR", "Long", 1000, "Fixed", 0.04, DATE_SERIAL_MAX, _
                  Empty, "GOV", "AAA to AA-", ""), _
            Array("R4", "", "Bond", "EUR", "Long", 1000, "Fixed", 0.04, DateSerial(2031, 10, 1), _
                  Empty, "GOV", "AAA to AA-", "")))

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Calculate the interest rate class
        CalculateInterestRate InputRows, AsOfDate, Parameters, FxCodes, FxRates, "XAU", SrKeys, _
                              SrCharges, Ladder, Labels, RowOutput, LadderOutput, ZoneOutput, Totals

'------------------------------------------------------------------------------
' CHECK ROW DATES
'------------------------------------------------------------------------------
    'Check the status of each row and the counts of used and rejected rows
        TEST_Harness.AssertEqualText "maturity-typed-as-a-number", _
            ST_INVALID & "Final maturity date is not a date", CStr(RowOutput(1, 2))
        TEST_Harness.AssertEqualText "repricing-date-typed-as-a-number", _
            ST_INVALID & "Next repricing date is not a date", CStr(RowOutput(2, 2))
        TEST_Harness.AssertEqualText "maturity-on-the-last-date", ST_OK, CStr(RowOutput(3, 2))
        TEST_Harness.AssertEqualText "row-after-a-rejected-date", ST_OK, CStr(RowOutput(4, 2))
        TEST_Harness.AssertEqualDouble "rows-used", 2, Totals(TOT_USED)
        TEST_Harness.AssertEqualDouble "rows-rejected", 2, Totals(TOT_REJECTED)

'------------------------------------------------------------------------------
' BUILD RAW PARAMETERS
'------------------------------------------------------------------------------
    'Copy the reference parameters into the raw vector
        ReDim Raw(1 To PI_COUNT)
        For i = 1 To PI_COUNT
            Raw(i) = Parameters(i)
        Next i

'------------------------------------------------------------------------------
' CHECK REPORTING DATE
'------------------------------------------------------------------------------
    'Check a reporting date typed as a number and a reporting date of zero
        TEST_Harness.AssertEqualDouble "reporting-date-typed-as-a-number", 1, _
            BuildParameters(Raw, 20260930, "EUR", "XAU", Validated, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount)
        TEST_Harness.AssertEqualDouble "reporting-date-of-zero", 1, _
            BuildParameters(Raw, 0, "EUR", "XAU", Validated, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount)
    'Check the last date as the reporting date
        TEST_Harness.AssertEqualDouble "reporting-date-on-the-last-date", 0, _
            BuildParameters(Raw, DATE_SERIAL_MAX, "EUR", "XAU", Validated, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseParameters()

'
'------------------------------------------------------------------------------
'                               CHECK PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   The reference parameters are valid and unchanged; an unusable value and a
'   faulty FX table are counted.
'
' WHY THIS EXISTS
'   Pins the reference parameter set, so that an edit of a value shows, and
'   checks the counting of unusable values and of the faults of an FX table.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case parameters, builds the validated parameters from the
'   reference set and checks that they are valid and unchanged, changes three
'   raw values and checks the count of unusable ones, then builds an FX table
'   with a reporting currency not at 1 and a duplicate code and checks the
'   counts of rules C05, C06 and C07.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   ReferenceParameters, BuildParameters, ChangedParameterCount, BuildFxTable;
'   TEST_Harness (BeginCase, AssertEqualDouble, RecordCaseError);
'   PRM_REFERENCE_END_DATE
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Reference()     As Double   'MAR40 parameters
    Dim Raw()           As Variant  'Raw numeric parameters
    Dim RawFx()         As Variant  'Raw FX table
    Dim Parameters()    As Double   'Validated parameters
    Dim FxCodes()       As String   'FX codes
    Dim FxRates()       As Double   'FX rates
    Dim Findings()      As String   'Findings
    Dim FindingCount    As Long     'Findings in use
    Dim AsOfDate        As Double   'Reporting date
    Dim ReportingCcy    As String   'Reporting currency
    Dim GoldCode        As String   'Gold code
    Dim ReportingErrors As Long     'C05
    Dim DuplicateErrors As Long     'C06
    Dim RateErrors      As Long     'C07
    Dim i               As Long     'Parameter

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "parameters"

'------------------------------------------------------------------------------
' BUILD RAW PARAMETERS
'------------------------------------------------------------------------------
    'Load the reference parameters and copy them into the raw vector
        ReferenceParameters Reference
        ReDim Raw(1 To PI_COUNT)
        For i = 1 To PI_COUNT
            Raw(i) = Reference(i)
        Next i

'------------------------------------------------------------------------------
' CHECK REFERENCE SET
'------------------------------------------------------------------------------
    'Check that the reference set is valid and unchanged
        TEST_Harness.AssertEqualDouble "reference-set-is-valid", 0, _
            BuildParameters(Raw, DateSerial(2026, 9, 30), "EUR", "XAU", Parameters, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount)
        TEST_Harness.AssertEqualDouble "reference-set-unchanged", 0, _
            ChangedParameterCount(Parameters, GoldCode, PRM_REFERENCE_END_DATE)

'------------------------------------------------------------------------------
' CHECK UNUSABLE VALUES
'------------------------------------------------------------------------------
    'Change a day count, a vertical weight and a zone weight
        Raw(PI_DAYS_PER_YEAR) = 360
        Raw(PI_VERTICAL) = "ten"
        Raw(PI_ZONE1) = 1.5
    'Check the number of unusable values
        TEST_Harness.AssertEqualDouble "two-unusable-values", 2, _
            BuildParameters(Raw, DateSerial(2026, 9, 30), "EUR", "XAU", Parameters, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount)

'------------------------------------------------------------------------------
' CHECK FX TABLE
'------------------------------------------------------------------------------
    'Enter an FX table with a duplicate code and a reporting currency not at 1
        ReDim RawFx(1 To 3, 1 To 2)
        RawFx(1, 1) = "USD": RawFx(1, 2) = 0.9
        RawFx(2, 1) = "usd": RawFx(2, 2) = 0.9
        RawFx(3, 1) = "EUR": RawFx(3, 2) = 1.1
    'Validate the FX table
        BuildFxTable RawFx, "EUR", FxCodes, FxRates, ReportingErrors, DuplicateErrors, RateErrors, _
                     Findings, FindingCount
    'Check the number of errors of each kind
        TEST_Harness.AssertEqualDouble "reporting-currency-not-at-1", 1, ReportingErrors
        TEST_Harness.AssertEqualDouble "duplicate-code", 1, DuplicateErrors
        TEST_Harness.AssertEqualDouble "no-rate-error", 0, RateErrors

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseRegime()

'
'------------------------------------------------------------------------------
'                                 CHECK REGIME
'------------------------------------------------------------------------------
' PURPOSE
'   The regime of the workbook: Basel or CRR without regard to case, the last
'   date and the total own funds, read whatever the regime and required under
'   the CRR; the transitional factor on and after its last date; the multiplier
'   as a rate; the last date counted as a what-if value; and the de minimis of
'   Article 351 on, above and below its threshold.
'
' WHY THIS EXISTS
'   The regime selects the rules of the CRR, so its reading, the transitional
'   factor and the de minimis are checked on both sides of each boundary.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case regime, checks BuildRegime on valid and on invalid input,
'   checks TransitionalFactor and DeMinimisWaived on and around their
'   boundaries, and checks the count of changed parameters and the validation
'   of the multiplier on the reference parameters.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   BuildRegime, TransitionalFactor, DeMinimisWaived, ReferenceParameters,
'   BuildParameters, ChangedParameterCount; TEST_Harness (BeginCase,
'   AssertEqualDouble, AssertEqualText, AssertTrue, RecordCaseError);
'   REGIME_BASEL, REGIME_CRR
'
'   Reference: Article 495s(2) of Regulation (EU) No 575/2013, from Delegated
'   Regulation (EU) 2026/1221: 0.9 until 2029-12-31. Article 351: no FX
'   requirement unless the overall net open position exceeds 2% of total own
'   funds.
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
    Dim Reference()     As Double   'MAR40 parameters
    Dim Raw()           As Variant  'Raw numeric parameters
    Dim Parameters()    As Double   'Validated parameters
    Dim Findings()      As String   'Findings
    Dim FindingCount    As Long     'Findings in use
    Dim Regime          As String   'REGIME_BASEL or REGIME_CRR
    Dim EndDate         As Double   'Last date of the multiplier
    Dim OwnFunds        As Double   'Total own funds
    Dim AsOfDate        As Double   'Reporting date
    Dim ReportingCcy    As String   'Reporting currency
    Dim GoldCode        As String   'Gold code
    Dim LastDate        As Double   'Last date of the multiplier, 2029-12-31
    Dim i               As Long     'Parameter

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "regime"

'------------------------------------------------------------------------------
' SET LAST DATE
'------------------------------------------------------------------------------
    'Fix the last date of the multiplier, 2029-12-31
        LastDate = CDbl(DateSerial(2029, 12, 31))

'------------------------------------------------------------------------------
' CHECK VALID REGIMES
'------------------------------------------------------------------------------
    'Basel is valid and is read in upper case; without an end date or own funds both are zero
        TEST_Harness.AssertEqualDouble "basel-is-valid", 0, _
            BuildRegime("Basel", Empty, Empty, Regime, EndDate, OwnFunds, Findings, FindingCount)
        TEST_Harness.AssertEqualText "basel-is-read-in-upper-case", REGIME_BASEL, Regime
        TEST_Harness.AssertEqualDouble "no-end-date-is-zero", 0, EndDate
        TEST_Harness.AssertEqualDouble "no-own-funds-is-zero", 0, OwnFunds
    'Basel reads an end date that is a date, and the own funds
        TEST_Harness.AssertEqualDouble "basel-reads-an-end-date-that-is-a-date", 0, _
            BuildRegime("Basel", DateSerial(2030, 12, 31), 5000000, Regime, EndDate, OwnFunds, _
                        Findings, FindingCount)
        TEST_Harness.AssertEqualDouble "the-end-date-is-read-under-basel", _
            CDbl(DateSerial(2030, 12, 31)), EndDate
        TEST_Harness.AssertEqualDouble "the-own-funds-are-read-under-basel", 5000000, OwnFunds
    'CRR is read without regard to case, in upper case, with its end date
        TEST_Harness.AssertEqualDouble "crr-without-regard-to-case", 0, _
            BuildRegime(" crr ", DateSerial(2029, 12, 31), 5000000, Regime, EndDate, OwnFunds, _
                        Findings, FindingCount)
        TEST_Harness.AssertEqualText "crr-is-read-in-upper-case", REGIME_CRR, Regime
        TEST_Harness.AssertEqualDouble "crr-reads-the-end-date", LastDate, EndDate
    'Basel ignores an end date and own funds that are not dates or numbers
        TEST_Harness.AssertEqualDouble "basel-ignores-an-end-date-that-is-no-date", 0, _
            BuildRegime("Basel", "never", "many", Regime, EndDate, OwnFunds, Findings, FindingCount)

'------------------------------------------------------------------------------
' CHECK INVALID REGIMES
'------------------------------------------------------------------------------
    'An unknown regime and a blank regime are errors
        TEST_Harness.AssertEqualDouble "an-unknown-regime", 1, _
            BuildRegime("EU", Empty, Empty, Regime, EndDate, OwnFunds, Findings, FindingCount)
        TEST_Harness.AssertEqualDouble "a-blank-regime", 1, _
            BuildRegime("", DateSerial(2029, 12, 31), 5000000, Regime, EndDate, OwnFunds, _
                        Findings, FindingCount)
    'CRR with an end date typed as text or missing is an error
        TEST_Harness.AssertEqualDouble "crr-with-an-end-date-typed-as-text", 1, _
            BuildRegime("CRR", "31.12.2029", 5000000, Regime, EndDate, OwnFunds, Findings, FindingCount)
        TEST_Harness.AssertEqualDouble "crr-with-no-end-date", 1, _
            BuildRegime("CRR", Empty, 5000000, Regime, EndDate, OwnFunds, Findings, FindingCount)
    'CRR with own funds missing, zero or typed as text is an error
        TEST_Harness.AssertEqualDouble "crr-with-no-own-funds", 1, _
            BuildRegime("CRR", DateSerial(2029, 12, 31), Empty, Regime, EndDate, OwnFunds, _
                        Findings, FindingCount)
        TEST_Harness.AssertEqualDouble "crr-with-own-funds-of-zero", 1, _
            BuildRegime("CRR", DateSerial(2029, 12, 31), 0, Regime, EndDate, OwnFunds, _
                        Findings, FindingCount)
        TEST_Harness.AssertEqualDouble "crr-with-own-funds-typed-as-text", 1, _
            BuildRegime("CRR", DateSerial(2029, 12, 31), "five billion", Regime, EndDate, OwnFunds, _
                        Findings, FindingCount)

'------------------------------------------------------------------------------
' CHECK TRANSITIONAL FACTOR
'------------------------------------------------------------------------------
    'Check the factor on the last date, the day after it and under Basel
        TEST_Harness.AssertEqualDouble "factor-on-the-last-date", 0.9, _
            TransitionalFactor(REGIME_CRR, LastDate, LastDate, 0.9)
        TEST_Harness.AssertEqualDouble "factor-the-day-after", 1, _
            TransitionalFactor(REGIME_CRR, LastDate + 1, LastDate, 0.9)
        TEST_Harness.AssertEqualDouble "factor-under-basel", 1, _
            TransitionalFactor(REGIME_BASEL, CDbl(DateSerial(2026, 9, 30)), LastDate, 0.9)

'------------------------------------------------------------------------------
' CHECK DE MINIMIS
'------------------------------------------------------------------------------
    'Check the threshold of 2% of 100 million on it, below it and above it, and under Basel
        TEST_Harness.AssertTrue "de-minimis-on-the-threshold", _
            DeMinimisWaived(REGIME_CRR, 2000000#, 100000000#, 0.02), _
            "2% of 100 million is 2 million: the requirement is waived on it"
        TEST_Harness.AssertTrue "de-minimis-below-the-threshold", _
            DeMinimisWaived(REGIME_CRR, 1999999#, 100000000#, 0.02), _
            "the requirement is waived below the threshold"
        TEST_Harness.AssertTrue "no-de-minimis-above-the-threshold", _
            Not DeMinimisWaived(REGIME_CRR, 2000001#, 100000000#, 0.02), _
            "the requirement applies above the threshold"
        TEST_Harness.AssertTrue "no-de-minimis-under-basel", _
            Not DeMinimisWaived(REGIME_BASEL, 1#, 100000000#, 0.02), _
            "the Basel regime never waives"

'------------------------------------------------------------------------------
' LOAD REFERENCE PARAMETERS
'------------------------------------------------------------------------------
    'Reload the reference parameters into the raw vector
        ReferenceParameters Reference
        ReDim Raw(1 To PI_COUNT)
        For i = 1 To PI_COUNT
            Raw(i) = Reference(i)
        Next i

'------------------------------------------------------------------------------
' CHECK REFERENCE VALUES
'------------------------------------------------------------------------------
    'Check the multiplier and the de minimis rate of the reference set
        TEST_Harness.AssertEqualDouble "multiplier-of-reference", 0.9, Reference(PI_TRANSITIONAL)
        TEST_Harness.AssertEqualDouble "de-minimis-rate-of-reference", 0.02, Reference(PI_FX_DE_MINIMIS)
    'Validate the reference set
        BuildParameters Raw, DateSerial(2026, 9, 30), "EUR", "XAU", Parameters, AsOfDate, _
                        ReportingCcy, GoldCode, Findings, FindingCount
    'Check that the reference end date is unchanged, and that a later date and a missing date are counted
        TEST_Harness.AssertEqualDouble "reference-end-date-is-unchanged", 0, _
            ChangedParameterCount(Parameters, GoldCode, LastDate)
        TEST_Harness.AssertEqualDouble "a-later-end-date-is-counted", 1, _
            ChangedParameterCount(Parameters, GoldCode, LastDate + 366)
        TEST_Harness.AssertEqualDouble "an-end-date-that-is-no-date-is-counted", 1, _
            ChangedParameterCount(Parameters, GoldCode, 0#)

'------------------------------------------------------------------------------
' CHECK MULTIPLIER ABOVE ONE
'------------------------------------------------------------------------------
    'A multiplier above one is an error
        Raw(PI_TRANSITIONAL) = 1.5
        TEST_Harness.AssertEqualDouble "multiplier-above-one", 1, _
            BuildParameters(Raw, DateSerial(2026, 9, 30), "EUR", "XAU", Parameters, AsOfDate, _
                            ReportingCcy, GoldCode, Findings, FindingCount)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseVolume()

'
'------------------------------------------------------------------------------
'                            CHECK POSITION VOLUME
'------------------------------------------------------------------------------
' PURPOSE
'   The largest volume the workbook supports: a position table of exactly
'   MAX_POSITION_ROWS rows is within the limit, and one row more is a finding
'   that names the sheet and the rows found.
'
' WHY THIS EXISTS
'   The workbook supports at most MAX_POSITION_ROWS rows on a position sheet,
'   so the limit and the row beyond it are checked.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case volume, fills the four position tables to the limit and
'   checks that no sheet is over it, adds one row to the interest rate table
'   and checks the finding, then adds one row to the other three tables and
'   checks that all four sheets are over the limit.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   FilledRows, PositionVolumeErrors; TEST_Harness (BeginCase,
'   AssertEqualDouble, AssertEqualText, RecordCaseError); MAX_POSITION_ROWS
'
'   Reference: The data contract's Volume section: at most 500 rows on a
'   position sheet.
'
' NOTES
'   None
'
' UPDATED
'   2026-11-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Inputs(1 To IN_COUNT)   As Variant  'Engine input vector, position tables only
    Dim Findings()              As String   'Findings
    Dim FindingCount            As Long     'Findings in use

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "volume"

'------------------------------------------------------------------------------
' BUILD TABLES AT THE LIMIT
'------------------------------------------------------------------------------
    'Fill each position table with MAX_POSITION_ROWS rows
        Inputs(IN_IR_ROWS) = FilledRows(IR_INPUT_WIDTH, MAX_POSITION_ROWS)
        Inputs(IN_EQ_ROWS) = FilledRows(EQ_INPUT_WIDTH, MAX_POSITION_ROWS)
        Inputs(IN_CO_ROWS) = FilledRows(CO_INPUT_WIDTH, MAX_POSITION_ROWS)
        Inputs(IN_FX_ROWS) = FilledRows(FX_INPUT_WIDTH, MAX_POSITION_ROWS)

'------------------------------------------------------------------------------
' CHECK THE LIMIT
'------------------------------------------------------------------------------
    'Check that no sheet is beyond the limit
        TEST_Harness.AssertEqualDouble "sheets-at-the-limit", 0, _
            PositionVolumeErrors(Inputs, Findings, FindingCount)

'------------------------------------------------------------------------------
' CHECK ONE ROW BEYOND
'------------------------------------------------------------------------------
    'Add one row to the interest rate table and check the count and the finding
        Inputs(IN_IR_ROWS) = FilledRows(IR_INPUT_WIDTH, MAX_POSITION_ROWS + 1)
        TEST_Harness.AssertEqualDouble "sheets-one-row-beyond", 1, _
            PositionVolumeErrors(Inputs, Findings, FindingCount)
        TEST_Harness.AssertEqualDouble "findings-recorded", 1, FindingCount
        TEST_Harness.AssertEqualText "finding-sheet", SH_IR_POS, Findings(1, 1)
        TEST_Harness.AssertEqualText "finding-reason", _
            ST_INVALID & "501 rows found; the workbook supports at most 500", Findings(4, 1)

'------------------------------------------------------------------------------
' CHECK ALL SHEETS BEYOND
'------------------------------------------------------------------------------
    'Add one row to the other three tables
        Inputs(IN_EQ_ROWS) = FilledRows(EQ_INPUT_WIDTH, MAX_POSITION_ROWS + 1)
        Inputs(IN_CO_ROWS) = FilledRows(CO_INPUT_WIDTH, MAX_POSITION_ROWS + 1)
        Inputs(IN_FX_ROWS) = FilledRows(FX_INPUT_WIDTH, MAX_POSITION_ROWS + 1)
    'Count the findings again and check that all four sheets are beyond the limit
        FindingCount = 0
        TEST_Harness.AssertEqualDouble "all-four-sheets-beyond", 4, _
            PositionVolumeErrors(Inputs, Findings, FindingCount)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseFacade()

'
'------------------------------------------------------------------------------
'                          CHECK WORKSHEET FUNCTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   The worksheet functions on the inputs of the cases above, and the error
'   value each gives for an input outside its domain.
'
' WHY THIS EXISTS
'   The worksheet functions are the facade of the core, so they are checked on
'   inputs whose result is known from the cases above.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case facade, fills the weighted longs and shorts of ladder A and
'   checks the time band, the ladder charge, the FX charge and the commodity
'   charge, and that a negative maturity and a gross below the net give an
'   error value.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   FRTBSSA_TimeBand, FRTBSSA_LadderCharge, FRTBSSA_FXCharge,
'   FRTBSSA_CommodityCharge, NumberOf; TEST_Harness (BeginCase,
'   AssertEqualDouble, AssertNear, AssertTrue, RecordCaseError)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Longs(1 To 15)  As Variant  'Weighted longs, ladder A
    Dim Shorts(1 To 15) As Variant  'Weighted shorts, ladder A
    Dim b               As Long     'Band

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "facade"

'------------------------------------------------------------------------------
' BUILD LADDER VECTORS
'------------------------------------------------------------------------------
    'Zero every band, then weight ladder A
        For b = 1 To BAND_COUNT
            Longs(b) = 0
            Shorts(b) = 0
        Next b
        Longs(4) = 100: Shorts(4) = 60: Shorts(6) = 50: Longs(10) = 80

'------------------------------------------------------------------------------
' CHECK TIME BANDS
'------------------------------------------------------------------------------
    'Check the band of a maturity and a coupon, and of a floating-rate position
        TEST_Harness.AssertEqualDouble "time-band-six-months-at-5-percent", 3, _
            NumberOf(FRTBSSA_TimeBand(0.5, 0.05))
        TEST_Harness.AssertEqualDouble "time-band-25-years-at-1-percent", 15, _
            NumberOf(FRTBSSA_TimeBand(25, 0.01))
        TEST_Harness.AssertEqualDouble "time-band-floating-2-and-a-half-years", 6, _
            NumberOf(FRTBSSA_TimeBand(2.5, Empty, True))
    'Check that a negative maturity gives an error value
        TEST_Harness.AssertTrue "time-band-negative-maturity", IsError(FRTBSSA_TimeBand(-1, 0.05)), _
            "a negative maturity must give an error value"

'------------------------------------------------------------------------------
' CHECK CHARGES
'------------------------------------------------------------------------------
    'Check the ladder charge of ladder A
        TEST_Harness.AssertNear "ladder-charge-of-ladder-a", 96, _
            NumberOf(FRTBSSA_LadderCharge(Longs, Shorts)), TOL_ABSOLUTE, TOL_RELATIVE
    'Check the FX charge of table 9
        TEST_Harness.AssertNear "fx-charge-of-table-9", 26.8, _
            NumberOf(FRTBSSA_FXCharge(300, 200, -35)), TOL_ABSOLUTE, TOL_RELATIVE
    'Check the commodity charge
        TEST_Harness.AssertNear "commodity-charge", 14.4, _
            NumberOf(FRTBSSA_CommodityCharge(70, 130)), TOL_ABSOLUTE, TOL_RELATIVE
    'Check that a gross below the net gives an error value
        TEST_Harness.AssertTrue "commodity-charge-gross-below-net", _
            IsError(FRTBSSA_CommodityCharge(70, 10)), "a gross below the net must give an error value"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseFingerprint()

'
'------------------------------------------------------------------------------
'                              CHECK FINGERPRINT
'------------------------------------------------------------------------------
' PURPOSE
'   The fingerprint of a vector is the one its rule gives, and a changed value,
'   or two values that change places, give another one.
'
' WHY THIS EXISTS
'   The fingerprint identifies a vector of inputs, and its expected values were
'   calculated apart from the VBA code, so a change of its rule shows.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Begins the case fingerprint, builds a vector of scalars, a table and a
'   column, and checks the fingerprint of the vector, the same fingerprint
'   again, the fingerprint after a text changes its case and after two values
'   change places.
'
' ERROR POLICY
'   An unexpected error inside the case is recorded as a failure of the case
'   with TEST_Harness.RecordCaseError, and the run continues.
'
' DEPENDENCIES
'   Fingerprint; TEST_Harness (BeginCase, AssertEqualText, RecordCaseError)
'
'   Reference: Each value enters as its key and code 30: "N" and Str$ of a
'   number, "S", the length, ":" and the text, "_" for an empty value, "B1" for
'   True; an array as "A" and its rows, then its values by row. Two hashes
'   start at 131 and 137 and take each code unit as h = (h * base + code)
'   modulo 2^31 - 1 (base 131) and 2^31 - 19 (base 137).
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Values(1 To 6)          As Variant  'Vector: scalars, a table and a column
    Dim Table(1 To 2, 1 To 2)   As Variant  'Table; its last cell stays empty
    Dim Column(1 To 3)          As Variant  'Array of one dimension

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route an error of the case to its handler and begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "fingerprint"

'------------------------------------------------------------------------------
' BUILD VECTOR
'------------------------------------------------------------------------------
    'Fill a vector with scalars, a table and a column; the third item and the last table cell stay empty
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

'------------------------------------------------------------------------------
' CHECK KNOWN VECTOR
'------------------------------------------------------------------------------
    'Check the fingerprint the rule gives, and that the same vector gives it again
        TEST_Harness.AssertEqualText "known-vector", "F1-68BD702560607F96", Fingerprint(Values)
        TEST_Harness.AssertEqualText "same-vector-again", "F1-68BD702560607F96", Fingerprint(Values)

'------------------------------------------------------------------------------
' CHECK CHANGED TEXT
'------------------------------------------------------------------------------
    'Change a text to another case and check that the fingerprint changes
        Column(2) = "B"
        Values(6) = Column
        TEST_Harness.AssertEqualText "text-in-another-case", "F1-7833FF864C712519", _
            Fingerprint(Values)

'------------------------------------------------------------------------------
' CHECK SWAPPED VALUES
'------------------------------------------------------------------------------
    'Restore the text
        Column(2) = "b"
        Values(6) = Column
    'Swap two values of the table
        Table(1, 2) = -2
        Table(2, 1) = "a"
        Values(5) = Table
    'Check that two values that change places give another fingerprint
        TEST_Harness.AssertEqualText "two-values-change-places", "F1-4BF78E8A19460933", _
            Fingerprint(Values)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Leave before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the error as a failure of the case
CaseFailed:
        TEST_Harness.RecordCaseError

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function NumberOf( _
    ByVal Value As Variant) _
    As Double
'
'------------------------------------------------------------------------------
'                           CONVERT RESULT TO NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   The number a result holds, for an assertion that compares numbers.
'
' WHY THIS EXISTS
'   An assertion must never compare an empty cell, a text or an error value as
'   0, so a result that is not a number is refused.
'
' INPUTS
'   Value
'     A cell of an output table, or what a worksheet function returned.
'
' RETURNS
'   The value as a Double.
'
' BEHAVIOR
'   Refuses a value that IsNumberValue does not accept, and converts any other
'   with CDbl.
'
' ERROR POLICY
'   Raises error 13 when Value is not a number, so that an empty cell, a text
'   or an error value is never compared as 0; the case records it.
'
' DEPENDENCIES
'   IsNumberValue
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE VALUE
'------------------------------------------------------------------------------
    'Refuse a value that is not a number
        If Not IsNumberValue(Value) Then
            Err.Raise 13, "TEST_Suite.NumberOf", "The result is not a number."
        End If

'------------------------------------------------------------------------------
' RETURN NUMBER
'------------------------------------------------------------------------------
    'Return the value as a Double
        NumberOf = CDbl(Value)

End Function

Private Function FilledRows( _
    ByVal Width As Long, _
    ByVal Count As Long) _
    As Variant
'
'------------------------------------------------------------------------------
'                              BUILD FILLED TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Build an input table of Count rows, each with a PositionID in its first
'   column, as Range.Value would return it.
'
' WHY THIS EXISTS
'   The volume case needs tables of exactly the supported size and one row
'   more, and only the number of rows matters.
'
' INPUTS
'   Width
'     Columns of the table.
'
'   Count
'     Rows of the table (>= 1).
'
' RETURNS
'   Variant (1 To Count, 1 To Width).
'
' BEHAVIOR
'   Sizes the table and writes P1, P2 and so on in its first column; every
'   other cell stays Empty.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Table() As Variant  'Result
    Dim r       As Long     'Row

'------------------------------------------------------------------------------
' SIZE TABLE
'------------------------------------------------------------------------------
    'Size the table
        ReDim Table(1 To Count, 1 To Width)

'------------------------------------------------------------------------------
' FILL POSITION IDS
'------------------------------------------------------------------------------
    'Write a PositionID in the first column of each row
        For r = 1 To Count
            Table(r, 1) = "P" & r
        Next r

'------------------------------------------------------------------------------
' RETURN TABLE
'------------------------------------------------------------------------------
    'Return the table
        FilledRows = Table

End Function

Private Function RowsFrom( _
    ByVal Width As Long, _
    ByVal Records As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                           BUILD TABLE FROM RECORDS
'------------------------------------------------------------------------------
' PURPOSE
'   Build an input table, as Range.Value would return it, from an array of
'   record arrays; missing trailing fields are Empty.
'
' WHY THIS EXISTS
'   A case writes its positions as records, one Array(...) per row, and the
'   calculation reads a table of cells.
'
' INPUTS
'   Width
'     Columns of the table.
'
'   Records
'     Array of Array(...) records, one per row.
'
' RETURNS
'   Variant (1 To n, 1 To Width).
'
' BEHAVIOR
'   Sizes the table to the number of records and the width, and copies the
'   fields of each record into its row, up to the width.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Table() As Variant  'Result
    Dim Record  As Variant  'One record
    Dim r       As Long     'Row
    Dim c       As Long     'Field

'------------------------------------------------------------------------------
' SIZE TABLE
'------------------------------------------------------------------------------
    'Size the table to the number of records and the width
        ReDim Table(1 To UBound(Records) - LBound(Records) + 1, 1 To Width)

'------------------------------------------------------------------------------
' COPY RECORDS
'------------------------------------------------------------------------------
    'Copy the fields of each record into its row; a missing trailing field stays Empty
        For r = 1 To UBound(Table, 1)
            Record = Records(LBound(Records) + r - 1)
            For c = 1 To Width
                If c <= UBound(Record) - LBound(Record) + 1 Then
                    Table(r, c) = Record(LBound(Record) + c - 1)
                End If
            Next c
        Next r

'------------------------------------------------------------------------------
' RETURN TABLE
'------------------------------------------------------------------------------
    'Return the table
        RowsFrom = Table

End Function

Private Sub FxTable( _
    ByVal CodeList As String, _
    ByRef Codes() As String, _
    ByRef Rates() As Double)
'
'------------------------------------------------------------------------------
'                                BUILD FX TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Build an FX table with every rate 1.
'
' WHY THIS EXISTS
'   The cases need an FX table in which every currency converts at 1, so that
'   an amount is the same in the reporting currency.
'
' INPUTS
'   CodeList
'     Currency codes separated by "|".
'
'   Codes
'     Receives the currency codes, base 1.
'
'   Rates
'     Receives the rates, base 1; every rate is 1.
'
' RETURNS
'   Codes and Rates, base 1.
'
' BEHAVIOR
'   Splits the list at "|", sizes both arrays to the number of codes, and sets
'   each code and a rate of 1.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Items() As String   'Codes
    Dim i       As Long     'Code

'------------------------------------------------------------------------------
' SPLIT CODES
'------------------------------------------------------------------------------
    'Split the list of codes
        Items = Split(CodeList, "|")

'------------------------------------------------------------------------------
' SIZE TABLES
'------------------------------------------------------------------------------
    'Size the code and rate arrays to the number of codes
        ReDim Codes(1 To UBound(Items) + 1)
        ReDim Rates(1 To UBound(Items) + 1)

'------------------------------------------------------------------------------
' FILL TABLE
'------------------------------------------------------------------------------
    'Set each code with a rate of 1
        For i = 1 To UBound(Items) + 1
            Codes(i) = Items(i - 1)
            Rates(i) = 1#
        Next i

End Sub
