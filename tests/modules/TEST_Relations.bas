Attribute VB_Name = "TEST_Relations"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Relations
'------------------------------------------------------------------------------
' PURPOSE
'   Illustrative relation tests: checks that need no external expected
'   value, run on a stand-in for the calculation under test, the total of a
'   list of amounts. Each case asserts that its relation holds, then that
'   the same comparison rejects a deliberately corrupted output, so that a
'   relation which would pass any output is seen. Nothing here is a
'   calculation of the workbook or a regulatory value
'
' WHY THIS EXISTS
'   Shows the pattern of a relation test for a calculation that has no
'   external expected value: the relation is asserted, and the same
'   comparison is then shown to reject a corrupted output, so that a
'   comparison which could never fail is seen. The pattern is described in
'   docs/methodology/TEST_HARNESS.md (Relation tests)
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunRelationTests runs four cases in a fixed order between BeginRun and
'   EndRun of TEST_Harness:
'     - CaseZeroInput: zero inputs give a zero total
'     - CaseAdditivity: the total of two lists together equals the sum of
'       their totals
'     - CaseSymmetry: the order of the inputs does not change the total
'     - CaseMonotonicity: raising one input never lowers the total
'
'   Each case runs its relation on IllustrativeTotal, the stand-in for the
'   calculation under test, and then runs the same comparison on an output
'   shifted by CORRUPTION, which must be rejected
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises
'   TEST_Harness.ERR_TEST_FAILED, or the unexpected error that escaped the
'   runner. A refused start raises TEST_Harness.ERR_TEST_DIRTY_START
'
' DEPENDENCIES
'   TEST_Harness
'
' NOTES
'   Public surface: RunRelationTests runs the suite. Option Private Module
'   keeps it out of the workbook's external surface. The suite is not
'   certified: it is not listed in .github/excel-evidence-policy.json
'
'   Option Explicit requires explicit declarations, and Option Private
'   Module keeps the tests inside the project
'
'   State ownership: None. The cases build their inputs in memory and
'   change no Excel state
'
'   Usage: in a development workbook holding the modules of tests/modules/,
'   type TEST_Relations.RunRelationTests in the Immediate window. Replace
'   IllustrativeTotal with a call to the core procedure the relations test
'   (docs/methodology/TEST_HARNESS.md, Relation tests)
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-03
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "relations" 'Name of the suite in the log
    Private Const EXPECTED_CASES        As Long = 4             'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 8             'Assertions a complete run evaluates

    'The tolerance of every comparison, in currency units and as a proportion
    '(docs/methodology/TEST_CASES.md, Tolerance)
    Private Const ABSOLUTE_TOLERANCE    As Double = 0.000001    'Absolute bound, in currency units
    Private Const RELATIVE_TOLERANCE    As Double = 0.000000001 'Relative bound, as a proportion

    'The shift that corrupts an output: one currency unit, far outside both bounds for the totals
    'of these cases, which stay below 2,000,000
    Private Const CORRUPTION            As Double = 1#          'Corruption shift, in currency units

    'The amount by which the monotonicity case raises one input
    Private Const INCREMENT             As Double = 1000#       'Raise of one input, in currency units


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunRelationTests()

'
'------------------------------------------------------------------------------
'                              RUN RELATION TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every case in a fixed order between BeginRun and EndRun. A passing
'   run ends with RESULT=PASS; suite=relations; operation=PASS;
'   completeness=COMPLETE; cases=4; assertions=8; failures=0; cleanup=PASS;
'   recovery=NONE
'
' WHY THIS EXISTS
'   Starts the run, calls each case and ends the run on one path, so that
'   an error which escapes a case is saved and handed to EndRun, which
'   reports it
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - begins the run for SUITE_NAME in mode NORMAL with EXPECTED_CASES
'       and EXPECTED_ASSERTIONS
'     - enables the handler RunFailed and calls CaseZeroInput,
'       CaseAdditivity, CaseSymmetry and CaseMonotonicity
'     - disables the handler and ends the run with the saved error number,
'       source and description (0 and empty text when no error escaped)
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it
'   and re-raises it unchanged. BeginRun's refusal propagates directly
'
' DEPENDENCIES
'   TEST_Harness (BeginRun, EndRun); the four Case procedures of this
'   module; SUITE_NAME, EXPECTED_CASES and EXPECTED_ASSERTIONS
'
' NOTES
'   RunFailed saves the error and resumes at Finish, so EndRun runs on the
'   normal path and on the error path
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' START RUN
'------------------------------------------------------------------------------
    'Begin the run; a refused start propagates to the caller
        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Save an error that escapes a case instead of ending the run
        On Error GoTo RunFailed
    'Run every case in a fixed order
        CaseZeroInput
        CaseAdditivity
        CaseSymmetry
        CaseMonotonicity

'------------------------------------------------------------------------------
' FINISH RUN
'------------------------------------------------------------------------------
Finish:
    'Disable the runner's handler before ending the run
        On Error GoTo 0
    'End the run, reporting the saved error if there is one
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
RunFailed:
    'Save the error that escaped a case
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Continue at the end of the run
        Resume Finish

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseZeroInput()

'
'------------------------------------------------------------------------------
'                               TEST ZERO INPUT
'------------------------------------------------------------------------------
' PURPOSE
'   Zero inputs give a zero total, matched within the absolute bound only;
'   a total shifted by CORRUPTION is rejected. 2 assertions.
'
' WHY THIS EXISTS
'   The total of zero amounts is known without an external expected value,
'   so it checks the stand-in against a relation; the corrupted total shows
'   that the comparison can fail
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - begins the case zero-input
'     - totals three zero amounts
'     - asserts that the total is 0 within the absolute bound only
'     - asserts that a total shifted by CORRUPTION is rejected by the same
'       comparison
'
' ERROR POLICY
'   Does not raise: an error of the case goes to the handler CaseFailed,
'   which records it with TEST_Harness.RecordCaseError and returns to the
'   runner
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertNear, AssertTrue, WithinTolerance,
'   RecordCaseError); IllustrativeTotal; ABSOLUTE_TOLERANCE,
'   RELATIVE_TOLERANCE and CORRUPTION
'
' NOTES
'   The amounts are three zeros, as VBA initializes a Double array.
'   The case builds its inputs in memory and changes no Excel state
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Amounts(1 To 3) As Double   'Three zero amounts, as VBA initializes them
    Dim Total           As Double   'Total of Amounts

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Handle an error of the case in CaseFailed, then begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "zero-input"

'------------------------------------------------------------------------------
' COMPUTE TOTAL
'------------------------------------------------------------------------------
    'Total the three zero amounts
        Total = IllustrativeTotal(Amounts)

'------------------------------------------------------------------------------
' ASSERT RELATION
'------------------------------------------------------------------------------
    'Check that the total is zero within the absolute bound only
        TEST_Harness.AssertNear "holds", 0#, Total, ABSOLUTE_TOLERANCE, RELATIVE_TOLERANCE

'------------------------------------------------------------------------------
' ASSERT CORRUPTION
'------------------------------------------------------------------------------
    'Check that the same comparison rejects a total shifted by CORRUPTION
        TEST_Harness.AssertTrue "detects-corruption", _
            Not TEST_Harness.WithinTolerance(0#, Total + CORRUPTION, ABSOLUTE_TOLERANCE, _
            RELATIVE_TOLERANCE), "a corrupted total of zero amounts was accepted"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error of the case and return to the runner
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseAdditivity()

'
'------------------------------------------------------------------------------
'                               TEST ADDITIVITY
'------------------------------------------------------------------------------
' PURPOSE
'   The total of two lists together equals the sum of their totals; a
'   combined total shifted by CORRUPTION is rejected. 2 assertions.
'
' WHY THIS EXISTS
'   The total of a combined list must equal the sum of the totals of its
'   parts, which needs no external expected value; the corrupted total
'   shows that the comparison can fail
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - begins the case additivity
'     - fills a list of 2 amounts and a list of 3 amounts, and copies both
'       into a combined list of 5
'     - totals the combined list, and sums the totals of the two lists
'     - asserts that the two agree within both bounds
'     - asserts that a combined total shifted by CORRUPTION is rejected by
'       the same comparison
'
' ERROR POLICY
'   Does not raise: an error of the case goes to the handler CaseFailed,
'   which records it with TEST_Harness.RecordCaseError and returns to the
'   runner
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertNear, AssertTrue, WithinTolerance,
'   RecordCaseError); IllustrativeTotal; ABSOLUTE_TOLERANCE,
'   RELATIVE_TOLERANCE and CORRUPTION
'
' NOTES
'   The amounts are not exact in binary, so the sums differ in their last
'   bits and the comparison needs its tolerance.
'   The case builds its inputs in memory and changes no Excel state
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FirstList(1 To 2)   As Double   'The first list
    Dim SecondList(1 To 3)  As Double   'The second list
    Dim BothLists(1 To 5)   As Double   'Both lists together
    Dim Index               As Long     'Position in a list
    Dim Whole               As Double   'Total of BothLists
    Dim Parts               As Double   'Sum of the totals of FirstList and SecondList

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Handle an error of the case in CaseFailed, then begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "additivity"

'------------------------------------------------------------------------------
' PREPARE LISTS
'------------------------------------------------------------------------------
    'Use amounts that are not exact in binary, so the sums differ in their last bits
        FirstList(1) = 600000.1
        FirstList(2) = 400000.2
        SecondList(1) = 250000.3
        SecondList(2) = 125000.4
        SecondList(3) = 0.5
    'Copy the first list to the start of the combined list
        For Index = 1 To 2
            BothLists(Index) = FirstList(Index)
        Next Index
    'Copy the second list after the first
        For Index = 1 To 3
            BothLists(2 + Index) = SecondList(Index)
        Next Index

'------------------------------------------------------------------------------
' COMPUTE TOTALS
'------------------------------------------------------------------------------
    'Total the combined list
        Whole = IllustrativeTotal(BothLists)
    'Sum the totals of the two lists
        Parts = IllustrativeTotal(FirstList) + IllustrativeTotal(SecondList)

'------------------------------------------------------------------------------
' ASSERT RELATION
'------------------------------------------------------------------------------
    'Check that the combined total equals the sum of the totals
        TEST_Harness.AssertNear "holds", Parts, Whole, ABSOLUTE_TOLERANCE, RELATIVE_TOLERANCE

'------------------------------------------------------------------------------
' ASSERT CORRUPTION
'------------------------------------------------------------------------------
    'Check that the same comparison rejects a combined total shifted by CORRUPTION
        TEST_Harness.AssertTrue "detects-corruption", _
            Not TEST_Harness.WithinTolerance(Parts, Whole + CORRUPTION, ABSOLUTE_TOLERANCE, _
            RELATIVE_TOLERANCE), "a corrupted combined total was accepted"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error of the case and return to the runner
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseSymmetry()

'
'------------------------------------------------------------------------------
'                                TEST SYMMETRY
'------------------------------------------------------------------------------
' PURPOSE
'   The order of the inputs does not change the total; a total of the
'   reversed list shifted by CORRUPTION is rejected. 2 assertions.
'
' WHY THIS EXISTS
'   A total must not depend on the order of its inputs, which needs no
'   external expected value; the corrupted total shows that the comparison
'   can fail
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - begins the case symmetry
'     - fills a list of 4 amounts and builds the same list in reverse order
'     - totals both lists
'     - asserts that the two totals agree within both bounds
'     - asserts that a total of the reversed list shifted by CORRUPTION is
'       rejected by the same comparison
'
' ERROR POLICY
'   Does not raise: an error of the case goes to the handler CaseFailed,
'   which records it with TEST_Harness.RecordCaseError and returns to the
'   runner
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertNear, AssertTrue, WithinTolerance,
'   RecordCaseError); IllustrativeTotal; ABSOLUTE_TOLERANCE,
'   RELATIVE_TOLERANCE and CORRUPTION
'
' NOTES
'   The case builds its inputs in memory and changes no Excel state
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Forward(1 To 4)     As Double   'The list
    Dim Reversed(1 To 4)    As Double   'The same list in reverse order
    Dim Index               As Long     'Position in the list
    Dim Expected            As Double   'Total of Forward
    Dim Actual              As Double   'Total of Reversed

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Handle an error of the case in CaseFailed, then begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "symmetry"

'------------------------------------------------------------------------------
' PREPARE LISTS
'------------------------------------------------------------------------------
    'Fill the list
        Forward(1) = 0.1
        Forward(2) = 1000000.7
        Forward(3) = 0.2
        Forward(4) = 333333.3
    'Build the same list in reverse order
        For Index = 1 To 4
            Reversed(Index) = Forward(5 - Index)
        Next Index

'------------------------------------------------------------------------------
' COMPUTE TOTALS
'------------------------------------------------------------------------------
    'Total the list in its order
        Expected = IllustrativeTotal(Forward)
    'Total the reversed list
        Actual = IllustrativeTotal(Reversed)

'------------------------------------------------------------------------------
' ASSERT RELATION
'------------------------------------------------------------------------------
    'Check that both orders give the same total
        TEST_Harness.AssertNear "holds", Expected, Actual, ABSOLUTE_TOLERANCE, RELATIVE_TOLERANCE

'------------------------------------------------------------------------------
' ASSERT CORRUPTION
'------------------------------------------------------------------------------
    'Check that the same comparison rejects a reversed total shifted by CORRUPTION
        TEST_Harness.AssertTrue "detects-corruption", _
            Not TEST_Harness.WithinTolerance(Expected, Actual + CORRUPTION, ABSOLUTE_TOLERANCE, _
            RELATIVE_TOLERANCE), "a corrupted total of the reversed list was accepted"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error of the case and return to the runner
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseMonotonicity()

'
'------------------------------------------------------------------------------
'                              TEST MONOTONICITY
'------------------------------------------------------------------------------
' PURPOSE
'   Raising one input by INCREMENT never lowers the total; a raised total
'   corrupted as if the increment had been subtracted is rejected.
'   2 assertions.
'
' WHY THIS EXISTS
'   A total must not fall when one input rises, which needs no external
'   expected value; the corrupted total shows that the comparison can fail
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - begins the case monotonicity
'     - fills a list of 3 amounts and copies it, raising the second amount
'       by INCREMENT
'     - totals both lists
'     - asserts that the raised total is not below the original total
'     - asserts that the raised total minus twice INCREMENT is rejected by
'       the same comparison
'
' ERROR POLICY
'   Does not raise: an error of the case goes to the handler CaseFailed,
'   which records it with TEST_Harness.RecordCaseError and returns to the
'   runner
'
' DEPENDENCIES
'   TEST_Harness (BeginCase, AssertTrue, RecordCaseError);
'   IllustrativeTotal; INCREMENT
'
' NOTES
'   The case builds its inputs in memory and changes no Excel state
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Original(1 To 3)    As Double   'The list
    Dim Raised(1 To 3)      As Double   'The same list with its second amount raised
    Dim Index               As Long     'Position in the list
    Dim OriginalTotal       As Double   'Total of Original
    Dim RaisedTotal         As Double   'Total of Raised

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Handle an error of the case in CaseFailed, then begin the case
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "monotonicity"

'------------------------------------------------------------------------------
' PREPARE LISTS
'------------------------------------------------------------------------------
    'Fill the original list
        Original(1) = 500000#
        Original(2) = 250000.25
        Original(3) = 125000.125
    'Copy it to the raised list
        For Index = 1 To 3
            Raised(Index) = Original(Index)
        Next Index
    'Raise the second amount by INCREMENT
        Raised(2) = Raised(2) + INCREMENT

'------------------------------------------------------------------------------
' COMPUTE TOTALS
'------------------------------------------------------------------------------
    'Total the original list
        OriginalTotal = IllustrativeTotal(Original)
    'Total the raised list
        RaisedTotal = IllustrativeTotal(Raised)

'------------------------------------------------------------------------------
' ASSERT RELATION
'------------------------------------------------------------------------------
    'Check that raising an input does not lower the total
        TEST_Harness.AssertTrue "holds", RaisedTotal >= OriginalTotal, _
            "raised total " & CStr(RaisedTotal) & " is below " & CStr(OriginalTotal)

'------------------------------------------------------------------------------
' ASSERT CORRUPTION
'------------------------------------------------------------------------------
    'Check that a raised total corrupted by twice INCREMENT is rejected
        TEST_Harness.AssertTrue "detects-corruption", _
            Not (RaisedTotal - 2# * INCREMENT >= OriginalTotal), _
            "a corrupted raised total was accepted"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the error of the case and return to the runner
        TEST_Harness.RecordCaseError

End Sub

'
'------------------------------------------------------------------------------
'
'                         EXAMPLE PROCEDURE UNDER TEST
'
'------------------------------------------------------------------------------
'

Private Function IllustrativeTotal( _
    ByRef Amounts() As Double) _
    As Double
'
'------------------------------------------------------------------------------
'                                TOTAL AMOUNTS
'------------------------------------------------------------------------------
' PURPOSE
'   The sum of Amounts: a stand-in for the calculation the relations test.
'   Replace it with a call to that calculation's core procedure
'
' WHY THIS EXISTS
'   Gives the relation cases a calculation to run on, so that the suite
'   works without a core procedure
'
' INPUTS
'   Amounts
'     Amounts in one currency, read only; VBA passes an array by
'     reference
'
' RETURNS
'   The sum of the elements of Amounts from LBound to UBound, in the
'   currency of the amounts
'
' BEHAVIOR
'   Adds the elements of Amounts one by one, from LBound to UBound, and
'   returns the sum
'
' ERROR POLICY
'   Raises nothing of its own: LBound and UBound raise their error for an
'   array that is not allocated
'
' DEPENDENCIES
'   None
'
' NOTES
'   Owns no state
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index   As Long     'Position in Amounts
    Dim Total   As Double   'Sum so far

'------------------------------------------------------------------------------
' ADD UP
'------------------------------------------------------------------------------
    'Add the amounts one by one
        For Index = LBound(Amounts) To UBound(Amounts)
            Total = Total + Amounts(Index)
        Next Index
    'Return the sum
        IllustrativeTotal = Total

End Function
