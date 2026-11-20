Attribute VB_Name = "CORE_Util"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Util
'------------------------------------------------------------------------------
' PURPOSE
'   Shared helpers of the K-FRTB-SSA core: classify cell values (number, date,
'   text), normalise codes, look codes up in arrays, parse the delimited lists
'   of CORE_Config, collect Checks findings, total a Buckets block and take the
'   fingerprint of an input vector
'
' WHY THIS EXISTS
'   Puts the small helpers that read and compare cell values, grow arrays and
'   hash values in one place, with the rules of Excel's ISNUMBER, MATCH and
'   COUNTIF written once. The code does not show which components call them
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Groups the helpers in six families:
'     - value classification: IsNumberValue, NumberValue, IsDateSerial,
'       PlainText, CodeText, DirectionSign and IsInvalidStatus
'     - lookups: FindCode and KeyPosition
'     - delimited lists: ParseList and ListItem
'     - findings: AddFinding
'     - buckets: WriteTotalRow
'     - fingerprint: Fingerprint, with the private helpers HasColumns,
'       ValueKey, AddToHash and HexWord
'
' ERROR POLICY
'   No procedure raises on cell content: an error value in a cell reads as
'   text "#ERROR" or as "not a number". ParseList raises ERR_ARGUMENT for a
'   list with the wrong number of items
'
' DEPENDENCIES
'   CORE_Config (DATE_SERIAL_MIN, DATE_SERIAL_MAX, ST_INVALID and ERR_ARGUMENT)
'
' NOTES
'   Public surface: None outside this VBA project (Option Private Module). The
'   module owns no state: every procedure depends only on its arguments
'
'   Option Explicit requires explicit declarations, and Option Private Module
'   keeps the helpers inside this project
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
    Private Const ERROR_TEXT            As String = "#ERROR"    'Text of a cell error value
    Private Const FINGERPRINT_VERSION   As String = "F1"        'Form of a fingerprint
    Private Const HASH_BASE_1           As Double = 131#        'Multiplier of the first hash
    Private Const HASH_MODULUS_1        As Double = 2147483647# 'Prime 2^31 - 1
    Private Const HASH_BASE_2           As Double = 137#        'Multiplier of the second hash
    Private Const HASH_MODULUS_2        As Double = 2147483629# 'Prime 2^31 - 19
    Private Const HASH_SEPARATOR        As Long = 30            'Code added after each value
    Private Const CODE_UNITS            As Long = 65536         'Values of a UTF-16 code unit


'
'------------------------------------------------------------------------------
'
'                             VALUE CLASSIFICATION
'
'------------------------------------------------------------------------------
'

Public Function IsNumberValue( _
    ByVal Value As Variant) _
    As Boolean
'
'------------------------------------------------------------------------------
'                              TEST NUMBER VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Tells whether a cell value is a number, as Excel's ISNUMBER does: numeric
'   types and dates are numbers; text that looks like a number, booleans, empty
'   cells and error values are not
'
' WHY THIS EXISTS
'   Callers read cell values as Variant and need one rule for what counts as a
'   number, the rule of Excel's ISNUMBER
'
' INPUTS
'   Value
'     Cell value as read from Range.Value
'
' RETURNS
'   True for a number or a date
'
' BEHAVIOR
'   Classifies the VarType of the value: Integer, Long, Single, Double,
'   Currency, Decimal and Date are numbers; any other type is not
'
' ERROR POLICY
'   Never raises
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLASSIFY VALUE
'------------------------------------------------------------------------------
    'Treat the numeric types and dates as numbers, and every other type as not a number
        Select Case VarType(Value)
            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbDate
                IsNumberValue = True
            Case Else
                IsNumberValue = False
        End Select

End Function

Public Function NumberValue( _
    ByVal Value As Variant) _
    As Double
'
'------------------------------------------------------------------------------
'                           CONVERT VALUE TO NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a cell value as a Double: a number as itself, a date as its serial
'   number, an empty cell as 0 (as Excel arithmetic does)
'
' WHY THIS EXISTS
'   Callers add up cell values, and Excel arithmetic reads an empty cell as 0
'
' INPUTS
'   Value
'     Cell value; IsNumberValue(Value) or IsEmpty(Value)
'
' RETURNS
'   The value, or 0 for anything else
'
' BEHAVIOR
'   Converts the value with CDbl when IsNumberValue accepts it, and returns 0
'   otherwise
'
' ERROR POLICY
'   Never raises for a value that IsNumberValue accepts or for an empty cell
'
' DEPENDENCIES
'   IsNumberValue
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CONVERT VALUE
'------------------------------------------------------------------------------
    'Convert a number or a date, and give 0 to anything else
        If IsNumberValue(Value) Then
            NumberValue = CDbl(Value)
        Else
            NumberValue = 0#
        End If

End Function

Public Function IsDateSerial( _
    ByVal Value As Variant) _
    As Boolean
'
'------------------------------------------------------------------------------
'                               TEST DATE SERIAL
'------------------------------------------------------------------------------
' PURPOSE
'   Tells whether a cell value is a date: a date, or a number that is the
'   serial of a date
'
' WHY THIS EXISTS
'   A date typed as a number such as 20311231 is not a date serial, and the
'   range test tells it apart
'
' INPUTS
'   Value
'     Cell value
'
' RETURNS
'   True for a number from DATE_SERIAL_MIN to DATE_SERIAL_MAX
'
' BEHAVIOR
'   Returns False for a value that is not a number, and otherwise tests the
'   number against the serial range of the dates
'
' ERROR POLICY
'   Never raises
'
' DEPENDENCIES
'   IsNumberValue, DATE_SERIAL_MIN and DATE_SERIAL_MAX of CORE_Config
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' TEST DATE RANGE
'------------------------------------------------------------------------------
    'Accept a number that lies in the serial range of the dates
        If IsNumberValue(Value) Then
            IsDateSerial = (CDbl(Value) >= DATE_SERIAL_MIN And CDbl(Value) <= DATE_SERIAL_MAX)
        Else
            IsDateSerial = False
        End If

End Function

Public Function PlainText( _
    ByVal Value As Variant) _
    As String
'
'------------------------------------------------------------------------------
'                              RETURN PLAIN TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a cell value as trimmed text, case kept
'
' WHY THIS EXISTS
'   Callers compare and report cell content as text, and an error value must
'   not raise when it is read as text
'
' INPUTS
'   Value
'     Cell value of any type
'
' RETURNS
'   Trimmed text; "" for an empty cell; "#ERROR" for an error value
'
' BEHAVIOR
'   Returns ERROR_TEXT for an error value, an empty string for an Empty or Null
'   value, and the trimmed text of the value otherwise
'
' ERROR POLICY
'   Never raises: an error value reads as ERROR_TEXT
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' NORMALISE TEXT
'------------------------------------------------------------------------------
    'Read an error as the error text, an empty value as an empty string, and any other value as trimmed text
        If IsError(Value) Then
            PlainText = ERROR_TEXT
        ElseIf IsEmpty(Value) Or IsNull(Value) Then
            PlainText = vbNullString
        Else
            PlainText = Trim$(CStr(Value))
        End If

End Function

Public Function CodeText( _
    ByVal Value As Variant) _
    As String
'
'------------------------------------------------------------------------------
'                               RETURN CODE TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a cell value as a code: trimmed and upper case, so that codes
'   compare without regard to case or surrounding blanks (as Excel's MATCH and
'   COUNTIF do)
'
' WHY THIS EXISTS
'   Codes are compared the way Excel's MATCH and COUNTIF compare them, which
'   ignores case and surrounding blanks
'
' INPUTS
'   Value
'     Cell value of any type
'
' RETURNS
'   Upper-case trimmed text; "" for an empty cell
'
' BEHAVIOR
'   Converts the value with PlainText and puts the result in upper case
'
' ERROR POLICY
'   Never raises: PlainText reads an error value as text
'
' DEPENDENCIES
'   PlainText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' NORMALISE CODE
'------------------------------------------------------------------------------
    'Trim the value and put it in upper case
        CodeText = UCase$(PlainText(Value))

End Function

Public Function DirectionSign( _
    ByVal Value As Variant) _
    As Double
'
'------------------------------------------------------------------------------
'                            RETURN DIRECTION SIGN
'------------------------------------------------------------------------------
' PURPOSE
'   Turns a Direction cell into the sign of the position
'
' WHY THIS EXISTS
'   A position's direction is written as text in a sheet, and the calculation
'   needs it as a sign
'
' INPUTS
'   Value
'     "Long" or "Short", any case
'
' RETURNS
'   1 for Long, -1 for Short, 0 for anything else
'
' BEHAVIOR
'   Normalises the value with CodeText and maps LONG to 1 and SHORT to -1; any
'   other text maps to 0
'
' ERROR POLICY
'   Never raises: an unrecognised direction returns 0
'
' DEPENDENCIES
'   CodeText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' MAP DIRECTION
'------------------------------------------------------------------------------
    'Map the direction code to its sign, and an unknown code to zero
        Select Case CodeText(Value)
            Case "LONG"
                DirectionSign = 1#
            Case "SHORT"
                DirectionSign = -1#
            Case Else
                DirectionSign = 0#
        End Select

End Function

Public Function IsInvalidStatus( _
    ByVal Status As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                             TEST INVALID STATUS
'------------------------------------------------------------------------------
' PURPOSE
'   Tells whether a row status rejects the row
'
' WHY THIS EXISTS
'   A row rejected by the core carries a status that starts with ST_INVALID,
'   and callers need one test for it
'
' INPUTS
'   Status
'     Row status as written by the core
'
' RETURNS
'   True when the status starts with "INVALID: "
'
' BEHAVIOR
'   Compares the first Len(ST_INVALID) characters of the status with
'   ST_INVALID, case-sensitively
'
' ERROR POLICY
'   Never raises
'
' DEPENDENCIES
'   ST_INVALID of CORE_Config
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' TEST STATUS PREFIX
'------------------------------------------------------------------------------
    'Check whether the status starts with the invalid-row prefix
        IsInvalidStatus = (Left$(Status, Len(ST_INVALID)) = ST_INVALID)

End Function

'
'------------------------------------------------------------------------------
'
'                                   LOOKUPS
'
'------------------------------------------------------------------------------
'

Public Function FindCode( _
    ByRef Codes() As String, _
    ByVal Code As String) _
    As Long
'
'------------------------------------------------------------------------------
'                                  FIND CODE
'------------------------------------------------------------------------------
' PURPOSE
'   Finds a code in a list of codes, ignoring case and surrounding blanks
'
' WHY THIS EXISTS
'   Codes are compared as Excel's MATCH compares them, without regard to case
'   or surrounding blanks
'
' INPUTS
'   Codes
'     List of codes, any bounds; blank entries never match
'
'   Code
'     Code to find
'
' RETURNS
'   Index of the first match in Codes, or LBound(Codes) - 1 when absent
'
' BEHAVIOR
'   Normalises the code, returns the not-found index for a blank code, and
'   scans the list from its lower bound to its upper bound
'
' ERROR POLICY
'   Never raises for a list with bounds
'
' DEPENDENCIES
'   None
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
    Dim Wanted  As String   'Normalised code
    Dim i       As Long     'Index in Codes

'------------------------------------------------------------------------------
' PREPARE SEARCH
'------------------------------------------------------------------------------
    'Normalise the code to find
        Wanted = UCase$(Trim$(Code))
    'Start from the not-found index
        FindCode = LBound(Codes) - 1
    'Leave a blank code unmatched
        If Len(Wanted) = 0 Then Exit Function

'------------------------------------------------------------------------------
' SEARCH CODES
'------------------------------------------------------------------------------
    'Return the index of the first entry that matches the code
        For i = LBound(Codes) To UBound(Codes)
            If UCase$(Trim$(Codes(i))) = Wanted Then
                FindCode = i
                Exit Function
            End If
        Next i

End Function

Public Function KeyPosition( _
    ByRef Keys() As String, _
    ByRef KeyCount As Long, _
    ByVal Key As String, _
    ByVal AddMissing As Boolean) _
    As Long
'
'------------------------------------------------------------------------------
'                               FIND OR ADD KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Finds a key in a growing list of keys, optionally appending it. Keys are
'   compared without regard to case, in order of first appearance
'
' WHY THIS EXISTS
'   Keeps a list of distinct keys in order of first appearance, growing it
'   as keys are added
'
' INPUTS
'   Keys
'     List, base 1; grown with ReDim Preserve when a key is added
'
'   KeyCount
'     Number of keys in use (>= 0); incremented when one is added
'
'   Key
'     Key to find; "" is never found or added
'
'   AddMissing
'     True to append a missing key
'
' RETURNS
'   Position of the key (1 to KeyCount), or 0 when absent and not added
'
' BEHAVIOR
'   Returns 0 for a blank key, scans the keys in use for a match and returns
'   its position, and when the key is absent and AddMissing is True appends it,
'   starting the list at one item and doubling it when it is full
'
' ERROR POLICY
'   Never raises for a list with bounds
'
' DEPENDENCIES
'   None
'
' NOTES
'   The matching ignores case, but the key is stored as it was passed
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Wanted  As String   'Normalised key
    Dim i       As Long     'Position in Keys

'------------------------------------------------------------------------------
' PREPARE SEARCH
'------------------------------------------------------------------------------
    'Start from the not-found position
        KeyPosition = 0
    'Normalise the key to find
        Wanted = UCase$(Key)
    'Leave a blank key unmatched
        If Len(Wanted) = 0 Then Exit Function

'------------------------------------------------------------------------------
' SEARCH KEYS
'------------------------------------------------------------------------------
    'Return the position of the first key in use that matches
        For i = 1 To KeyCount
            If UCase$(Keys(i)) = Wanted Then
                KeyPosition = i
                Exit Function
            End If
        Next i

'------------------------------------------------------------------------------
' ADD MISSING KEY
'------------------------------------------------------------------------------
    'Append the key when asked to, growing the list when it is empty or full
        If AddMissing Then
            KeyCount = KeyCount + 1
            If KeyCount = 1 Then
                ReDim Keys(1 To 1)
            ElseIf KeyCount > UBound(Keys) Then
                ReDim Preserve Keys(1 To 2 * UBound(Keys))
            End If
            Keys(KeyCount) = Key
            KeyPosition = KeyCount
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                               DELIMITED LISTS
'
'------------------------------------------------------------------------------
'

Public Sub ParseList( _
    ByVal ListText As String, _
    ByVal ExpectedCount As Long, _
    ByVal BlankValue As Double, _
    ByRef Values() As Double)
'
'------------------------------------------------------------------------------
'                             PARSE DELIMITED LIST
'------------------------------------------------------------------------------
' PURPOSE
'   Turns a "|"-delimited list of decimal numbers (decimal point) into an array.
'   Val is used, so the Windows locale cannot change the result
'
' WHY THIS EXISTS
'   The lists of CORE_Config hold their numbers as text, and Val reads them with
'   a decimal point whatever the regional setting
'
' INPUTS
'   ListText
'     Items separated by "|"
'
'   ExpectedCount
'     Number of items the list must have (>= 1)
'
'   BlankValue
'     Value given to an empty item
'
'   Values
'     Output array, redimensioned here to 1 To ExpectedCount
'
' RETURNS
'   Nothing. Values(1 To ExpectedCount) holds the numbers
'
' BEHAVIOR
'   Splits the text at the delimiter, checks the item count, and converts each
'   item with Val, giving BlankValue to an item that is empty after trimming
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when the item count differs from ExpectedCount
'
' DEPENDENCIES
'   ERR_ARGUMENT of CORE_Config
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
    Dim Items() As String   'Items of the list
    Dim i       As Long     'Item index

'------------------------------------------------------------------------------
' SPLIT LIST
'------------------------------------------------------------------------------
    'Split the text at the delimiter
        Items = Split(ListText, "|")

'------------------------------------------------------------------------------
' VALIDATE ITEM COUNT
'------------------------------------------------------------------------------
    'Reject a list whose item count is not the expected one
        If UBound(Items) - LBound(Items) + 1 <> ExpectedCount Then
            Err.Raise ERR_ARGUMENT, "CORE_Util.ParseList", _
                "List has " & (UBound(Items) - LBound(Items) + 1) & " items, " & _
                ExpectedCount & " expected: " & ListText
        End If

'------------------------------------------------------------------------------
' CONVERT ITEMS
'------------------------------------------------------------------------------
    'Size the output array
        ReDim Values(1 To ExpectedCount)
    'Convert each item, and give the blank value to an empty one
        For i = 1 To ExpectedCount
            If Len(Trim$(Items(LBound(Items) + i - 1))) = 0 Then
                Values(i) = BlankValue
            Else
                Values(i) = Val(Items(LBound(Items) + i - 1))
            End If
        Next i

End Sub

Public Function ListItem( _
    ByVal ListText As String, _
    ByVal Position As Long) _
    As String
'
'------------------------------------------------------------------------------
'                               RETURN LIST ITEM
'------------------------------------------------------------------------------
' PURPOSE
'   Returns one item of a "|"-delimited list
'
' WHY THIS EXISTS
'   Callers read one item of a delimited list of CORE_Config without parsing
'   the whole list
'
' INPUTS
'   ListText
'     Items separated by "|"
'
'   Position
'     Item number, 1-based; 1 <= Position <= item count
'
' RETURNS
'   The item text
'
' BEHAVIOR
'   Splits the text at the delimiter and returns the item at the position
'
' ERROR POLICY
'   Raises the VBA subscript error when Position is outside the list; the
'   procedure does not check it
'
' DEPENDENCIES
'   None
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
    Dim Items() As String   'Items of the list

'------------------------------------------------------------------------------
' SELECT ITEM
'------------------------------------------------------------------------------
    'Split the text at the delimiter
        Items = Split(ListText, "|")
    'Return the item at the 1-based position
        ListItem = Items(LBound(Items) + Position - 1)

End Function

'
'------------------------------------------------------------------------------
'
'                                   FINDINGS
'
'------------------------------------------------------------------------------
'

Public Sub AddFinding( _
    ByRef Findings() As String, _
    ByRef FindingCount As Long, _
    ByVal SheetName As String, _
    ByVal Item As String, _
    ByVal RowText As String, _
    ByVal Reason As String)
'
'------------------------------------------------------------------------------
'                                 ADD FINDING
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one finding (a rejected row or an unusable parameter) to the list
'   written to the Checks sheet
'
' WHY THIS EXISTS
'   Keeps the findings of the Checks sheet in one matrix that grows as
'   findings are added
'
' INPUTS
'   Findings
'     Matrix (1 To 4, 1 To n): sheet, item, row, reason; grown here.
'     Undimensioned when FindingCount = 0
'
'   FindingCount
'     Findings in use (>= 0); incremented
'
'   SheetName
'     Sheet the finding is about
'
'   Item
'     PositionID or parameter code
'
'   RowText
'     Sheet row as text, or "" when not applicable
'
'   Reason
'     Message, starting "INVALID: " for a rejected row
'
' RETURNS
'   Nothing. The finding is stored in Findings
'
' BEHAVIOR
'   Increments the count, creates the matrix with 8 columns for the first
'   finding or doubles its columns when it is full, and stores the four fields
'   in the new column
'
' ERROR POLICY
'   Never raises for a list that follows the rules of Findings
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' GROW FINDINGS
'------------------------------------------------------------------------------
    'Count the finding
        FindingCount = FindingCount + 1
    'Create the matrix for the first finding, and double its columns when it is full
        If FindingCount = 1 Then
            ReDim Findings(1 To 4, 1 To 8)
        ElseIf FindingCount > UBound(Findings, 2) Then
            ReDim Preserve Findings(1 To 4, 1 To 2 * UBound(Findings, 2))
        End If

'------------------------------------------------------------------------------
' STORE FINDING
'------------------------------------------------------------------------------
    'Store the sheet, item, row and reason in the new column
        Findings(1, FindingCount) = SheetName
        Findings(2, FindingCount) = Item
        Findings(3, FindingCount) = RowText
        Findings(4, FindingCount) = Reason

End Sub

'
'------------------------------------------------------------------------------
'
'                                   BUCKETS
'
'------------------------------------------------------------------------------
'

Public Sub WriteTotalRow( _
    ByRef BucketOutput() As Variant, _
    ByVal BucketCount As Long, _
    ByVal Label As String)
'
'------------------------------------------------------------------------------
'                               WRITE TOTAL ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Fills the total row of an equity or commodity Buckets block: the label in
'   the second column, sums from the third, the net column (6) left blank
'
' WHY THIS EXISTS
'   The equity and commodity Buckets blocks end in a total row, and one
'   procedure fills it for both
'
' INPUTS
'   BucketOutput
'     (1 To BucketCount + 1, 1 To width); rows 1 to BucketCount are
'     filled
'
'   BucketCount
'     Bucket rows above the total row (>= 0)
'
'   Label
'     Text of the total row
'
' RETURNS
'   Nothing. Row BucketCount + 1 of BucketOutput is filled
'
' BEHAVIOR
'   Writes the label in column 2 and, for each column from 3 to the last, sums
'   the bucket rows; column 6 is set to Empty instead
'
' ERROR POLICY
'   Never raises for an array that follows the rules of BucketOutput
'
' DEPENDENCIES
'   None
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
    Dim Total   As Double   'Column sum
    Dim m       As Long     'Bucket row
    Dim c       As Long     'Column

'------------------------------------------------------------------------------
' WRITE LABEL
'------------------------------------------------------------------------------
    'Put the label in the second column of the total row
        BucketOutput(BucketCount + 1, 2) = Label

'------------------------------------------------------------------------------
' TOTAL COLUMNS
'------------------------------------------------------------------------------
    'Sum each column over the bucket rows, and leave the net column blank
        For c = 3 To UBound(BucketOutput, 2)
            If c = 6 Then
                BucketOutput(BucketCount + 1, c) = Empty
            Else
                Total = 0#
                For m = 1 To BucketCount
                    Total = Total + BucketOutput(m, c)
                Next m
                BucketOutput(BucketCount + 1, c) = Total
            End If
        Next c

End Sub

'
'------------------------------------------------------------------------------
'
'                                 FINGERPRINT
'
'------------------------------------------------------------------------------
'

Public Function Fingerprint( _
    ByRef Values() As Variant) _
    As String
'
'------------------------------------------------------------------------------
'                              RETURN FINGERPRINT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a fingerprint of a vector of values: the same values in the same
'   order give the same text, and a changed, added or removed value gives
'   another one, short of a collision of two 31-bit hashes. It tells a change;
'   it is not a cryptographic digest
'
' WHY THIS EXISTS
'   It tells whether an input vector has changed, with a text that does not
'   depend on the regional setting
'
' INPUTS
'   Values
'     Vector whose items are cell values, or arrays of cell values of
'     one or two dimensions
'
' RETURNS
'   FINGERPRINT_VERSION, "-" and sixteen hexadecimal digits
'
' BEHAVIOR
'   Starts two polynomial hashes from their bases and, for each item, adds its
'   key to them; an array adds a key with its row count and then the key of each
'   of its values, by row and column for a two-dimensional array. Returns the
'   version, a hyphen and the two hashes as eight hexadecimal digits each
'
' ERROR POLICY
'   Never raises for a vector of cell values or of arrays of them
'
' DEPENDENCIES
'   ValueKey, HasColumns, AddToHash and HexWord
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
    Dim First   As Double   'First running hash
    Dim Second  As Double   'Second running hash
    Dim Table   As Variant  'One array of Values
    Dim i       As Long     'Item of Values
    Dim r       As Long     'Row of Table
    Dim c       As Long     'Column of Table

'------------------------------------------------------------------------------
' SEED HASHES
'------------------------------------------------------------------------------
    'Start each hash from its base
        First = HASH_BASE_1
        Second = HASH_BASE_2

'------------------------------------------------------------------------------
' HASH VALUES
'------------------------------------------------------------------------------
    'Add each item of the vector to the hashes
        For i = LBound(Values) To UBound(Values)
            If Not IsArray(Values(i)) Then
                AddToHash ValueKey(Values(i)), First, Second
            Else
                Table = Values(i)
            'Add the row count, so that a row moved to another table changes the sequence of values
                AddToHash "A" & CStr(UBound(Table, 1) - LBound(Table, 1) + 1), First, Second
            'Add the values of the array, by row and column when it has two dimensions
                If HasColumns(Table) Then
                    For r = LBound(Table, 1) To UBound(Table, 1)
                        For c = LBound(Table, 2) To UBound(Table, 2)
                            AddToHash ValueKey(Table(r, c)), First, Second
                        Next c
                    Next r
                Else
                    For r = LBound(Table, 1) To UBound(Table, 1)
                        AddToHash ValueKey(Table(r)), First, Second
                    Next r
                End If
            End If
        Next i

'------------------------------------------------------------------------------
' RETURN FINGERPRINT
'------------------------------------------------------------------------------
    'Join the version and the two hashes
        Fingerprint = FINGERPRINT_VERSION & "-" & HexWord(First) & HexWord(Second)

End Function

Private Function HasColumns( _
    ByRef Table As Variant) _
    As Boolean
'
'------------------------------------------------------------------------------
'                            TEST SECOND DIMENSION
'------------------------------------------------------------------------------
' PURPOSE
'   Tells whether an array has a second dimension
'
' WHY THIS EXISTS
'   Asking an array of one dimension for its second bound raises an error, and
'   Fingerprint needs to tell the two shapes apart
'
' INPUTS
'   Table
'     Array of one or two dimensions
'
' RETURNS
'   True when UBound(Table, 2) is available and not negative
'
' BEHAVIOR
'   Asks for the upper bound of the second dimension, with -1 as the value when
'   the dimension does not exist, and compares it with 0
'
' ERROR POLICY
'   The error of asking an array of one dimension for its second bound is
'   contained on the single statement that can raise it
'
' DEPENDENCIES
'   None
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
    Dim Upper   As Long 'Upper bound of the second dimension

'------------------------------------------------------------------------------
' PROBE SECOND DIMENSION
'------------------------------------------------------------------------------
    'Assume no second dimension
        Upper = -1
    'Ask for the second bound, suppressing the error of a one-dimensional array on this statement only
        On Error Resume Next
        Upper = UBound(Table, 2)
        On Error GoTo 0

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Report whether the second bound exists
        HasColumns = (Upper >= 0)

End Function

Private Function ValueKey( _
    ByVal Value As Variant) _
    As String
'
'------------------------------------------------------------------------------
'                               RETURN VALUE KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the text a cell value enters a fingerprint with: its kind and its
'   content, the same in every regional setting
'
' WHY THIS EXISTS
'   The fingerprint of a vector must change when a value or its kind changes,
'   and must not depend on the regional setting
'
' INPUTS
'   Value
'     Cell value
'
' RETURNS
'   "E" for an error value, "_" for an empty cell, "S", the length, ":" and the
'   text for a text, "B0" or "B1" for a Boolean, "N" and the number for a number
'   or a date, and "?" with the type name for anything else
'
' BEHAVIOR
'   Tests the value in this order: error, empty or null, text, Boolean,
'   number or date, anything else, and builds the key of the first that applies
'
' ERROR POLICY
'   Never raises: an error value gets the key "E"
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' BUILD KEY
'------------------------------------------------------------------------------
    'Give each kind of value its own key
        If IsError(Value) Then
            ValueKey = "E"
        ElseIf IsEmpty(Value) Or IsNull(Value) Then
            ValueKey = "_"
        ElseIf VarType(Value) = vbString Then
            ValueKey = "S" & CStr(Len(Value)) & ":" & Value
        ElseIf VarType(Value) = vbBoolean Then
            If Value Then
                ValueKey = "B1"
            Else
                ValueKey = "B0"
            End If
        ElseIf IsNumeric(Value) Or VarType(Value) = vbDate Then
        'Write the number with Str$, which gives a decimal point whatever the regional setting
            ValueKey = "N" & Str$(CDbl(Value))
        Else
            ValueKey = "?" & TypeName(Value)
        End If

End Function

Private Sub AddToHash( _
    ByVal Text As String, _
    ByRef First As Double, _
    ByRef Second As Double)
'
'------------------------------------------------------------------------------
'                              ADD TEXT TO HASHES
'------------------------------------------------------------------------------
' PURPOSE
'   Adds a text and a separator to two running polynomial hashes
'
' WHY THIS EXISTS
'   Fingerprint builds its two hashes one key at a time, and the separator
'   keeps the end of one key apart from the start of the next
'
' INPUTS
'   Text
'     Text to add; may be empty
'
'   First
'     Hash modulo HASH_MODULUS_1, updated
'
'   Second
'     Hash modulo HASH_MODULUS_2, updated
'
' RETURNS
'   Nothing. First and Second are updated
'
' BEHAVIOR
'   Takes the code unit of each character of the text, then HASH_SEPARATOR, and
'   folds each code into both hashes: multiplies by the base, adds the code and
'   reduces modulo the prime
'
' ERROR POLICY
'   Never raises
'
' DEPENDENCIES
'   None
'
' NOTES
'   A product stays far below 2^53, so the Double arithmetic is exact
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Code    As Long 'Code unit of one character
    Dim p       As Long 'Position in Text

'------------------------------------------------------------------------------
' HASH CHARACTERS
'------------------------------------------------------------------------------
    'Hash each character, then the separator
        For p = 1 To Len(Text) + 1
            If p > Len(Text) Then
                Code = HASH_SEPARATOR
            Else
                Code = AscW(Mid$(Text, p, 1))
            'Turn the signed Integer that AscW returns into an unsigned code unit
                If Code < 0 Then Code = Code + CODE_UNITS
            End If
        'Fold the code into both hashes; a product stays far below 2^53, so the arithmetic is exact
            First = First * HASH_BASE_1 + Code
            First = First - Int(First / HASH_MODULUS_1) * HASH_MODULUS_1
            Second = Second * HASH_BASE_2 + Code
            Second = Second - Int(Second / HASH_MODULUS_2) * HASH_MODULUS_2
        Next p

End Sub

Private Function HexWord( _
    ByVal Value As Double) _
    As String
'
'------------------------------------------------------------------------------
'                               RETURN HEX WORD
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a whole number from 0 to 2^31 - 1 as eight hexadecimal digits
'
' WHY THIS EXISTS
'   A hash is written in the fingerprint as a fixed-width hexadecimal word
'
' INPUTS
'   Value
'     Whole number from 0 to 2^31 - 1
'
' RETURNS
'   Eight hexadecimal digits, left-padded with zeros
'
' BEHAVIOR
'   Converts the value with CLng and Hex$, pads the digits with zeros on the
'   left and keeps the last eight
'
' ERROR POLICY
'   Raises the VBA overflow error for a value outside the range of a Long
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FORMAT HEX
'------------------------------------------------------------------------------
    'Write the value in hexadecimal, padded with zeros to eight digits
        HexWord = Right$("0000000" & Hex$(CLng(Value)), 8)

End Function
