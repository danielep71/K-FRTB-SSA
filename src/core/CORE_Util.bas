Attribute VB_Name = "CORE_Util"
'==============================================================================
' MODULE: CORE_Util
'------------------------------------------------------------------------------
' PURPOSE
'   Shared helpers of the K-FRTB-SSA core: classify cell values (number, date,
'   text), normalise codes, look codes up in arrays, parse the delimited
'   lists of CORE_Config, collect Checks findings, total a Buckets block and
'   take the fingerprint of an input vector.
'
' PUBLIC SURFACE
'   None outside this VBA project (Option Private Module).
'
' DEPENDENCIES
'   CORE_Config.
'
' STATE OWNERSHIP
'   None. Every procedure depends only on its arguments.
'
' ERROR POLICY
'   No procedure raises on cell content: an error value in a cell reads as
'   text "#ERROR" or as "not a number". ParseList raises ERR_ARGUMENT for a
'   list with the wrong number of items.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults. Never touches the Excel
'   object model.
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
    'Require explicit declarations; keep the helpers inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const ERROR_TEXT As String = "#ERROR"    'Text of a cell error value
        Private Const FINGERPRINT_VERSION As String = "F1"      'Form of a fingerprint
        Private Const HASH_BASE_1    As Double = 131#           'Multiplier of the first hash
        Private Const HASH_MODULUS_1 As Double = 2147483647#    'Prime 2^31 - 1
        Private Const HASH_BASE_2    As Double = 137#           'Multiplier of the second hash
        Private Const HASH_MODULUS_2 As Double = 2147483629#    'Prime 2^31 - 19
        Private Const HASH_SEPARATOR As Long = 30               'Code added after each value
        Private Const CODE_UNITS     As Long = 65536            'Values of a UTF-16 code unit


'------------------------------------------------------------------------------
'                              VALUE CLASSIFICATION
'------------------------------------------------------------------------------

Public Function IsNumberValue( _
    ByVal Value As Variant) _
    As Boolean
'
'==============================================================================
'                                 IsNumberValue
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a cell value is a number, as Excel's ISNUMBER does: numeric
'   types and dates are numbers; text that looks like a number, booleans,
'   empty cells and error values are not.
'
' INPUTS
'   Value   Cell value as read from Range.Value.
'
' RETURNS
'   True for a number or a date.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
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
'==============================================================================
'                                  NumberValue
'------------------------------------------------------------------------------
' PURPOSE
'   Return a cell value as a Double: a number as itself, a date as its
'   serial number, an empty cell as 0 (as Excel arithmetic does).
'
' INPUTS
'   Value   Cell value; IsNumberValue(Value) or IsEmpty(Value).
'
' RETURNS
'   The value, or 0 for anything else.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
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
'==============================================================================
'                                 IsDateSerial
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a cell value is a date: a date, or a number that is the
'   serial of a date. A date typed as a number such as 20311231 is not.
'
' INPUTS
'   Value   Cell value.
'
' RETURNS
'   True for a number from DATE_SERIAL_MIN to DATE_SERIAL_MAX.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
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
'==============================================================================
'                                   PlainText
'------------------------------------------------------------------------------
' PURPOSE
'   Return a cell value as trimmed text, case kept.
'
' INPUTS
'   Value   Cell value of any type.
'
' RETURNS
'   Trimmed text; "" for an empty cell; "#ERROR" for an error value.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
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
'==============================================================================
'                                   CodeText
'------------------------------------------------------------------------------
' PURPOSE
'   Return a cell value as a code: trimmed and upper case, so that codes
'   compare without regard to case or surrounding blanks (as Excel's MATCH
'   and COUNTIF do).
'
' INPUTS
'   Value   Cell value of any type.
'
' RETURNS
'   Upper-case trimmed text; "" for an empty cell.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        CodeText = UCase$(PlainText(Value))

End Function


Public Function DirectionSign( _
    ByVal Value As Variant) _
    As Double
'
'==============================================================================
'                                 DirectionSign
'------------------------------------------------------------------------------
' PURPOSE
'   Turn a Direction cell into the sign of the position.
'
' INPUTS
'   Value   "Long" or "Short", any case.
'
' RETURNS
'   1 for Long, -1 for Short, 0 for anything else.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
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
'==============================================================================
'                                IsInvalidStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a row status rejects the row.
'
' INPUTS
'   Status   Row status as written by the core.
'
' RETURNS
'   True when the status starts with "INVALID: ".
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        IsInvalidStatus = (Left$(Status, Len(ST_INVALID)) = ST_INVALID)

End Function


'------------------------------------------------------------------------------
'                                    LOOKUPS
'------------------------------------------------------------------------------

Public Function FindCode( _
    ByRef Codes() As String, _
    ByVal Code As String) _
    As Long
'
'==============================================================================
'                                   FindCode
'------------------------------------------------------------------------------
' PURPOSE
'   Find a code in a list of codes, ignoring case and surrounding blanks.
'
' INPUTS
'   Codes   List of codes, any bounds; blank entries never match.
'   Code    Code to find.
'
' RETURNS
'   Index of the first match in Codes, or LBound(Codes) - 1 when absent.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Wanted      As String    'Normalised code
    Dim i           As Long      'Index in Codes

        Wanted = UCase$(Trim$(Code))
        FindCode = LBound(Codes) - 1
        If Len(Wanted) = 0 Then Exit Function
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
'==============================================================================
'                                  KeyPosition
'------------------------------------------------------------------------------
' PURPOSE
'   Find a key in a growing list of keys, optionally appending it. Keys are
'   compared without regard to case, in order of first appearance.
'
' INPUTS
'   Keys       List, base 1; grown with ReDim Preserve when a key is added.
'   KeyCount   Number of keys in use (>= 0); incremented when one is added.
'   Key        Key to find; "" is never found or added.
'   AddMissing True to append a missing key.
'
' RETURNS
'   Position of the key (1 to KeyCount), or 0 when absent and not added.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Wanted      As String    'Normalised key
    Dim i           As Long      'Position in Keys

        KeyPosition = 0
        Wanted = UCase$(Key)
        If Len(Wanted) = 0 Then Exit Function
        For i = 1 To KeyCount
            If UCase$(Keys(i)) = Wanted Then
                KeyPosition = i
                Exit Function
            End If
        Next i
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


'------------------------------------------------------------------------------
'                                 DELIMITED LISTS
'------------------------------------------------------------------------------

Public Sub ParseList( _
    ByVal ListText As String, _
    ByVal ExpectedCount As Long, _
    ByVal BlankValue As Double, _
    ByRef Values() As Double)
'
'==============================================================================
'                                   ParseList
'------------------------------------------------------------------------------
' PURPOSE
'   Turn a "|"-delimited list of decimal numbers (decimal point) into an
'   array. Val is used, so the Windows locale cannot change the result.
'
' INPUTS
'   ListText       Items separated by "|".
'   ExpectedCount  Number of items the list must have (>= 1).
'   BlankValue     Value given to an empty item.
'
' RETURNS
'   Values(1 To ExpectedCount).
'
' ERROR POLICY
'   Raises ERR_ARGUMENT when the item count differs from ExpectedCount.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Items()     As String    'Items of the list
    Dim i           As Long      'Item index

        Items = Split(ListText, "|")
        If UBound(Items) - LBound(Items) + 1 <> ExpectedCount Then
            Err.Raise ERR_ARGUMENT, "CORE_Util.ParseList", _
                "List has " & (UBound(Items) - LBound(Items) + 1) & " items, " & _
                ExpectedCount & " expected: " & ListText
        End If
        ReDim Values(1 To ExpectedCount)
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
'==============================================================================
'                                   ListItem
'------------------------------------------------------------------------------
' PURPOSE
'   Return one item of a "|"-delimited list.
'
' INPUTS
'   ListText   Items separated by "|".
'   Position   Item number, 1-based; 1 <= Position <= item count.
'
' RETURNS
'   The item text.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Items()     As String    'Items of the list

        Items = Split(ListText, "|")
        ListItem = Items(LBound(Items) + Position - 1)

End Function


'------------------------------------------------------------------------------
'                                    FINDINGS
'------------------------------------------------------------------------------

Public Sub AddFinding( _
    ByRef Findings() As String, _
    ByRef FindingCount As Long, _
    ByVal SheetName As String, _
    ByVal Item As String, _
    ByVal RowText As String, _
    ByVal Reason As String)
'
'==============================================================================
'                                  AddFinding
'------------------------------------------------------------------------------
' PURPOSE
'   Append one finding (a rejected row or an unusable parameter) to the list
'   written to the Checks sheet.
'
' INPUTS
'   Findings      Matrix (1 To 4, 1 To n): sheet, item, row, reason; grown
'                 here. Undimensioned when FindingCount = 0.
'   FindingCount  Findings in use (>= 0); incremented.
'   SheetName     Sheet the finding is about.
'   Item          PositionID or parameter code.
'   RowText       Sheet row as text, or "" when not applicable.
'   Reason        Message, starting "INVALID: " for a rejected row.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
        FindingCount = FindingCount + 1
        If FindingCount = 1 Then
            ReDim Findings(1 To 4, 1 To 8)
        ElseIf FindingCount > UBound(Findings, 2) Then
            ReDim Preserve Findings(1 To 4, 1 To 2 * UBound(Findings, 2))
        End If
        Findings(1, FindingCount) = SheetName
        Findings(2, FindingCount) = Item
        Findings(3, FindingCount) = RowText
        Findings(4, FindingCount) = Reason

End Sub


'------------------------------------------------------------------------------
'                                    BUCKETS
'------------------------------------------------------------------------------

Public Sub WriteTotalRow( _
    ByRef BucketOutput() As Variant, _
    ByVal BucketCount As Long, _
    ByVal Label As String)
'
'==============================================================================
'                                 WriteTotalRow
'------------------------------------------------------------------------------
' PURPOSE
'   Fill the total row of an equity or commodity Buckets block: the label in
'   the second column, sums from the third, the net column (6) left blank.
'
' INPUTS
'   BucketOutput  (1 To BucketCount + 1, 1 To width); rows 1 to BucketCount
'                 are filled.
'   BucketCount   Bucket rows above the total row (>= 0).
'   Label         Text of the total row.
'
' UPDATED
'   2026-11-14
'==============================================================================
'
    Dim Total       As Double    'Column sum
    Dim m           As Long      'Bucket row
    Dim c           As Long      'Column

        BucketOutput(BucketCount + 1, 2) = Label
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


'------------------------------------------------------------------------------
'                                  FINGERPRINT
'------------------------------------------------------------------------------

Public Function Fingerprint( _
    ByRef Values() As Variant) _
    As String
'
'==============================================================================
'                                  Fingerprint
'------------------------------------------------------------------------------
' PURPOSE
'   Return a fingerprint of a vector of values: the same values in the same
'   order give the same text, and a changed, added or removed value gives
'   another one, short of a collision of two 31-bit hashes. It tells a
'   change; it is not a cryptographic digest.
'
' INPUTS
'   Values   Vector whose items are cell values, or arrays of cell values of
'            one or two dimensions.
'
' RETURNS
'   FINGERPRINT_VERSION, "-" and sixteen hexadecimal digits.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
    Dim First       As Double    'First running hash
    Dim Second      As Double    'Second running hash
    Dim Table       As Variant   'One array of Values
    Dim i           As Long      'Item of Values
    Dim r           As Long      'Row of Table
    Dim c           As Long      'Column of Table

        First = HASH_BASE_1
        Second = HASH_BASE_2
        For i = LBound(Values) To UBound(Values)
            If Not IsArray(Values(i)) Then
                AddToHash ValueKey(Values(i)), First, Second
            Else
                Table = Values(i)
                'The rows of the array, so that a row moved to another table
                'does not leave the sequence of values as it was.
                AddToHash "A" & CStr(UBound(Table, 1) - LBound(Table, 1) + 1), First, Second
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
        Fingerprint = FINGERPRINT_VERSION & "-" & HexWord(First) & HexWord(Second)

End Function


Private Function HasColumns( _
    ByRef Table As Variant) _
    As Boolean
'
'==============================================================================
'                                  HasColumns
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether an array has a second dimension.
'
' INPUTS
'   Table   Array of one or two dimensions.
'
' ERROR POLICY
'   The error of asking an array of one dimension for its second bound is
'   contained on the single statement that can raise it.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
    Dim Upper       As Long      'Upper bound of the second dimension

        Upper = -1
        On Error Resume Next
        Upper = UBound(Table, 2)
        On Error GoTo 0
        HasColumns = (Upper >= 0)

End Function


Private Function ValueKey( _
    ByVal Value As Variant) _
    As String
'
'==============================================================================
'                                   ValueKey
'------------------------------------------------------------------------------
' PURPOSE
'   Return the text a cell value enters a fingerprint with: its kind and its
'   content, the same in every regional setting.
'
' INPUTS
'   Value   Cell value.
'
' RETURNS
'   "E" for an error value, "_" for an empty cell, "S", the length, ":" and
'   the text for a text, "B0" or "B1" for a Boolean, "N" and the number for
'   a number or a date, and "?" with the type name for anything else.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
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
            'Str$ writes a decimal point whatever the regional setting.
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
'==============================================================================
'                                   AddToHash
'------------------------------------------------------------------------------
' PURPOSE
'   Add a text and a separator to two running polynomial hashes.
'
' INPUTS
'   Text     Text to add; may be empty.
'   First    Hash modulo HASH_MODULUS_1, updated.
'   Second   Hash modulo HASH_MODULUS_2, updated.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
    Dim Code        As Long      'Code unit of one character
    Dim p           As Long      'Position in Text

        For p = 1 To Len(Text) + 1
            If p > Len(Text) Then
                Code = HASH_SEPARATOR
            Else
                Code = AscW(Mid$(Text, p, 1))
                'AscW returns a signed Integer.
                If Code < 0 Then Code = Code + CODE_UNITS
            End If
            'A product stays far below 2^53, so the Double arithmetic is exact.
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
'==============================================================================
'                                    HexWord
'------------------------------------------------------------------------------
' PURPOSE
'   Return a whole number from 0 to 2^31 - 1 as eight hexadecimal digits.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
        HexWord = Right$("0000000" & Hex$(CLng(Value)), 8)

End Function
