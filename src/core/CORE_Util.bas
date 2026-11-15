Attribute VB_Name = "CORE_Util"
'==============================================================================
' MODULE: CORE_Util
'------------------------------------------------------------------------------
' PURPOSE
'   Shared helpers of the K-FRTB-SSA core: classify cell values (number, date,
'   text), normalise codes, look codes up in arrays, parse the delimited
'   lists of CORE_Config, collect Checks findings and total a Buckets block.
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
