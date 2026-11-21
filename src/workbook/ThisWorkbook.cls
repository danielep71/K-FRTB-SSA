VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: ThisWorkbook
'------------------------------------------------------------------------------
' PURPOSE
'   Event procedures of the K-FRTB-SSA workbook: check, when the workbook is
'   opened and before it is saved, that the results shown are those of the
'   inputs on the sheets
'
' WHY THIS EXISTS
'   An application that does not run this module can save the workbook after
'   an input changed without a check. These event procedures call the check
'   of M_Main on open and before save, so that the status line of Results
'   shows OUTDATED instead of a result that no longer matches the inputs
'
' INPUTS
'   None at class level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Workbook_Open and Workbook_BeforeSave each call M_Main.CheckStatus_OnEvent,
'   once per event. The module has no other procedure
'
' ERROR POLICY
'   Raises nothing and shows no message: M_Main.CheckStatus_OnEvent contains
'   every error, so that an unattended open or save is never stopped
'
' DEPENDENCIES
'   M_Main (CheckStatus_OnEvent)
'
' NOTES
'   Public surface: None; event procedures only
'
'   The module owns no state of its own. M_Main.CheckStatus_OnEvent writes the
'   status line of Results when it must change
'
'   Option Explicit requires explicit declarations
'
'   Compatibility: Excel for Windows, VBA7; no references beyond the defaults
'
'   Author: Daniele Penza
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------


'
'------------------------------------------------------------------------------
'
'                               EVENT PROCEDURES
'
'------------------------------------------------------------------------------
'

Private Sub Workbook_Open()
'
'------------------------------------------------------------------------------
'                              CHECK ON OPEN
'------------------------------------------------------------------------------
' PURPOSE
'   Marks the results OUTDATED when an input changed after the last run and
'   the workbook was saved without a check, as an application that does not
'   run this module can do
'
' WHY THIS EXISTS
'   Excel raises Workbook_Open when the workbook is opened, which is the first
'   point at which the saved status line can be compared with the inputs
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls M_Main.CheckStatus_OnEvent once
'
' ERROR POLICY
'   Raises nothing: M_Main.CheckStatus_OnEvent contains every error
'
' DEPENDENCIES
'   M_Main (CheckStatus_OnEvent)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK STATUS
'------------------------------------------------------------------------------
    'Mark the results OUTDATED when an input changed after the last run
        M_Main.CheckStatus_OnEvent

End Sub

Private Sub Workbook_BeforeSave( _
    ByVal SaveAsUI As Boolean, _
    ByRef Cancel As Boolean)
'
'------------------------------------------------------------------------------
'                              CHECK BEFORE SAVE
'------------------------------------------------------------------------------
' PURPOSE
'   Marks the results OUTDATED before the workbook is saved when an input
'   changed after the last run, so that the saved file does not show them as
'   current
'
' WHY THIS EXISTS
'   Excel raises Workbook_BeforeSave before every save, so the saved file
'   carries a status line that matches its inputs
'
' INPUTS
'   SaveAsUI
'     True when Excel shows the Save As dialog; not used
'
'   Cancel
'     Left as it is: the save is never cancelled
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls M_Main.CheckStatus_OnEvent once, before Excel saves the workbook
'
' ERROR POLICY
'   Raises nothing: M_Main.CheckStatus_OnEvent contains every error, so that a
'   save is never stopped
'
' DEPENDENCIES
'   M_Main (CheckStatus_OnEvent)
'
' NOTES
'   None
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK STATUS
'------------------------------------------------------------------------------
    'Mark the results OUTDATED when an input changed after the last run
        M_Main.CheckStatus_OnEvent

End Sub
