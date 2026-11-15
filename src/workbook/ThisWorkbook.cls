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
'==============================================================================
' MODULE: ThisWorkbook
'------------------------------------------------------------------------------
' PURPOSE
'   Event procedures of the K-FRTB-SSA workbook: check, when the workbook is
'   opened and before it is saved, that the results shown are those of the
'   inputs on the sheets.
'
' PUBLIC SURFACE
'   None: event procedures only.
'
' DEPENDENCIES
'   M_Main.
'
' STATE OWNERSHIP
'   None of its own. M_Main.CheckStatus_OnEvent writes the status line of
'   Results when it must change.
'
' ERROR POLICY
'   Raises nothing and shows no message: M_Main.CheckStatus_OnEvent contains
'   every error, so that an unattended open or save is never stopped.
'
' COMPATIBILITY
'   Excel for Windows, VBA7; no references beyond the defaults.
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
    'Require explicit declarations.
    Option Explicit


'------------------------------------------------------------------------------
'                               EVENT PROCEDURES
'------------------------------------------------------------------------------

Private Sub Workbook_Open()
'
'==============================================================================
'                                 Workbook_Open
'------------------------------------------------------------------------------
' PURPOSE
'   Mark the results OUTDATED when an input changed after the last run and
'   the workbook was saved without a check, as an application that does not
'   run this module can do.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
        M_Main.CheckStatus_OnEvent

End Sub


Private Sub Workbook_BeforeSave( _
    ByVal SaveAsUI As Boolean, _
    ByRef Cancel As Boolean)
'
'==============================================================================
'                              Workbook_BeforeSave
'------------------------------------------------------------------------------
' PURPOSE
'   Mark the results OUTDATED before the workbook is saved when an input
'   changed after the last run, so that the saved file does not show them
'   as current.
'
' INPUTS
'   SaveAsUI   True when Excel shows the Save As dialog; not used.
'   Cancel     Left as it is: the save is never cancelled.
'
' UPDATED
'   2026-11-15
'==============================================================================
'
        M_Main.CheckStatus_OnEvent

End Sub
