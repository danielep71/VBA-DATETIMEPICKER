VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: ThisWorkbook (add-in)
'------------------------------------------------------------------------------
' PURPOSE
'   Workbook lifecycle code of the DATETIMEPICKER add-in package: starts the
'   runtime on open, stops it before close, and clears transient UI before
'   printing.
'
' WHY THIS EXISTS
'   A workbook's document module lives inside the Office package and cannot be
'   imported as a component. Keeping its code here makes the add-in buildable
'   from tracked source only (RELEASING.md section 7, #211).
'
' HOW IT IS USED
'   tools/Build-Addin.ps1 writes everything after the Attribute lines into the
'   new workbook's own ThisWorkbook module. Do not import this file in the VBE:
'   File > Import would create a class module named ThisWorkbook1 instead.
'
' ORIGIN
'   The procedures are those of the published DATETIMEPICKER.v1.2.2.xlam
'   ThisWorkbook module, read with tools/inspect_vba_package.py; only trailing
'   whitespace changed.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------
'

Private Sub Workbook_Open()

'
'------------------------------------------------------------------------------
'                           WORKBOOK OPEN
'------------------------------------------------------------------------------
' PURPOSE
'   Starts the DatePicker runtime, hides the demo worksheet, and preloads the
'   DatePicker UserForm for immediate later display
'
' WHY THIS EXISTS
'   The DatePicker manager and integrations must be initialized when the
'   workbook opens, while the first UserForm load cost should be paid during
'   startup rather than during the first user click
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Hides the demo sheet on a best-effort basis, starts DatePicker runtime
'   integrations through DP_Start, and preloads the DatePicker form hidden
'
' ERROR POLICY
'   Logs startup failures to the Immediate Window and does not block workbook
'   opening
'
' DEPENDENCIES
'   DP_DemoSheet_HideVeryHidden
'   DP_Start
'   DP_Preload
'
' NOTES
'   Do not duplicate worksheet SelectionChange logic here
'
'   The manager owns Application-level events
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' HIDE DEMO SHEET
'------------------------------------------------------------------------------
    'Hide the demo worksheet during normal workbook startup
        DP_DemoSheet_HideVeryHidden

'------------------------------------------------------------------------------
' START DATEPICKER
'------------------------------------------------------------------------------
    'Start DatePicker runtime integrations
        DP_Start

'------------------------------------------------------------------------------
' PRELOAD DATEPICKER FORM
'------------------------------------------------------------------------------
    'Preload the DatePicker form hidden for faster first display
        DP_Preload

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Write startup diagnostics without blocking workbook open
        Debug.Print "ThisWorkbook.Workbook_Open" & _
            " | Error=" & VBA.CStr(Err.Number) & _
            " | " & Err.Description

    'Clear the handled startup error
        Err.Clear

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

'
'------------------------------------------------------------------------------
'                           WORKBOOK BEFORE CLOSE
'------------------------------------------------------------------------------
' PURPOSE
'   Stops the DatePicker runtime and hides the demo worksheet before close
'
' WHY THIS EXISTS
'   DatePicker uses application-wide and transient UI artifacts that should not
'   survive after the workbook project is closed. The demo sheet should also
'   return to its normal very-hidden state
'
' INPUTS
'   Cancel
'     Excel close-cancel flag
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Hides the demo worksheet and delegates DatePicker teardown to DP_Stop
'
' ERROR POLICY
'   Best-effort workbook shutdown. Suppresses teardown failures and does not
'   modify Cancel
'
' DEPENDENCIES
'   DP_DemoSheet_HideVeryHidden
'   DP_Stop
'
' NOTES
'   This routine does not set Cancel
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress shutdown errors
        On Error Resume Next

'------------------------------------------------------------------------------
' RESPECT EXISTING CANCELLATION
'------------------------------------------------------------------------------
    'Exit if another process has already cancelled the workbook close
        If Cancel Then Exit Sub

'------------------------------------------------------------------------------
' HIDE DEMO SHEET
'------------------------------------------------------------------------------
    'Return the demo worksheet to very hidden state
        DP_DemoSheet_HideVeryHidden

'------------------------------------------------------------------------------
' STOP DATEPICKER
'------------------------------------------------------------------------------
    'Stop DatePicker runtime integrations
        DP_Stop

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Clear any suppressed shutdown error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)

'
'------------------------------------------------------------------------------
'                           WORKBOOK BEFORE PRINT
'------------------------------------------------------------------------------
' PURPOSE
'   Clears transient DatePicker UI before printing
'
' WHY THIS EXISTS
'   Runtime UserForms and worksheet icon shapes should not interfere with print
'   output or page-preview behavior
'
' INPUTS
'   Cancel
'     Excel print-cancel flag
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Closes the DatePicker form and purges grid icons before print
'
' ERROR POLICY
'   Best-effort print cleanup
'
'   Suppresses cleanup failures and does not modify Cancel
'
' DEPENDENCIES
'   DP_Close
'   M_GridIcon_PurgeAll
'
' NOTES
'   cDatePickerManager already handles Application-level WorkbookBeforePrint
'   when the manager is alive. This workbook event is an additional safe fallback
'
' UPDATED
'   2026-05-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress print-cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' RESPECT EXISTING CANCELLATION
'------------------------------------------------------------------------------
    'Exit if another process has already cancelled the print
        If Cancel Then Exit Sub

'------------------------------------------------------------------------------
' CLEAR TRANSIENT UI
'------------------------------------------------------------------------------
    'Close any visible DatePicker form
        DP_Close
    'Purge in-grid icon shapes before print
        M_GridIcon_PurgeAll

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub
