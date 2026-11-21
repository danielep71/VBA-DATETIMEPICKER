Attribute VB_Name = "M_DP_Keyboard"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_DP_Keyboard
'------------------------------------------------------------------------------
' PURPOSE
'   Registers and removes the Ctrl+Shift+D shortcut through
'   Application.OnKey
'
' WHY THIS EXISTS
'   OnKey is application-wide and has no getter. One module owns the key
'   and every registration and removal of it (#57)
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - M_KeyboardShortcut_Update and M_KeyboardShortcut_Remove
'     - M_KeyboardShortcut_TryRemove, the removal used by lifecycle
'       cleanup, which reports a failure instead of hiding it
'
' ERROR POLICY
'   Each procedure documents its own policy
'
' DEPENDENCIES
'   Application.OnKey
'   M_DP_MacroNames (M_GetQualifiedMacroName)
'   M_DP_Core (gDP_EnableKeyboardShortcut, M_Settings_EnsureLoaded)
'
' NOTES
'   The shortcut runs the facade callback DP_OpenForActiveCell
'
'   Option Private Module keeps every procedure out of other VBA projects
'   and out of the macro dialog. Procedures are Public only where another
'   module of the project calls them
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const DP_KEYBOARD_SHORTCUT_KEY As String = "^+d"    'Ctrl + Shift + D


'
'------------------------------------------------------------------------------
'
'                              KEYBOARD SHORTCUT
'
'------------------------------------------------------------------------------
'

Public Sub M_KeyboardShortcut_Update()

'
'------------------------------------------------------------------------------
'                         UPDATE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Registers or removes the DatePicker keyboard shortcut according to settings
'
' WHY THIS EXISTS
'   The keyboard shortcut is one of three interactive entry points, alongside the
'   in-grid icon and the right-click menu. Disabling the other two no longer
'   enables it: Application.OnKey is a session-wide binding, and the component
'   does not take it on behalf of a user who did not ask for it
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads settings if needed, registers Ctrl + Shift + D when enabled, and
'   restores the key to Excel when disabled
'
' ERROR POLICY
'   Raises a descriptive runtime error if shortcut synchronization fails
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_KeyboardShortcut_Register
'   M_KeyboardShortcut_Remove
'   gDP_EnableKeyboardShortcut
'
' NOTES
'   Application.OnKey is application-wide for the current Excel session, so it
'   must be removed during teardown
'
'   This routine owns only the setting-driven synchronization policy
'
'   The actual Application.OnKey assignment and removal remain delegated to the
'   dedicated register / remove routines
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_KeyboardShortcut_Update"

    Dim HandlerStep             As String       'Current handler step for diagnostics
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"
    'Ensure settings are available before reading the keyboard shortcut flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' SYNCHRONIZE SHORTCUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Synchronize keyboard shortcut"
    'Register the keyboard shortcut when the feature is enabled
        If gDP_EnableKeyboardShortcut Then
            M_KeyboardShortcut_Register
    'Otherwise restore the key to Excel default handling
        Else
            M_KeyboardShortcut_Remove
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Keyboard shortcut synchronization failed: " & ErrorDescription

End Sub

Private Sub M_KeyboardShortcut_Register()

'
'------------------------------------------------------------------------------
'                         REGISTER KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Registers the DatePicker keyboard shortcut for the current Excel session
'
' WHY THIS EXISTS
'   Registration happens only when the user enabled the shortcut. Nothing else
'   turns it on, because taking a session-wide key binding is not a decision the
'   component makes on the user's behalf
'
'   Excel exposes no getter for Application.OnKey, so a binding this displaces
'   cannot be captured and cannot be restored. That limitation is documented
'   rather than worked around
'
'   Application.OnKey requires an explicit public macro callback and the
'   assignment is not persisted by Excel across sessions
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Assigns Ctrl + Shift + D to the workbook-qualified DP_OpenForActiveCell
'   callback
'
' ERROR POLICY
'   Raises a descriptive runtime error if the shortcut cannot be registered
'
' DEPENDENCIES
'   Application.OnKey
'   M_GetQualifiedMacroName
'   DP_OpenForActiveCell
'   DP_KEYBOARD_SHORTCUT_KEY
'
' NOTES
'   The callback is workbook-qualified so the shortcut resolves to this project
'
'   Application.OnKey is application-wide for the current Excel session
'
'   The shortcut should be removed during teardown through
'   M_KeyboardShortcut_Remove
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_KeyboardShortcut_Register"

    Dim CallbackMacroName       As String       'Workbook-qualified callback macro name
    Dim HandlerStep             As String       'Current handler step for diagnostics
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' RESOLVE CALLBACK MACRO NAME
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve callback macro name"
    'Build the workbook-qualified DatePicker launcher macro name
        CallbackMacroName = M_GetQualifiedMacroName("DP_OpenForActiveCell")

'------------------------------------------------------------------------------
' REGISTER SHORTCUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Register keyboard shortcut"
    'Assign Ctrl + Shift + D to the DatePicker public launcher
        Application.OnKey DP_KEYBOARD_SHORTCUT_KEY, CallbackMacroName

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Keyboard shortcut registration failed: " & ErrorDescription

End Sub

Public Sub M_KeyboardShortcut_Remove()

'
'------------------------------------------------------------------------------
'                         REMOVE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Restores the DatePicker keyboard shortcut to Excel default handling
'
' WHY THIS EXISTS
'   Application.OnKey assignments are application-wide. If the workbook or add-in
'   is closed, reset, or unloaded, the shortcut must not remain bound to a macro
'   that may no longer exist
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Restores Ctrl + Shift + D to Excel's default handling
'
' ERROR POLICY
'   Best-effort cleanup. Does not raise outward
'
'   Shortcut-removal failures are suppressed and written to the Immediate Window
'   for diagnostics
'
' DEPENDENCIES
'   Application.OnKey
'   DP_KEYBOARD_SHORTCUT_KEY
'
' NOTES
'   This restores Excel's default handling rather than any binding the DatePicker
'   displaced. Excel exposes no getter for Application.OnKey, so the displaced
'   assignment cannot be captured and cannot be put back. Of the three possible
'   teardown behaviors this is the least damaging: binding the key to an empty
'   macro swallows it, and leaving the DatePicker callback in place points it at
'   a project that may be unloading
'
'   Calling Application.OnKey with only the key argument restores normal Excel
'   behavior for that key combination
'
'   This routine intentionally does not call M_Settings_EnsureLoaded because it
'   is a teardown-safe cleanup routine
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_KeyboardShortcut_Remove"

    Dim RemoveErrNumber         As Long         'Captured shortcut-removal error number
    Dim RemoveErrDescription    As String       'Captured shortcut-removal error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors because shortcut removal is best-effort
        On Error Resume Next

'------------------------------------------------------------------------------
' REMOVE SHORTCUT
'------------------------------------------------------------------------------
    'Restore the key to Excel default handling
        Application.OnKey DP_KEYBOARD_SHORTCUT_KEY

'------------------------------------------------------------------------------
' CAPTURE DIAGNOSTICS
'------------------------------------------------------------------------------
    'Capture shortcut-removal error number
        RemoveErrNumber = Err.Number
    'Capture shortcut-removal error description
        RemoveErrDescription = Err.Description
    'Clear any suppressed shortcut-removal error
        Err.Clear

'------------------------------------------------------------------------------
' WRITE DIAGNOSTICS
'------------------------------------------------------------------------------
    'Write diagnostics only when shortcut removal failed
        If RemoveErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=Application.OnKey" & _
                " | Key=" & DP_KEYBOARD_SHORTCUT_KEY & _
                " | Error=" & VBA.CStr(RemoveErrNumber) & _
                " | " & RemoveErrDescription
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                              VERIFIED TEARDOWN
'
'------------------------------------------------------------------------------
'

Public Function M_KeyboardShortcut_TryRemove( _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'                         REMOVE THE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Restores Excel's default handling of the DatePicker shortcut key
'
' WHY THIS EXISTS
'   The shortcut is an application-wide Application.OnKey binding and must not
'   outlive the runtime that registered it
'
' INPUTS
'   ErrorNumber, ErrorDescription
'     Receive the failure when the OnKey call raises
'
' RETURNS
'   True when the OnKey call completed
'
' BEHAVIOR
'   Calls Application.OnKey with the shortcut key and no procedure, which hands
'   the key back to Excel
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   DP_KEYBOARD_SHORTCUT_KEY
'
' NOTES
'   This is the one cleanup step that cannot verify its own outcome. Excel
'   exposes no getter for Application.OnKey, so success here means only that the
'   call did not raise, not that the binding is gone. Every other step in this
'   transaction reads its result back
'
'   That is the same limitation #42 recorded for the registration side, and it is
'   why the keyboard shortcut is covered by manual validation rather than by a
'   harness assertion
'
'   Moved from M_DP_Core by #57, where it was named
'   M_Lifecycle_TryRemoveKeyboardShortcut, so the module that owns this
'   integration's state also verifies its teardown
'
'   Public only so the lifecycle cleanup in M_DP_Core can call it
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

    On Error GoTo Failed
    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString

    Excel.Application.OnKey DP_KEYBOARD_SHORTCUT_KEY
    M_KeyboardShortcut_TryRemove = True
    Exit Function

Failed:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    Err.Clear
    M_KeyboardShortcut_TryRemove = False

End Function
