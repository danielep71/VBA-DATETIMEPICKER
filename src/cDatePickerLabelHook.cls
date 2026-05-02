VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDatePickerLabelHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'------------------------------------------------------------------------------
' CLASS: cDatePickerLabelHook
'------------------------------------------------------------------------------
' PURPOSE
'   Provides a generic event hook for runtime-created DatePicker labels
'
' WHY THIS EXISTS
'   Runtime-created MSForms labels do not expose normal UserForm event
'   procedures. A WithEvents class is required to capture their Click and
'   MouseMove events
'
' INPUTS
'   None at class level
'
' RETURNS
'   Nothing at class level
'
' BEHAVIOR
'   Supports:
'     - optional click-action routing
'     - day-label hover routing
'     - header-label hover routing
'     - picker-panel item hover routing
'     - footer-label hover routing
'     - labels with no hover behavior
'
' ERROR POLICY
'   Raises descriptive runtime errors during initialization and event routing
'
' DEPENDENCIES
'   MSForms.Label
'
'   OwnerForm.UF_PickerPanel_HandleAction
'   OwnerForm.UF_DayCell_HoverApply
'   OwnerForm.UF_Header_HoverApply
'   OwnerForm.UF_PickerPanel_HoverApply
'   OwnerForm.UF_Footer_HoverApply
'
' NOTES
'   Instances of this class must be stored in a module-level Collection in the
'   UserForm. Otherwise they are destroyed and events stop firing
'
'   The owner form is stored as Object to keep this class generic and reusable
'   across UserForms that expose the expected public routing methods
'
'   The owning UserForm should release the hook collections during termination
'
'   The optional HoverMode default is the literal string "NONE" rather than the
'   private class constant DP_HOVER_NONE because VBA can be fragile with
'   class-level constants in public optional-argument defaults
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const DP_HOVER_NONE         As String = "NONE"      'No hover behavior
    Private Const DP_HOVER_DAY          As String = "DAY"       'Day-label hover behavior
    Private Const DP_HOVER_HEADER       As String = "HEADER"    'Header-label hover behavior
    Private Const DP_HOVER_PICKER       As String = "PICKER"    'Picker-panel hover behavior
    Private Const DP_HOVER_FOOTER       As String = "FOOTER"    'Footer-label hover behavior

'------------------------------------------------------------------------------
' PRIVATE VARIABLES
'------------------------------------------------------------------------------
    Private WithEvents mLabel           As MSForms.Label        'Hooked runtime label
Attribute mLabel.VB_VarHelpID = -1
    Private mOwnerForm                  As Object               'Owner UserForm
    Private mActionName                 As String               'Click action name
    Private mHoverMode                  As String               'Hover behavior mode

'
'------------------------------------------------------------------------------
'
'                                INITIALIZATION
'
'------------------------------------------------------------------------------
'

Public Sub Initialize( _
    ByVal OwnerForm As Object, _
    ByVal TargetLabel As MSForms.Label, _
    Optional ByVal ActionName As String = vbNullString, _
    Optional ByVal HoverMode As String = "NONE")

'
'------------------------------------------------------------------------------
'                           INITIALIZE LABEL HOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Connects one runtime-created MSForms.Label to this generic DatePicker event
'   hook
'
' WHY THIS EXISTS
'   Runtime-created MSForms labels do not expose normal UserForm event
'   procedures unless their events are captured through a WithEvents class
'
' INPUTS
'   OwnerForm
'     UserForm that owns the label and exposes the routed action / hover methods
'
'   TargetLabel
'     Runtime-created MSForms.Label to monitor
'
'   ActionName
'     Optional click action routed to the owner form
'
'   HoverMode
'     Optional hover behavior
'       - NONE
'       - DAY
'       - HEADER
'       - PICKER
'       - FOOTER
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the owner form, target label, click action, and hover mode
'
' ERROR POLICY
'   Raises a descriptive runtime error if OwnerForm or TargetLabel is missing,
'   or if HoverMode is unsupported
'
' DEPENDENCIES
'   MSForms.Label
'
' NOTES
'   Instances of this class must be stored in a module-level Collection in the
'   UserForm. Otherwise they are destroyed and events stop firing
'
'   The optional HoverMode default is the literal string "NONE" rather than the
'   private class constant DP_HOVER_NONE because VBA can be fragile with
'   class-level constants in public optional-argument defaults
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "cDatePickerLabelHook.Initialize"    'Current procedure name

    Dim NormalizedHoverMode    As String                                        'Normalized hover mode

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing owner form
        If OwnerForm Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "OwnerForm cannot be Nothing."
        End If

    'Reject a missing target label
        If TargetLabel Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, "TargetLabel cannot be Nothing."
        End If

'------------------------------------------------------------------------------
' NORMALIZE HOVER MODE
'------------------------------------------------------------------------------
    'Normalize the requested hover mode
        NormalizedHoverMode = UCase$(Trim$(CStr(HoverMode)))

    'Default blank hover mode to NONE
        If Len(NormalizedHoverMode) = 0 Then
            NormalizedHoverMode = DP_HOVER_NONE
        End If

'------------------------------------------------------------------------------
' VALIDATE HOVER MODE
'------------------------------------------------------------------------------
    'Reject unsupported hover modes
        Select Case NormalizedHoverMode

            Case DP_HOVER_NONE, _
                 DP_HOVER_DAY, _
                 DP_HOVER_HEADER, _
                 DP_HOVER_PICKER, _
                 DP_HOVER_FOOTER

            Case Else
                Err.Raise vbObjectError + 515, PROC_NAME, _
                    "Unsupported HoverMode '" & CStr(HoverMode) & _
                    "'. HoverMode must be NONE, DAY, HEADER, PICKER, or FOOTER."

        End Select

'------------------------------------------------------------------------------
' STORE REFERENCES
'------------------------------------------------------------------------------
    'Store the owner form reference
        Set mOwnerForm = OwnerForm

    'Store the target label reference
        Set mLabel = TargetLabel

    'Store the click action name
        mActionName = Trim$(CStr(ActionName))

    'Store the normalized hover mode
        mHoverMode = NormalizedHoverMode

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub

Private Sub mLabel_Click()

'
'==============================================================================
'                           LABEL CLICK
'------------------------------------------------------------------------------
' PURPOSE
'   Routes the label click event to the owner form
'
' WHY THIS EXISTS
'   Runtime-created DatePicker labels use a shared hook class to route click
'   behavior back to the UserForm
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls UF_PickerPanel_HandleAction when this hook has a non-empty action name
'
' ERROR POLICY
'   Raises a descriptive runtime error if the owner form is missing, if the
'   target label is missing, or if click-action routing fails
'
' DEPENDENCIES
'   OwnerForm.UF_PickerPanel_HandleAction
'
' NOTES
'   This routine is only the event router
'
'   Errors raised here normally originate inside the owner form action handler
'
' UPDATED
'   2026-04-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerLabelHook.mLabel_Click"     'Current procedure name

    Dim RoutedActionName        As String       'Action name snapshot
    Dim RoutedLabelName         As String       'Label name snapshot
    Dim RoutedOwnerForm         As Object       'Owner form snapshot
    
    Dim InnerErrNumber          As Long         'Inner error number
    Dim InnerErrSource          As String       'Inner error source
    Dim InnerErrDescription     As String       'Inner error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT ACTION
'------------------------------------------------------------------------------
    'Capture the action name before routing
        RoutedActionName = Trim$(CStr(mActionName))

'------------------------------------------------------------------------------
' EXIT IF NO ACTION IS ASSIGNED
'------------------------------------------------------------------------------
    'Exit if this label has no click action
        If Len(RoutedActionName) = 0 Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' VALIDATE STATE
'------------------------------------------------------------------------------
    'Reject a missing owner form
        If mOwnerForm Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "OwnerForm is not initialized."
        End If
    'Reject a missing target label
        If mLabel Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, "Target label is not initialized."
        End If

'------------------------------------------------------------------------------
' SNAPSHOT REFERENCES
'------------------------------------------------------------------------------
    'Capture the owner form reference before routing
        Set RoutedOwnerForm = mOwnerForm
    'Capture the clicked label name before routing
        RoutedLabelName = CStr(mLabel.Name)

'------------------------------------------------------------------------------
' ROUTE CLICK ACTION
'------------------------------------------------------------------------------
    'Route the click action to the owner form
        RoutedOwnerForm.UF_PickerPanel_HandleAction RoutedActionName

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the inner error number
        InnerErrNumber = Err.Number
    'Capture the inner error source
        InnerErrSource = Err.Source
    'Capture the inner error description
        InnerErrDescription = Err.Description
    'Raise a diagnostic error without losing the inner source
        Err.Raise InnerErrNumber, _
            PROC_NAME & " | Label=" & RoutedLabelName & " | Action=" & RoutedActionName, _
            "InnerSource=[" & InnerErrSource & "] | InnerDescription=[" & InnerErrDescription & "]"

End Sub
Private Sub mLabel_MouseMove( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

'
'------------------------------------------------------------------------------
'                           LABEL MOUSE MOVE
'------------------------------------------------------------------------------
' PURPOSE
'   Routes label hover behavior to the owner form
'
' WHY THIS EXISTS
'   Runtime-created DatePicker labels use a shared hook class to route hover
'   behavior back to the UserForm
'
' INPUTS
'   Button
'     Standard MSForms mouse button state
'
'   Shift
'     Standard MSForms shift-key state
'
'   X
'     Mouse X position
'
'   Y
'     Mouse Y position
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Routes hover behavior according to mHoverMode
'
' ERROR POLICY
'   Best-effort hover routing. Suppresses hover errors so MouseMove cannot
'   interrupt UserForm interaction
'
' DEPENDENCIES
'   OwnerForm.UF_DayCell_HoverApply
'   OwnerForm.UF_Header_HoverApply
'   OwnerForm.UF_PickerPanel_HoverApply
'   OwnerForm.UF_Footer_HoverApply
'
' NOTES
'   Labels with HoverMode = NONE safely exit without routing
'
'   Hover errors are intentionally written to the Immediate Window instead of
'   being raised outward
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "cDatePickerLabelHook.mLabel_MouseMove"    'Current procedure name

    Dim RoutedLabelName        As String                                             'Label name snapshot
    Dim RoutedHoverMode        As String                                             'Hover mode snapshot

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress hover-routing failures
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' EXIT IF NO HOVER IS ASSIGNED
'------------------------------------------------------------------------------
    'Exit when this label has no hover behavior
        If mHoverMode = DP_HOVER_NONE Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' VALIDATE STATE
'------------------------------------------------------------------------------
    'Exit safely when the owner form is unavailable
        If mOwnerForm Is Nothing Then
            Exit Sub
        End If

    'Exit safely when the target label is unavailable
        If mLabel Is Nothing Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' SNAPSHOT STATE
'------------------------------------------------------------------------------
    'Capture the label name before routing
        RoutedLabelName = CStr(mLabel.Name)

    'Capture the hover mode before routing
        RoutedHoverMode = CStr(mHoverMode)

'------------------------------------------------------------------------------
' ROUTE HOVER ACTION
'------------------------------------------------------------------------------
    'Route hover behavior according to the configured hover mode
        Select Case RoutedHoverMode

            Case DP_HOVER_DAY
                mOwnerForm.UF_DayCell_HoverApply RoutedLabelName

            Case DP_HOVER_HEADER
                mOwnerForm.UF_Header_HoverApply RoutedLabelName

            Case DP_HOVER_PICKER
                mOwnerForm.UF_PickerPanel_HoverApply RoutedLabelName

            Case DP_HOVER_FOOTER
                mOwnerForm.UF_Footer_HoverApply RoutedLabelName

            Case Else
                Debug.Print PROC_NAME & ": unsupported hover mode '" & RoutedHoverMode & "'"

        End Select

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful hover routing
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Write hover diagnostics without interrupting the UserForm
        Debug.Print PROC_NAME & _
            " | Label=" & RoutedLabelName & _
            " | HoverMode=" & RoutedHoverMode & _
            " | Error=" & CStr(Err.Number) & _
            " | " & Err.Description

End Sub

'
'------------------------------------------------------------------------------
'
'                                  CLEANUP
'
'------------------------------------------------------------------------------
'

Private Sub Class_Terminate()

'
'------------------------------------------------------------------------------
'                           TERMINATE CLASS
'------------------------------------------------------------------------------
' PURPOSE
'   Releases references held by the label hook
'
' WHY THIS EXISTS
'   The hook stores references to the owner form and target label while the
'   DatePicker is open. These references should be released when the hook is
'   destroyed
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Releases object references and clears stored routing state
'
' ERROR POLICY
'   Suppresses cleanup errors
'
' DEPENDENCIES
'   None
'
' NOTES
'   The owning UserForm should still release the hook collection during
'   UserForm_Terminate
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAN UP
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next
    'Release the target label reference
        Set mLabel = Nothing
    'Release the owner form reference
        Set mOwnerForm = Nothing
    'Clear the click action name
        mActionName = vbNullString
    'Clear the hover mode
        mHoverMode = vbNullString
    'Restore normal error handling
        On Error GoTo 0

End Sub

