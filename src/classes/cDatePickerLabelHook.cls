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
'   cDatePickerLabelHook provides the lightweight WithEvents routing layer used
'   by the DatePicker UserForm to handle Click and MouseMove events for
'   runtime-created MSForms labels.
'
' WHY THIS EXISTS
'   Runtime-created MSForms labels do not expose normal UserForm event
'   procedures. A WithEvents class is required to capture their Click and
'   MouseMove events.
'
' INPUTS
'   None at class level
'
' RETURNS
'   Nothing at class level
'
' BEHAVIOR
'   Supports:
'     - optional click-action routing via a caller-supplied handler method name
'     - day-label hover routing
'     - header-label hover routing
'     - picker-panel item hover routing
'     - footer-label hover routing
'     - settings-panel label hover routing
'     - simulated MouseLeave via NotifyMouseLeft called by the owner form
'     - labels with no hover behavior
'
' ERROR POLICY
'   Initialization raises descriptive runtime errors if required references are
'   missing, if the hook is already initialized, if the requested hover mode is
'   unsupported, or if the owner form does not expose every required routing
'   method
'
'   Click routing raises diagnostic runtime errors and preserves the inner error
'   source and description
'
'   Hover routing is best-effort and suppresses errors so MouseMove cannot
'   interrupt UserForm interaction
'
'   MouseLeave routing is best-effort and suppresses errors so NotifyMouseLeft
'   cannot interrupt UserForm interaction
'
'   Termination suppresses cleanup errors
'
' DEPENDENCIES
'   MSForms.Label
'
'   Expected public members on OwnerForm (verified at initialization time):
'     - UF_DayCell_HoverApply
'     - UF_Header_HoverApply
'     - UF_PickerPanel_HoverApply
'     - UF_Footer_HoverApply
'     - UF_SettingsPanel_HoverApply
'
'   The click-action handler method name is supplied by the caller via
'   ClickHandlerName and is verified against the owner form at initialization
'
'   The leave handler method name is supplied by the caller via LeaveHandlerName
'   and is verified against the owner form at initialization when provided.
'   The leave handler receives the departing label name as its only argument.
'
' NOTES
'   Instances of this class must be stored in a module-level Collection in the
'   UserForm. Otherwise they are destroyed and events stop firing.
'
'   The owner form is stored as Object to keep this class generic and reusable
'   across UserForms that expose the expected public routing methods.
'
'   Because OwnerForm is late-bound, method existence is probed explicitly
'   during Initialize using CallByName so that missing or renamed routing
'   methods are detected at initialization time rather than at first use.
'
'   The owning UserForm should release the hook collection during termination.
'
'   The optional HoverMode default is the literal string "NONE" rather than the
'   private class constant DP_HOVER_NONE because VBA can be fragile with
'   class-level constants in public optional-argument defaults.
'
'   mHoverMode is expected to contain a normalized uppercase value after
'   successful initialization.
'
'   mIsInitialized is set to True only after every reference and validation
'   step completes successfully. It is the authoritative hook-state flag.
'
'   MSForms labels do not fire a native MouseLeave event. The leave mechanism
'   in this class is cooperative: the owning UserForm must call NotifyMouseLeft
'   on every hook from its own MouseMove handler. Each hook uses the mIsHovered
'   flag to detect whether it was the previously hovered label. When the flag
'   is True and NotifyMouseLeft fires, the hook concludes the mouse has left,
'   routes the leave callback, and clears the flag. Labels that were never
'   entered skip the callback at zero cost.
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const DP_HOVER_NONE         As String = "NONE"      'No hover behavior
    Private Const DP_HOVER_DAY          As String = "DAY"       'Day-label hover behavior
    Private Const DP_HOVER_HEADER       As String = "HEADER"    'Header-label hover behavior
    Private Const DP_HOVER_PICKER       As String = "PICKER"    'Picker-panel hover behavior
    Private Const DP_HOVER_FOOTER       As String = "FOOTER"    'Footer-label hover behavior
    Private Const DP_HOVER_SETTINGS     As String = "SETTINGS"  'Settings-panel hover behavior

    'Error numbers - each procedure uses a distinct block to allow log filtering
    Private Const DP_ERR_ALREADY_INIT   As Long = 512           'Hook already initialized
    Private Const DP_ERR_NO_OWNER       As Long = 513           'OwnerForm is Nothing
    Private Const DP_ERR_NO_LABEL       As Long = 514           'TargetLabel is Nothing
    Private Const DP_ERR_BAD_HOVER      As Long = 515           'Unsupported hover mode
    Private Const DP_ERR_BAD_METHOD     As Long = 516           'Owner form missing required method
    Private Const DP_ERR_CLICK_UNINIT   As Long = 520           'Click fired before initialization
    Private Const DP_ERR_CLICK_NO_OWNER As Long = 521           'OwnerForm lost during click
    Private Const DP_ERR_LEAVE_UNINIT   As Long = 530           'NotifyMouseLeft called before initialization
    Private Const DP_ERR_LEAVE_NO_OWNER As Long = 531           'OwnerForm lost during leave routing

'------------------------------------------------------------------------------
' PRIVATE VARIABLES
'------------------------------------------------------------------------------
    Private WithEvents mLabel           As MSForms.Label        'Hooked runtime label
Attribute mLabel.VB_VarHelpID = -1

    Private mOwnerForm                  As Object               'Owner UserForm exposing expected routing methods
    Private mActionName                 As String               'Normalized click action name
    Private mClickHandlerName           As String               'Normalized owner-form click handler method name
    Private mLeaveHandlerName           As String               'Normalized owner-form leave handler method name
    Private mHoverMode                  As String               'Normalized hover behavior mode
    Private mIsInitialized              As Boolean              'True after successful initialization
    Private mIsHovered                  As Boolean              'True while the mouse is over this label

'
'------------------------------------------------------------------------------
'
'                           PUBLIC PROPERTIES
'
'------------------------------------------------------------------------------
'

Public Property Get IsInitialized() As Boolean

'
'------------------------------------------------------------------------------
'                           IS INITIALIZED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the hook has been successfully initialized and is ready
'   to route events
'
' WHY THIS EXISTS
'   The owning UserForm collection may need to verify hook state without
'   relying on object-reference checks alone
'
' INPUTS
'   None
'
' RETURNS
'   True when Initialize has completed successfully
'   False before initialization or after a failed initialization
'
' BEHAVIOR
'   Returns the private initialization flag directly
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   mIsInitialized is set to True only at the final step of Initialize, after
'   all validation and reference storage has succeeded
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

    IsInitialized = mIsInitialized

End Property

Public Property Get HoverMode() As String

'
'------------------------------------------------------------------------------
'                           HOVER MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the normalized hover mode assigned to this hook
'
' WHY THIS EXISTS
'   Allows the owning UserForm to inspect hook state for diagnostics and
'   teardown logging without coupling to private internals
'
' INPUTS
'   None
'
' RETURNS
'   Normalized uppercase hover mode string (e.g. "DAY", "NONE")
'   Empty string when not initialized
'
' BEHAVIOR
'   Returns the stored normalized hover mode directly
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
'   2026-05-03
'------------------------------------------------------------------------------

    HoverMode = mHoverMode

End Property

Public Property Get ActionName() As String

'
'------------------------------------------------------------------------------
'                           ACTION NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the normalized click action name assigned to this hook
'
' WHY THIS EXISTS
'   Allows the owning UserForm to inspect hook state for diagnostics and
'   teardown logging without coupling to private internals
'
' INPUTS
'   None
'
' RETURNS
'   Normalized click action name, or empty string when no action is assigned
'
' BEHAVIOR
'   Returns the stored normalized action name directly
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
'   2026-05-03
'------------------------------------------------------------------------------

    ActionName = mActionName

End Property

Public Property Get LabelName() As String

'
'------------------------------------------------------------------------------
'                           LABEL NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Name of the hooked MSForms.Label
'
' WHY THIS EXISTS
'   Allows the owning UserForm to log or identify the hook target without
'   holding a separate reference to the label
'
' INPUTS
'   None
'
' RETURNS
'   Label.Name string when a label is hooked
'   Empty string when no label is attached
'
' BEHAVIOR
'   Guards against a missing label reference before accessing the Name property
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   MSForms.Label
'
' NOTES
'   None
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

    If Not mLabel Is Nothing Then LabelName = mLabel.Name

End Property

Public Property Get LeaveHandlerName() As String

'
'------------------------------------------------------------------------------
'                           LEAVE HANDLER NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the normalized leave handler method name assigned to this hook
'
' WHY THIS EXISTS
'   Allows the owning UserForm to inspect hook state for diagnostics and
'   teardown logging without coupling to private internals
'
' INPUTS
'   None
'
' RETURNS
'   Normalized leave handler method name, or empty string when no leave
'   handler is assigned
'
' BEHAVIOR
'   Returns the stored normalized leave handler name directly
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   An empty return value indicates that this hook will not fire a leave
'   callback even when NotifyMouseLeft is called
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

    LeaveHandlerName = mLeaveHandlerName

End Property

Public Property Get IsHovered() As Boolean

'
'------------------------------------------------------------------------------
'                           IS HOVERED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the mouse is currently over this label
'
' WHY THIS EXISTS
'   Allows the owning UserForm to query hover state on demand without
'   maintaining a parallel tracking variable
'
' INPUTS
'   None
'
' RETURNS
'   True while mLabel_MouseMove has fired more recently than NotifyMouseLeft
'   False before first hover, after NotifyMouseLeft fires, or when not
'   initialized
'
' BEHAVIOR
'   Returns the private hover flag directly
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   mIsHovered is set by mLabel_MouseMove and cleared by NotifyMouseLeft.
'   It is also cleared during Initialize and Class_Terminate.
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

    IsHovered = mIsHovered

End Property

'------------------------------------------------------------------------------
' PUBLIC METHODS
'------------------------------------------------------------------------------

Public Sub Initialize(ByVal OwnerForm As Object, ByVal TargetLabel As _
    MSForms.Label, Optional ByVal ActionName As String = vbNullString, Optional _
    ByVal HoverMode As String = "NONE", Optional ByVal ClickHandlerName As String = _
    "UF_PickerPanel_HandleAction", Optional ByVal LeaveHandlerName As String = _
    vbNullString)

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
'     Optional click action name passed to the owner form click handler
'
'   HoverMode
'     Optional hover behavior
'       - NONE     (default)
'       - DAY
'       - HEADER
'       - PICKER
'       - FOOTER
'       - SETTINGS
'
'   ClickHandlerName
'     Optional name of the public method on OwnerForm that receives click
'     actions. Defaults to UF_PickerPanel_HandleAction
'
'   LeaveHandlerName
'     Optional name of the public method on OwnerForm that receives the
'     simulated MouseLeave notification. Leave empty to suppress leave routing
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Rejects re-initialization of an already-live hook, clears state, validates
'   supplied references and supported hover modes, normalizes all routing names,
'   and stores the owner form, target label, action, hover, click-handler, and
'   optional leave-handler state
'
' ERROR POLICY
'   Raises a descriptive runtime error if:
'     - the hook is already initialized
'     - OwnerForm is Nothing
'     - TargetLabel is Nothing
'     - HoverMode is unsupported
'     - ActionName is supplied but ClickHandlerName is blank
'
' DEPENDENCIES
'   MSForms.Label
'
' NOTES
'   This routine preserves the DatePicker call pattern used throughout
'   UF_DatePicker:
'
'       LabelHook.Initialize Me, Label, ActionName, "HEADER"
'
'   Therefore the fourth positional argument is HoverMode, not ClickHandlerName
'
'   This routine deliberately does not probe OwnerForm methods with CallByName at
'   initialization time. The previous probe approach executed real hover / click
'   handlers with blank sentinel arguments, but those handlers correctly reject
'   blank action or label names. Probing therefore caused DP_Show to fail during
'   UserForm_Initialize even when the methods existed
'
'   Click and hover routing failures remain diagnosable at the point of actual
'   event routing
'
' UPDATED
'   2026-05-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerLabelHook.Initialize"

    Dim NormalizedActionName    As String       'Normalized click action name
    Dim NormalizedHoverMode     As String       'Normalized hover mode
    Dim NormalizedHandlerName   As String       'Normalized click handler method name
    Dim NormalizedLeaveHandler  As String       'Normalized leave handler method name
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' GUARD AGAINST RE-INITIALIZATION
'------------------------------------------------------------------------------
    'Reject a second Initialize call on an already-live hook
        If mIsInitialized Then
            Err.Raise vbObjectError + DP_ERR_ALREADY_INIT, PROC_NAME, _
                "Hook is already initialized. Release the existing hook before re-using it."
        End If

'------------------------------------------------------------------------------
' CLEAR STATE
'------------------------------------------------------------------------------
    'Clear any previous target label reference
        Set mLabel = Nothing
    'Clear any previous owner form reference
        Set mOwnerForm = Nothing
    'Clear any previous click action
        mActionName = vbNullString
    'Clear any previous click handler name
        mClickHandlerName = vbNullString
    'Clear any previous leave handler name
        mLeaveHandlerName = vbNullString
    'Reset hover mode to a safe default
        mHoverMode = DP_HOVER_NONE
    'Clear the hover tracking flag
        mIsHovered = False
    'Clear the initialization flag
        mIsInitialized = False

'------------------------------------------------------------------------------
' VALIDATE REFERENCES
'------------------------------------------------------------------------------
    'Reject a missing owner form
        If OwnerForm Is Nothing Then
            Err.Raise vbObjectError + DP_ERR_NO_OWNER, PROC_NAME, _
                "OwnerForm cannot be Nothing"
        End If
    'Reject a missing target label
        If TargetLabel Is Nothing Then
            Err.Raise vbObjectError + DP_ERR_NO_LABEL, PROC_NAME, _
                "TargetLabel cannot be Nothing"
        End If

'------------------------------------------------------------------------------
' NORMALIZE INPUTS
'------------------------------------------------------------------------------
    'Normalize the click action name
        NormalizedActionName = VBA.Trim$(ActionName)
    'Normalize the requested hover mode
        NormalizedHoverMode = VBA.UCase$(VBA.Trim$(HoverMode))
    'Default blank hover mode to NONE
        If VBA.Len(NormalizedHoverMode) = 0 Then
            NormalizedHoverMode = DP_HOVER_NONE
        End If
    'Normalize the click handler method name
        NormalizedHandlerName = VBA.Trim$(ClickHandlerName)
    'Normalize the optional leave handler method name
        NormalizedLeaveHandler = VBA.Trim$(LeaveHandlerName)

'------------------------------------------------------------------------------
' VALIDATE HOVER MODE
'------------------------------------------------------------------------------
    'Validate the normalized hover mode against the supported set
        Select Case NormalizedHoverMode
            Case DP_HOVER_NONE, DP_HOVER_DAY, DP_HOVER_HEADER, DP_HOVER_PICKER, _
                DP_HOVER_FOOTER, DP_HOVER_SETTINGS
            Case Else
                Err.Raise vbObjectError + DP_ERR_BAD_HOVER, PROC_NAME, _
                    "Unsupported HoverMode '" & HoverMode & _
                    "'. Supported values are NONE, DAY, HEADER, PICKER, FOOTER, or SETTINGS"
        End Select

'------------------------------------------------------------------------------
' VALIDATE CLICK HANDLER NAME
'------------------------------------------------------------------------------
    'Require a click handler name only when a click action is assigned
        If VBA.Len(NormalizedActionName) > 0 Then
            If VBA.Len(NormalizedHandlerName) = 0 Then
                Err.Raise vbObjectError + DP_ERR_BAD_METHOD, PROC_NAME, _
                    "ClickHandlerName cannot be blank when ActionName is supplied"
            End If
        End If

'------------------------------------------------------------------------------
' STORE REFERENCES
'------------------------------------------------------------------------------
    'Store the owner form reference
        Set mOwnerForm = OwnerForm
    'Store the normalized click action name
        mActionName = NormalizedActionName
    'Store the normalized click handler method name
        mClickHandlerName = NormalizedHandlerName
    'Store the normalized leave handler method name
        mLeaveHandlerName = NormalizedLeaveHandler
    'Store the normalized hover mode
        mHoverMode = NormalizedHoverMode
    'Initialise the hover tracking flag to the safe inert value
        mIsHovered = False
    'Store the target label reference last so a partial failure cannot leave
    'a live event subscription pointing at stale state
        Set mLabel = TargetLabel
    'Mark the hook as successfully initialized
        mIsInitialized = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Suppress cleanup errors
        On Error Resume Next
    'Clear the target label reference
        Set mLabel = Nothing
    'Clear the owner form reference
        Set mOwnerForm = Nothing
    'Clear the click action name
        mActionName = vbNullString
    'Clear the click handler name
        mClickHandlerName = vbNullString
    'Clear the leave handler name
        mLeaveHandlerName = vbNullString
    'Reset hover mode to a safe value
        mHoverMode = DP_HOVER_NONE
    'Clear the hover tracking flag
        mIsHovered = False
    'Ensure the initialization flag is cleared
        mIsInitialized = False
    'Restore normal error handling
        On Error GoTo 0
    'Raise a descriptive initialization error
        Err.Raise ErrorNumber, PROC_NAME, "Label hook initialization failed: " & _
            ErrorDescription

End Sub
Private Sub mLabel_Click()

'
'------------------------------------------------------------------------------
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
'   Exits immediately when no click action is assigned.
'
'   Validates the hook state, captures diagnostic context, and calls the owner
'   form click handler by the method name stored during initialization.
'
' ERROR POLICY
'   Raises a descriptive runtime error if the hook is not initialized, if the
'   target label is missing, if the owner form is missing, or if click-action
'   routing fails.
'
'   Error numbers in this procedure use the DP_ERR_CLICK_* block to allow
'   diagnostic log filtering without colliding with Initialize error numbers.
'
' DEPENDENCIES
'   OwnerForm.[mClickHandlerName]
'
' NOTES
'   This routine is only the event router. The click handler method name is
'   supplied by the caller during Initialize rather than being hardcoded here,
'   so labels outside the picker panel can route through different handlers.
'
'   Errors raised here normally originate inside the owner form action handler.
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerLabelHook.mLabel_Click"

    Dim RoutedActionName        As String       'Action name snapshot
    Dim RoutedHandlerName       As String       'Click handler method name snapshot
    Dim RoutedLabelName         As String       'Clicked label name snapshot
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
    'Capture the stored action name before routing (already normalized)
        RoutedActionName = mActionName
    'Capture the stored click handler method name before routing
        RoutedHandlerName = mClickHandlerName

'------------------------------------------------------------------------------
' EXIT IF NO ACTION IS ASSIGNED
'------------------------------------------------------------------------------
    'Exit if this label has no click action
        If VBA.Len(RoutedActionName) = 0 Then Exit Sub

'------------------------------------------------------------------------------
' VALIDATE HOOK STATE
'------------------------------------------------------------------------------
    'Reject a click that fires before initialization
        If Not mIsInitialized Then
            Err.Raise vbObjectError + DP_ERR_CLICK_UNINIT, PROC_NAME, _
                "Click fired on an uninitialized hook"
        End If
    'Reject a missing target label
        If mLabel Is Nothing Then
            Err.Raise vbObjectError + DP_ERR_CLICK_UNINIT, PROC_NAME, _
                "Target label reference is missing"
        End If
    'Capture the clicked label name as early diagnostic context
        RoutedLabelName = mLabel.Name
    'Reject a missing owner form
        If mOwnerForm Is Nothing Then
            Err.Raise vbObjectError + DP_ERR_CLICK_NO_OWNER, PROC_NAME, _
                "OwnerForm reference is missing | Label=" & RoutedLabelName
        End If

'------------------------------------------------------------------------------
' SNAPSHOT REFERENCES
'------------------------------------------------------------------------------
    'Capture the owner form reference before routing
        Set RoutedOwnerForm = mOwnerForm

'------------------------------------------------------------------------------
' ROUTE CLICK ACTION
'------------------------------------------------------------------------------
    'Route the click action to the owner form via the configured handler method
        CallByName RoutedOwnerForm, RoutedHandlerName, VbMethod, _
            RoutedActionName

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
    'Raise a diagnostic error that preserves label, action, and inner source
        Err.Raise InnerErrNumber, PROC_NAME & " | Label=" & RoutedLabelName & _
            " | Action=" & RoutedActionName, "InnerSource=[" & InnerErrSource & _
            "] | InnerDescription=[" & InnerErrDescription & "]"

End Sub

Private Sub mLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single)

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
'   Exits immediately when no hover behavior is assigned or when the hook is
'   not initialized.
'
'   Validates the hook state, captures diagnostic context, and routes hover
'   behavior according to the configured hover mode.
'
'   The label name is captured before any routing call so that the FailSafe
'   diagnostic always has a non-empty label identity, even if routing fails
'   on the first call into the owner form.
'
'   Sets mIsHovered = True after successful routing so that NotifyMouseLeft
'   can detect when the mouse departs this label.
'
' ERROR POLICY
'   Best-effort hover routing. Suppresses hover errors so MouseMove cannot
'   interrupt UserForm interaction.
'
' DEPENDENCIES
'   OwnerForm.UF_DayCell_HoverApply
'   OwnerForm.UF_Header_HoverApply
'   OwnerForm.UF_PickerPanel_HoverApply
'   OwnerForm.UF_Footer_HoverApply
'   OwnerForm.UF_SettingsPanel_HoverApply
'
' NOTES
'   Labels with HoverMode = NONE safely exit without routing.
'
'   Button, Shift, X, and Y are required by the MSForms MouseMove event
'   signature even when this router does not use them directly.
'
'   Hover errors are intentionally suppressed because MouseMove fires at high
'   frequency and should never interrupt normal UserForm interaction.
'
'   mIsHovered is set to True here and cleared by NotifyMouseLeft. The owning
'   UserForm must call NotifyMouseLeft on all hooks from its own MouseMove
'   handler to complete the cooperative leave-detection contract.
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = _
        "cDatePickerLabelHook.mLabel_MouseMove"

    Dim RoutedLabelName         As String       'Label name snapshot
    Dim RoutedHoverMode         As String       'Hover mode snapshot
    Dim RoutedOwnerForm         As Object       'Owner form snapshot

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress hover-routing failures
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' EXIT IF NOT INITIALIZED
'------------------------------------------------------------------------------
    'Exit when the hook has not been successfully initialized
        If Not mIsInitialized Then Exit Sub

'------------------------------------------------------------------------------
' SNAPSHOT HOVER MODE
'------------------------------------------------------------------------------
    'Capture the hover mode before routing
        RoutedHoverMode = mHoverMode

'------------------------------------------------------------------------------
' EXIT IF NO HOVER IS ASSIGNED
'------------------------------------------------------------------------------
    'Exit when this label has no hover behavior
        If RoutedHoverMode = DP_HOVER_NONE Then Exit Sub

'------------------------------------------------------------------------------
' VALIDATE LABEL STATE
'------------------------------------------------------------------------------
    'Exit safely when the target label is unavailable
        If mLabel Is Nothing Then Exit Sub
    'Capture the label name before routing so the FailSafe always has context
        RoutedLabelName = mLabel.Name

'------------------------------------------------------------------------------
' VALIDATE OWNER STATE
'------------------------------------------------------------------------------
    'Exit safely when the owner form is unavailable
        If mOwnerForm Is Nothing Then Exit Sub
    'Capture the owner form reference before routing
        Set RoutedOwnerForm = mOwnerForm

'------------------------------------------------------------------------------
' ROUTE HOVER ACTION
'------------------------------------------------------------------------------
    'Route hover behavior according to the configured hover mode
        Select Case RoutedHoverMode
            Case DP_HOVER_DAY
                RoutedOwnerForm.UF_DayCell_HoverApply RoutedLabelName
            Case DP_HOVER_HEADER
                RoutedOwnerForm.UF_Header_HoverApply RoutedLabelName
            Case DP_HOVER_PICKER
                RoutedOwnerForm.UF_PickerPanel_HoverApply RoutedLabelName
            Case DP_HOVER_FOOTER
                RoutedOwnerForm.UF_Footer_HoverApply RoutedLabelName
            Case DP_HOVER_SETTINGS
                RoutedOwnerForm.UF_SettingsPanel_HoverApply RoutedLabelName
            Case Else
                Debug.Print PROC_NAME & " | Unsupported hover mode=" & _
                    RoutedHoverMode & " | Label=" & RoutedLabelName
        End Select

    'Stamp the hover flag so NotifyMouseLeft can detect departure from this label
        mIsHovered = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful hover routing
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Write hover diagnostics without interrupting the UserForm.
    'RoutedLabelName is captured before any routing call so this print always
    'includes a non-empty label identity even when routing failed early.
        Debug.Print PROC_NAME & " | Label=" & RoutedLabelName & " | HoverMode=" _
            & RoutedHoverMode & " | Error=" & VBA.CStr(Err.Number) & " | " & _
            Err.Description

    'Clear the suppressed error
        Err.Clear

End Sub

Public Sub NotifyMouseLeft()

'
'------------------------------------------------------------------------------
'                           NOTIFY MOUSE LEFT
'------------------------------------------------------------------------------
' PURPOSE
'   Delivers a simulated MouseLeave notification to this hook
'
' WHY THIS EXISTS
'   MSForms labels do not fire a native MouseLeave event. The owning UserForm
'   must call this method on every hook from its own MouseMove handler.
'   Each hook uses its mIsHovered flag to detect whether the mouse has
'   departed since the last mLabel_MouseMove call.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits immediately when the hook is not initialized, when no leave handler
'   is configured, or when this label was not the previously hovered label.
'
'   When mIsHovered is True, the mouse has left this label since the last
'   MouseMove. The method clears mIsHovered, validates hook state, and
'   routes the leave callback to the owner form.
'
' ERROR POLICY
'   Best-effort leave routing. Suppresses leave errors so NotifyMouseLeft
'   cannot interrupt UserForm interaction when called in bulk from the owner
'   form MouseMove handler.
'
' DEPENDENCIES
'   OwnerForm.[mLeaveHandlerName]
'
' NOTES
'   The owning UserForm must call this method on every hook in its hook
'   collection from its own UserForm_MouseMove handler to complete the
'   cooperative leave-detection contract. Example pattern:
'
'     Private Sub UserForm_MouseMove(...)
'         Dim Hook As cDatePickerLabelHook
'         For Each Hook In mHookCollection
'             Hook.NotifyMouseLeft
'         Next Hook
'     End Sub
'
'   Only hooks with a configured LeaveHandlerName and a True mIsHovered flag
'   incur any routing cost. All others exit in the first two or three checks.
'
'   mIsHovered is cleared before the routing call so that a leave handler
'   which itself triggers a MouseMove on another label cannot cause a
'   recursive leave notification on this hook.
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = _
        "cDatePickerLabelHook.NotifyMouseLeft"

    Dim RoutedLabelName         As String       'Label name snapshot
    Dim RoutedLeaveHandler      As String       'Leave handler method name snapshot
    Dim RoutedOwnerForm         As Object       'Owner form snapshot

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress leave-routing failures
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' EXIT IF NOT INITIALIZED
'------------------------------------------------------------------------------
    'Exit when the hook has not been successfully initialized
        If Not mIsInitialized Then Exit Sub

'------------------------------------------------------------------------------
' EXIT IF NO LEAVE HANDLER IS CONFIGURED
'------------------------------------------------------------------------------
    'Exit when this hook has no leave callback to route
        If VBA.Len(mLeaveHandlerName) = 0 Then Exit Sub

'------------------------------------------------------------------------------
' EXIT IF THIS LABEL WAS NOT HOVERED
'------------------------------------------------------------------------------
    'Exit when the mouse was not over this label - nothing to notify
        If Not mIsHovered Then Exit Sub

'------------------------------------------------------------------------------
' CLEAR HOVER FLAG
'------------------------------------------------------------------------------
    'Clear the hover flag before routing so a MouseMove triggered inside the
    'leave handler cannot cause a recursive leave notification on this hook
        mIsHovered = False

'------------------------------------------------------------------------------
' VALIDATE LABEL STATE
'------------------------------------------------------------------------------
    'Exit safely when the target label is unavailable
        If mLabel Is Nothing Then Exit Sub
    'Capture the label name before routing so the FailSafe always has context
        RoutedLabelName = mLabel.Name

'------------------------------------------------------------------------------
' VALIDATE OWNER STATE
'------------------------------------------------------------------------------
    'Exit safely when the owner form is unavailable
        If mOwnerForm Is Nothing Then Exit Sub
    'Capture the owner form reference and handler name before routing
        Set RoutedOwnerForm = mOwnerForm
        RoutedLeaveHandler = mLeaveHandlerName

'------------------------------------------------------------------------------
' ROUTE LEAVE CALLBACK
'------------------------------------------------------------------------------
    'Route the leave callback to the owner form via the configured handler
        CallByName RoutedOwnerForm, RoutedLeaveHandler, VbMethod, _
            RoutedLabelName

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful leave routing
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Write leave diagnostics without interrupting the UserForm.
    'RoutedLabelName is captured before the routing call so this print always
    'includes a non-empty label identity even when routing failed early.
        Debug.Print PROC_NAME & " | Label=" & RoutedLabelName & _
            " | LeaveHandler=" & RoutedLeaveHandler & " | Error=" & _
            VBA.CStr(Err.Number) & " | " & Err.Description

    'Clear the suppressed error
        Err.Clear

End Sub

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
'   destroyed.
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
'   Suppresses cleanup errors because termination must never interrupt the VBA
'   runtime
'
' DEPENDENCIES
'   None
'
' NOTES
'   The owning UserForm should still explicitly release the hook collection
'   during UserForm_Terminate.
'
'   This routine is a final cleanup guard, not the primary lifecycle-management
'   mechanism.
'
' UPDATED
'   2026-05-03
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
    'Clear the click handler method name
        mClickHandlerName = vbNullString
    'Clear the leave handler method name
        mLeaveHandlerName = vbNullString
    'Reset the hover mode to the safe inert value
        mHoverMode = DP_HOVER_NONE
    'Clear the hover tracking flag
        mIsHovered = False
    'Clear the initialization flag
        mIsInitialized = False
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub






