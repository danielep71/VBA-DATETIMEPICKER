Attribute VB_Name = "M_DP_ContextMenu"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_DP_ContextMenu
'------------------------------------------------------------------------------
' PURPOSE
'   Adds and removes the Date / Time Picker entry in the worksheet
'   right-click menus
'
' WHY THIS EXISTS
'   CommandBars are shared by every workbook and add-in in the Excel
'   process. One module owns the tagged entry, so adding, removing and
'   proving removal use the same identity (#57)
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - M_ContextMenu_Update and M_ContextMenu_EnsureForRightClick
'     - M_ContextMenu_Remove, best-effort removal
'     - M_ContextMenu_TryRemove, removal verified on every Cell and
'       List Range Popup bar, used by lifecycle cleanup
'
' ERROR POLICY
'   Each procedure documents its own policy
'
' DEPENDENCIES
'   Office CommandBars object model
'   M_DP_MacroNames (M_GetQualifiedMacroName)
'   M_DP_Core (gDP_ShowRightClick, M_Settings_EnsureLoaded,
'     M_Lease_IsOwner, M_Lease_TryReclaimVanished)
'
' NOTES
'   The entry runs the facade callback DP_Click. Its tag keeps the legacy
'   VBA_DATETIMEPICKER identifier for backward compatibility
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
    Private Const DP_CONTEXT_MENU_TAG     As String = "VBA_DATETIMEPICKER"  'Legacy context-menu tag
    Private Const DP_CONTEXT_MENU_CAPTION As String = "Date / Time Picker"  'Context-menu caption; matches DP_MSGBOX_TITLE (#88)
    Private Const DP_CONTEXT_MENU_FACEID  As Long = 1992                    'Context-menu icon FaceId
    Private Const DP_CONTEXT_MENU_BEFORE  As Long = 1                       'Context-menu insertion position


'
'------------------------------------------------------------------------------
'
'                                 CONTEXT MENU
'
'------------------------------------------------------------------------------
'

Public Sub M_ContextMenu_Update()

'
'------------------------------------------------------------------------------
'                           RIGHT-CLICK MENU UPDATE
'------------------------------------------------------------------------------
' PURPOSE
'   Synchronizes DatePicker right-click menu integration with the current setting
'
' WHY THIS EXISTS
'   Context-menu entries should be added, detected, and removed consistently
'   using the DatePicker stable tag. Centralizing the update policy keeps menu
'   state aligned with persisted settings and avoids duplicate or stale command
'   bar controls
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures settings are loaded, then adds DatePicker right-click entries when
'   gDP_ShowRightClick is True and removes them when gDP_ShowRightClick is False
'
' ERROR POLICY
'   Raises unexpected failures back to the caller with this procedure name as
'   the error source
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_ContextMenu_Add
'   M_ContextMenu_Remove
'   gDP_ShowRightClick
'   Application.CommandBars
'
' NOTES
'   This routine owns the high-level policy decision only
'
'   The detailed add / remove mechanics remain delegated to M_ContextMenu_Add
'   and M_ContextMenu_Remove
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_Update"

    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure settings are loaded before reading the right-click feature flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' APPLY CONTEXT-MENU POLICY
'------------------------------------------------------------------------------
    'Add right-click entries when the feature is enabled
        If gDP_ShowRightClick Then
            M_ContextMenu_Add
    'Otherwise remove DatePicker right-click entries
        Else
            M_ContextMenu_Remove
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
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Sub

Public Sub M_ContextMenu_EnsureForRightClick()

'
'------------------------------------------------------------------------------
'                           RIGHT-CLICK MENU RE-ENSURE
'------------------------------------------------------------------------------
' PURPOSE
'   Restores a missing DatePicker right-click entry just before Excel shows a
'   cell context menu
'
' WHY THIS EXISTS
'   The entry is a temporary command-bar control added once by DP_Start. When
'   the add-in loads at Excel startup, the entry can be gone by the time the
'   user first right-clicks, although the runtime started (#113). On the
'   affected host the provider lease bar, also a temporary command bar, was
'   gone as well, so the copy still held its owner token but could no longer
'   prove ownership. Re-ensuring on demand repairs both, whenever and by
'   whatever they were removed
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   When this project holds its owner token but the lease has vanished, reclaims
'   the free lease through the silent admission path. A lease held by anyone
'   else, or unreadable, is never taken, and a project without a token never
'   claims one here
'   Does nothing unless this project then owns the runtime lease and the
'   right-click setting is enabled
'   Otherwise delegates to M_ContextMenu_Add, which adds an entry only to bars
'   that do not already carry one
'
' ERROR POLICY
'   Best-effort. Never raises: it runs from an Excel event, before the menu is
'   shown, and a failure must not affect the right-click
'
' DEPENDENCIES
'   M_Lease_IsOwner
'   M_Lease_TryReclaimVanished
'   M_Settings_EnsureLoaded
'   M_ContextMenu_Add
'   gDP_ShowRightClick
'
' NOTES
'   Only adds. Removing entries when the setting is disabled stays with
'   M_ContextMenu_Update, so this path can never change the user's setting
'
'   A copy that does not own the runtime never touches the shared menus, which
'   keeps the one-provider rule intact. Reclaiming a vanished lease is the same
'   acquisition any copy may make of a free lease, restricted here to the copy
'   that last held it
'
' UPDATED
'   2026-10-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_EnsureForRightClick"

    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CHECK OWNERSHIP AND SETTING
'------------------------------------------------------------------------------
    'Only the copy that owns the runtime may touch the shared right-click menus
        If Not M_Lease_IsOwner() Then
        'Only a copy that held the lease may reclaim it, and only while it is free
            If Not M_Lease_TryReclaimVanished(PROC_NAME) Then Exit Sub
        End If
    'Ensure settings are loaded before reading the right-click feature flag
        M_Settings_EnsureLoaded
    'Leave a disabled entry alone; removal belongs to M_ContextMenu_Update
        If Not gDP_ShowRightClick Then Exit Sub

'------------------------------------------------------------------------------
' RESTORE MISSING ENTRIES
'------------------------------------------------------------------------------
    'Add the entry to every supported bar that lost it; existing entries stay
        M_ContextMenu_Add

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
    'Write diagnostics without interrupting the right-click
        Debug.Print PROC_NAME & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed error
        Err.Clear

End Sub

Private Sub M_ContextMenu_Add()

'
'==============================================================================
'                           ADD RIGHT-CLICK MENU
'------------------------------------------------------------------------------
' PURPOSE
'   Adds DatePicker entries to the supported Excel right-click menus
'
' WHY THIS EXISTS
'   DatePicker right-click integration targets more than one Excel context menu
'
'   Centralizing the add operation keeps the supported command-bar list
'   consistent and avoids duplicating menu wiring logic across the project
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds the DatePicker right-click entry to every standard cell context menu
'   Adds the DatePicker right-click entry to every table / list range context
'   menu
'   Delegates duplicate-control prevention to M_ContextMenu_AddToCommandBar
'   Delegates command-bar resolution to M_ContextMenu_GetCommandBars
'
' ERROR POLICY
'   Raises a descriptive runtime error if right-click menu synchronization fails
'   Preserves the original error number and description
'
' DEPENDENCIES
'   M_ContextMenu_GetCommandBars
'   M_ContextMenu_AddToCommandBar
'   Application.CommandBars
'
' NOTES
'   Duplicate-control prevention should remain inside
'   M_ContextMenu_AddToCommandBar because that routine owns command-bar mutation
'   logic and stable-tag checks
'
'   This routine owns only the list of supported right-click command bars
'
'   Excel has more than one command bar named "Cell": one serves Normal view and
'   another Page Layout view. Looking a bar up by name returns only the first, so
'   the entry used to be missing in Page Layout view (#114)
'
'   The first bar of each name keeps the raising contract it always had. Later
'   bars are best-effort: a failure there is logged rather than raised, so the
'   extra coverage can never make a start fail that used to succeed
'
' UPDATED
'   2026-10-01
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME                    As String = "M_ContextMenu_Add"  'Current procedure name
    Const CELL_COMMAND_BAR_NAME        As String = "Cell"               'Standard cell context menu
    Const LIST_RANGE_COMMAND_BAR_NAME  As String = "List Range Popup"   'Table / list range context menu

    Dim HandlerStep                    As String                        'Current handler step for diagnostics
    Dim ErrorNumber                    As Long                          'Captured error number
    Dim ErrorDescription               As String                        'Captured error description
    Dim TargetBar                      As Variant                       'Current command bar
    Dim BarOrdinal                     As Long                          'Position of the bar among same-named bars

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' ADD STANDARD CELL CONTEXT MENU ENTRIES
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Add standard cell context menu entry"
    'Add to every standard cell context menu
        BarOrdinal = 0
        For Each TargetBar In M_ContextMenu_GetCommandBars(CELL_COMMAND_BAR_NAME)
            BarOrdinal = BarOrdinal + 1
            If BarOrdinal = 1 Then
                M_ContextMenu_AddToCommandBar TargetBar
            Else
                M_ContextMenu_TryAddToCommandBar TargetBar
            End If
        Next TargetBar

'------------------------------------------------------------------------------
' ADD TABLE / LIST RANGE CONTEXT MENU ENTRIES
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Add table / list range context menu entry"
    'Add to every table / list range context menu
        BarOrdinal = 0
        For Each TargetBar In M_ContextMenu_GetCommandBars(LIST_RANGE_COMMAND_BAR_NAME)
            BarOrdinal = BarOrdinal + 1
            If BarOrdinal = 1 Then
                M_ContextMenu_AddToCommandBar TargetBar
            Else
                M_ContextMenu_TryAddToCommandBar TargetBar
            End If
        Next TargetBar

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Release the command-bar reference
        Set TargetBar = Nothing
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
    'Release the command-bar reference
        Set TargetBar = Nothing
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Right-click menu add failed: " & ErrorDescription

End Sub

Public Sub M_ContextMenu_Remove()

'
'------------------------------------------------------------------------------
'                           REMOVE RIGHT-CLICK MENU
'------------------------------------------------------------------------------
' PURPOSE
'   Removes DatePicker entries from the supported Excel right-click menus
'
' WHY THIS EXISTS
'   DatePicker right-click integration targets more than one Excel context menu.
'   Centralizing the remove operation keeps the supported command-bar list
'   consistent and avoids leaving stale DatePicker controls after settings
'   changes, workbook close, add-in unload, or manual reset
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Removes the DatePicker right-click entry from:
'     - every standard cell context menu
'     - every table / list range context menu
'
' ERROR POLICY
'   Delegates best-effort removal to M_ContextMenu_RemoveFromCommandBar
'   Does not normally raise for missing, protected, or stale command-bar controls
'
' DEPENDENCIES
'   M_ContextMenu_GetCommandBars
'   M_ContextMenu_RemoveFromCommandBar
'   Application.CommandBars
'
' NOTES
'   Stable-tag based deletion should remain inside
'   M_ContextMenu_RemoveFromCommandBar because that routine owns the command-bar
'   mutation logic and DatePicker control identification policy
'
'   This routine owns only the list of supported right-click command bars
'
'   Every bar sharing a supported name is cleaned, including the Page Layout
'   "Cell" bar, so an entry added there cannot outlive teardown (#114)
'
' UPDATED
'   2026-10-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const CELL_COMMAND_BAR_NAME         As String = "Cell"
    Const LIST_RANGE_COMMAND_BAR_NAME   As String = "List Range Popup"

    Dim TargetBar                       As Variant      'Current command bar

'------------------------------------------------------------------------------
' REMOVE STANDARD CELL CONTEXT MENU ENTRIES
'------------------------------------------------------------------------------
    'Remove from every standard cell context menu
        For Each TargetBar In M_ContextMenu_GetCommandBars(CELL_COMMAND_BAR_NAME)
            M_ContextMenu_RemoveFromCommandBar TargetBar
        Next TargetBar

'------------------------------------------------------------------------------
' REMOVE TABLE / LIST RANGE CONTEXT MENU ENTRIES
'------------------------------------------------------------------------------
    'Remove from every table / list range context menu
        For Each TargetBar In M_ContextMenu_GetCommandBars(LIST_RANGE_COMMAND_BAR_NAME)
            M_ContextMenu_RemoveFromCommandBar TargetBar
        Next TargetBar

'------------------------------------------------------------------------------
' RELEASE REFERENCES
'------------------------------------------------------------------------------
    'Release the command-bar reference
        Set TargetBar = Nothing

End Sub

Private Function M_ContextMenu_GetCommandBar(ByVal CommandBarName As String) As CommandBar

'
'------------------------------------------------------------------------------
'                           GET COMMAND BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Safely resolves an Excel command bar by name
'
' WHY THIS EXISTS
'   DatePicker right-click integration depends on Excel context menus. Command
'   bars may be unavailable, renamed, hidden, unsupported in a given host state,
'   or inaccessible in some Excel configurations, so lookup must fail safely
'
' INPUTS
'   CommandBarName
'     Name of the Excel command bar to retrieve
'
' RETURNS
'   Matching CommandBar object when available
'   Nothing when the command bar cannot be resolved
'
' BEHAVIOR
'   Attempts to retrieve Application.CommandBars(CommandBarName) and returns a
'   safe Nothing reference when lookup fails
'
' ERROR POLICY
'   Safe default. Command-bar lookup errors are suppressed and reported to the
'   Immediate Window for diagnostics
'
' DEPENDENCIES
'   Application.CommandBars
'   CommandBar
'
' NOTES
'   This routine does not create command bars
'
'   This routine deliberately does not raise outward because missing context
'   menus should not break DatePicker startup, settings synchronization, or
'   teardown
'
'   Add / remove routines must tolerate a Nothing command-bar reference
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_GetCommandBar"

    Dim LookupErrNumber         As Long          'Captured command-bar lookup error number
    Dim LookupErrDescription    As String        'Captured command-bar lookup error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default
        Set M_ContextMenu_GetCommandBar = Nothing
    'Suppress command-bar lookup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' RETRIEVE COMMAND BAR
'------------------------------------------------------------------------------
    'Retrieve the requested command bar
        Set M_ContextMenu_GetCommandBar = Application.CommandBars(CommandBarName)
    'Capture command-bar lookup error number
        LookupErrNumber = Err.Number
    'Capture command-bar lookup error description
        LookupErrDescription = Err.Description
    'Clear any suppressed lookup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' DIAGNOSTICS
'------------------------------------------------------------------------------
    'Write diagnostics only when command-bar lookup failed with an error
        If LookupErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | CommandBarName=" & CommandBarName & _
                " | Error=" & VBA.CStr(LookupErrNumber) & _
                " | " & LookupErrDescription
        End If

End Function

Private Function M_ContextMenu_GetCommandBars(ByVal CommandBarName As String) As Collection

'
'------------------------------------------------------------------------------
'                         GET ALL SAME-NAMED COMMAND BARS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns every Excel command bar that carries the requested name
'
' WHY THIS EXISTS
'   Excel exposes more than one command bar named "Cell": one serves Normal view
'   and another Page Layout view. Application.CommandBars(Name) returns only the
'   first, so a right-click entry registered by name was missing in Page Layout
'   view and a teardown verified by name could not see that bar at all (#114)
'
' INPUTS
'   CommandBarName
'     Name of the Excel command bars to collect
'
' RETURNS
'   A Collection of CommandBar objects, possibly empty. Never Nothing
'
' BEHAVIOR
'   Walks Application.CommandBars and collects every bar whose Name matches,
'   case-insensitively as Excel's own name lookup does. When the walk finds none,
'   falls back to M_ContextMenu_GetCommandBar so a host that cannot be enumerated
'   keeps the previous single-bar behavior
'
' ERROR POLICY
'   Safe default. Enumeration errors are suppressed and reported to the
'   Immediate Window. The caller always receives a Collection
'
' DEPENDENCIES
'   M_ContextMenu_GetCommandBar
'   Application.CommandBars
'
' NOTES
'   This routine does not create command bars
'
'   Callers must tolerate an empty Collection, exactly as they tolerated a
'   Nothing command bar before
'
' UPDATED
'   2026-10-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_GetCommandBars"

    Dim Result                  As Collection    'Collected command bars
    Dim Bar                     As CommandBar    'Current command bar while walking
    Dim FallbackBar             As CommandBar    'Single bar resolved by name
    Dim BarName                 As String        'Name read from the current bar
    Dim WalkErrNumber           As Long          'Captured enumeration error number
    Dim WalkErrDescription      As String        'Captured enumeration error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create the result so callers never receive Nothing
        Set Result = New Collection
    'Suppress enumeration errors
        On Error Resume Next

'------------------------------------------------------------------------------
' COLLECT MATCHING BARS
'------------------------------------------------------------------------------
    'Walk every command bar and keep the ones with the requested name. The name
    'is read into a variable first: under Resume Next, a failing If condition
    'would fall into its Then branch and collect an unreadable bar
        For Each Bar In Application.CommandBars
            BarName = VBA.vbNullString
            BarName = Bar.Name
            If VBA.StrComp(BarName, CommandBarName, vbTextCompare) = 0 Then
                Result.Add Bar
            End If
        Next Bar
    'Capture enumeration error number
        WalkErrNumber = Err.Number
    'Capture enumeration error description
        WalkErrDescription = Err.Description
    'Clear any suppressed enumeration error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' DIAGNOSTICS
'------------------------------------------------------------------------------
    'Write diagnostics only when enumeration failed with an error
        If WalkErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | CommandBarName=" & CommandBarName & _
                " | Error=" & VBA.CStr(WalkErrNumber) & _
                " | " & WalkErrDescription
        End If

'------------------------------------------------------------------------------
' FALL BACK TO NAME LOOKUP
'------------------------------------------------------------------------------
    'Keep the previous single-bar behavior when the walk found nothing
        If Result.Count = 0 Then
            Set FallbackBar = M_ContextMenu_GetCommandBar(CommandBarName)
            If Not FallbackBar Is Nothing Then Result.Add FallbackBar
        End If

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
    'Return the collected bars
        Set M_ContextMenu_GetCommandBars = Result
    'Release object references
        Set FallbackBar = Nothing
        Set Bar = Nothing
        Set Result = Nothing

End Function

Private Sub M_ContextMenu_TryAddToCommandBar(ByVal TargetCommandBar As CommandBar)

'
'------------------------------------------------------------------------------
'                      TRY TO ADD TO AN ADDITIONAL COMMAND BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Adds the DatePicker command to an additional same-named command bar without
'   raising
'
' WHY THIS EXISTS
'   Registration on the second "Cell" bar (Page Layout view) is new in #114.
'   A host that refuses it must not turn a start that used to succeed into a
'   failure
'
' INPUTS
'   TargetCommandBar
'     Command bar to add the DatePicker command to
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to M_ContextMenu_AddToCommandBar and reports any failure to the
'   Immediate Window
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_ContextMenu_AddToCommandBar
'
' NOTES
'   Removal already treats every bar as best-effort, and teardown verification
'   reads every bar back, so a control that was added here is never left behind
'   unnoticed
'
' UPDATED
'   2026-10-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_TryAddToCommandBar"

    Dim AddErrNumber            As Long          'Captured add error number
    Dim AddErrDescription       As String        'Captured add error description

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
    'Suppress add errors
        On Error Resume Next
    'Add the DatePicker command
        M_ContextMenu_AddToCommandBar TargetCommandBar
    'Capture add error number
        AddErrNumber = Err.Number
    'Capture add error description
        AddErrDescription = Err.Description
    'Clear any suppressed add error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' DIAGNOSTICS
'------------------------------------------------------------------------------
    'Write diagnostics only when the add failed
        If AddErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Error=" & VBA.CStr(AddErrNumber) & _
                " | " & AddErrDescription
        End If

End Sub

Private Sub M_ContextMenu_AddToCommandBar(ByVal TargetCommandBar As CommandBar)

'
'------------------------------------------------------------------------------
'                           ADD TO COMMAND BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Adds the DatePicker command to one Excel command bar
'
' WHY THIS EXISTS
'   DatePicker right-click integration can target multiple Excel context menus.
'   The actual command-bar mutation must remain centralized so duplicate
'   detection, stable tagging, button placement, and button configuration stay
'   consistent across all supported menus
'
' INPUTS
'   TargetCommandBar
'     CommandBar object that should receive the DatePicker entry
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits when no command bar is supplied, exits when the DatePicker control is
'   already present, otherwise adds a temporary command-bar button and configures
'   its action, icon, caption, tag, and grouping
'
' ERROR POLICY
'   Raises unexpected command-bar mutation failures back to the caller with this
'   procedure name as the error source
'
' DEPENDENCIES
'   M_ContextMenu_ContainsDatePicker
'   M_GetQualifiedMacroName
'   CommandBar.Controls.Add
'   CommandBarButton
'   msoControlButton
'   DP_CONTEXT_MENU_BEFORE
'   DP_CONTEXT_MENU_FACEID
'   DP_CONTEXT_MENU_CAPTION
'   DP_CONTEXT_MENU_TAG
'   DP_Click
'
' NOTES
'   Duplicate prevention is based on M_ContextMenu_ContainsDatePicker
'
'   The control is added as Temporary so Excel does not persist it permanently
'   into the command-bar customization layer
'
'   The stable tag is the authoritative identifier used later for cleanup
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_AddToCommandBar"

    Dim ButtonControl           As CommandBarButton     'Created command-bar button
    Dim ErrorNumber             As Long                 'Captured error number
    Dim ErrorDescription        As String               'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit when no command bar is supplied
        If TargetCommandBar Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' AVOID DUPLICATE CONTROL
'------------------------------------------------------------------------------
    'Exit when the DatePicker control is already present
        If M_ContextMenu_ContainsDatePicker(TargetCommandBar) Then Exit Sub

'------------------------------------------------------------------------------
' ADD BUTTON
'------------------------------------------------------------------------------
    'Add the DatePicker button
        Set ButtonControl = TargetCommandBar.Controls.Add( _
            Type:=msoControlButton, _
            Before:=DP_CONTEXT_MENU_BEFORE, _
            Temporary:=True)

'------------------------------------------------------------------------------
' CONFIGURE BUTTON
'------------------------------------------------------------------------------
    'Configure the DatePicker button
        With ButtonControl
            .OnAction = M_GetQualifiedMacroName("DP_Click")
            .FaceId = DP_CONTEXT_MENU_FACEID
            .Caption = DP_CONTEXT_MENU_CAPTION
            .Tag = DP_CONTEXT_MENU_TAG
            .BeginGroup = True
        End With

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Release the command-bar button reference
        Set ButtonControl = Nothing
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
    'Release the command-bar button reference
        Set ButtonControl = Nothing
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Sub

Private Sub M_ContextMenu_RemoveFromCommandBar(ByVal TargetCommandBar As CommandBar)

'
'------------------------------------------------------------------------------
'                           REMOVE FROM COMMAND BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Removes DatePicker controls from one Excel command bar
'
' WHY THIS EXISTS
'   DatePicker right-click menu entries are transient Excel UI controls. They
'   must be removable by stable tag so menu cleanup remains reliable after
'   settings changes, workbook close, add-in unload, or runtime repair
'
' INPUTS
'   TargetCommandBar
'     CommandBar object to clean
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits when no command bar is supplied
'   Reads the command-bar control count defensively
'   Exits safely when command-bar controls cannot be inspected
'   Scans the command-bar controls backwards
'   Deletes every control whose Tag matches DP_CONTEXT_MENU_TAG
'
' ERROR POLICY
'   Best-effort cleanup
'
'   Missing, protected, stale, or otherwise non-readable command-bar controls
'   are skipped and logged to the Immediate Window
'
'   Missing, protected, stale, or otherwise non-removable command-bar controls
'   are skipped and logged to the Immediate Window
'
'   Does not raise outward
'
' DEPENDENCIES
'   CommandBar
'   CommandBarControl
'   DP_CONTEXT_MENU_TAG
'
' NOTES
'   The backward scan is intentional because controls are deleted while the
'   collection is being traversed
'
'   Matching by Tag is preferred over matching by Caption because captions may
'   change with localization or UI wording
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_RemoveFromCommandBar" 'Current procedure name

    Dim Index                   As Long       'CommandBar control index
    Dim ControlCount            As Long       'CommandBar control count
    Dim ControlTag              As String     'Current control tag
    Dim CountErrNumber          As Long       'Captured controls-count error number
    Dim CountErrDescription     As String     'Captured controls-count error description
    Dim DeleteErrNumber         As Long       'Captured delete error number
    Dim DeleteErrDescription    As String     'Captured delete error description
    Dim ReadErrNumber           As Long       'Captured tag-read error number
    Dim ReadErrDescription      As String     'Captured tag-read error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors because this is a best-effort cleanup routine
        On Error Resume Next

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit when no command bar is supplied
        If TargetCommandBar Is Nothing Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' READ CONTROL COUNT
'------------------------------------------------------------------------------
    'Clear any pending error before reading the controls count
        Err.Clear
    'Read the command-bar controls count defensively
        ControlCount = TargetCommandBar.Controls.Count
    'Capture controls-count error number
        CountErrNumber = Err.Number
    'Capture controls-count error description
        CountErrDescription = Err.Description
    'Clear any suppressed controls-count error
        Err.Clear
    'Exit safely when the controls collection cannot be inspected
        If CountErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=ReadControlsCount" & _
                " | Error=" & VBA.CStr(CountErrNumber) & _
                " | " & CountErrDescription
            GoTo ExitProcedure
        End If

'------------------------------------------------------------------------------
' REMOVE DATEPICKER CONTROLS
'------------------------------------------------------------------------------
    'Loop backward through controls so deletion does not disturb pending indexes
        For Index = ControlCount To 1 Step -1
            'Reset the current control tag
                ControlTag = vbNullString
            'Reset tag-read diagnostics
                ReadErrNumber = 0
            'Reset tag-read diagnostic description
                ReadErrDescription = vbNullString
            'Clear any pending error before reading the tag
                Err.Clear
            'Read the current control tag
                ControlTag = VBA.CStr(TargetCommandBar.Controls(Index).Tag)
            'Capture tag-read error number
                ReadErrNumber = Err.Number
            'Capture tag-read error description
                ReadErrDescription = Err.Description
            'Clear any suppressed tag-read error
                Err.Clear
            'Write diagnostics only when tag read failed
                If ReadErrNumber <> 0 Then
                    Debug.Print PROC_NAME & _
                        " | Step=ReadTag" & _
                        " | Index=" & VBA.CStr(Index) & _
                        " | Error=" & VBA.CStr(ReadErrNumber) & _
                        " | " & ReadErrDescription
                End If
            'Delete DatePicker controls by stable tag
                If VBA.StrComp(ControlTag, DP_CONTEXT_MENU_TAG, vbBinaryCompare) = 0 Then
                    'Reset delete diagnostics
                        DeleteErrNumber = 0
                    'Reset delete diagnostic description
                        DeleteErrDescription = vbNullString
                    'Clear any pending error before deleting the control
                        Err.Clear
                    'Delete the matching DatePicker control
                        TargetCommandBar.Controls(Index).Delete
                    'Capture delete error number
                        DeleteErrNumber = Err.Number
                    'Capture delete error description
                        DeleteErrDescription = Err.Description
                    'Clear any suppressed delete error
                        Err.Clear
                    'Write diagnostics only when deletion failed
                        If DeleteErrNumber <> 0 Then
                            Debug.Print PROC_NAME & _
                                " | Step=DeleteControl" & _
                                " | Index=" & VBA.CStr(Index) & _
                                " | Error=" & VBA.CStr(DeleteErrNumber) & _
                                " | " & DeleteErrDescription
                        End If
                End If
        Next Index

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Function M_ContextMenu_ContainsDatePicker(ByVal TargetCommandBar As CommandBar) As Boolean

'
'------------------------------------------------------------------------------
'                           CONTAINS DATEPICKER CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether one Excel command bar already contains a DatePicker control
'
' WHY THIS EXISTS
'   DatePicker right-click menu integration must avoid duplicate controls when
'   menus are refreshed, settings are reapplied, workbooks are activated, or the
'   runtime is repaired
'
' INPUTS
'   TargetCommandBar
'     CommandBar object to inspect
'
' RETURNS
'   True when a control tagged with DP_CONTEXT_MENU_TAG is found
'   False when no matching control is found or the command bar cannot be scanned
'
' BEHAVIOR
'   Exits safely when no command bar is supplied, scans each command-bar control,
'   reads the control Tag defensively, and returns True as soon as the stable
'   DatePicker tag is found
'
' ERROR POLICY
'   Safe-default predicate
'
'   Unexpected command-bar or control-inspection errors return False and are
'   written to the Immediate Window for diagnostics
'
' DEPENDENCIES
'   CommandBar
'   CommandBarControl
'   DP_CONTEXT_MENU_TAG
'
' NOTES
'   Matching by Tag is preferred over matching by Caption because captions may
'   change with localization or UI wording
'
'   This function does not mutate the command bar
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_ContainsDatePicker"

    Dim ControlItem             As CommandBarControl     'CommandBar control being inspected
    Dim ControlTag              As String                'Current control tag
    Dim ReadErrNumber           As Long                  'Captured tag-read error number
    Dim ReadErrDescription      As String                'Captured tag-read error description
    Dim ErrorNumber             As Long                  'Captured runtime error number
    Dim ErrorDescription        As String                'Captured runtime error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default
        M_ContextMenu_ContainsDatePicker = False
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit when no command bar is supplied
        If TargetCommandBar Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' SCAN COMMAND-BAR CONTROLS
'------------------------------------------------------------------------------
    'Loop through command-bar controls
        For Each ControlItem In TargetCommandBar.Controls
            'Reset current control tag
                ControlTag = vbNullString
            'Reset tag-read diagnostics
                ReadErrNumber = 0
            'Reset tag-read diagnostic description
                ReadErrDescription = vbNullString
            'Suppress tag-read errors for unusual command-bar controls
                On Error Resume Next
            'Read the current control tag
                ControlTag = VBA.CStr(ControlItem.Tag)
            'Capture tag-read error number
                ReadErrNumber = Err.Number
            'Capture tag-read error description
                ReadErrDescription = Err.Description
            'Clear any suppressed tag-read error
                Err.Clear
            'Restore safe-default error handling
                On Error GoTo ErrorHandler
            'Write diagnostics only when tag read failed
                If ReadErrNumber <> 0 Then
                    Debug.Print PROC_NAME & _
                        " | Step=ReadTag" & _
                        " | Error=" & VBA.CStr(ReadErrNumber) & _
                        " | " & ReadErrDescription
                End If
            'Return True when the DatePicker stable tag is found
                If VBA.StrComp(ControlTag, DP_CONTEXT_MENU_TAG, vbBinaryCompare) = 0 Then
                    M_ContextMenu_ContainsDatePicker = True
                    Exit Function
                End If
        Next ControlItem

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Release object reference
        Set ControlItem = Nothing
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Release object reference
        Set ControlItem = Nothing
    'Return the safe default
        M_ContextMenu_ContainsDatePicker = False
    'Write diagnostics without interrupting context-menu synchronization
        Debug.Print PROC_NAME & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed predicate error
        Err.Clear

End Function

'
'------------------------------------------------------------------------------
'
'                              VERIFIED TEARDOWN
'
'------------------------------------------------------------------------------
'

Private Function M_ContextMenu_BarIsClean( _
    ByVal CommandBarName As String, _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'             ARE ALL SAME-NAMED COMMAND BARS FREE OF DATEPICKER CONTROLS
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether any command bar with the given name still carries a
'   DatePicker control
'
' WHY THIS EXISTS
'   M_ContextMenu_Remove suppresses its own errors, so calling it proves nothing.
'   The registration has to be read back, and it lives on several bars: Excel
'   has more than one bar named "Cell" (#114)
'
' INPUTS
'   CommandBarName
'     Name of the bars to inspect
'
'   ErrorNumber, ErrorDescription
'     Receive the failure when no bar resolves or a control remains
'
' RETURNS
'   True when at least one bar resolves and none carries a tagged control
'
' BEHAVIOR
'   Resolves every bar with the name and scans each one's controls for the
'   DatePicker tag
'
' ERROR POLICY
'   Never raises outward. An unresolvable bar is a failure, not a clean result
'
' DEPENDENCIES
'   M_ContextMenu_GetCommandBars
'   DP_CONTEXT_MENU_TAG
'
' NOTES
'   An unresolvable bar is deliberately not treated as clean. A bar that cannot
'   be read might still hold a control, and this transaction fails closed
'
'   Moved from M_DP_Core by #57, where it was named
'   M_Lifecycle_ContextMenuBarIsClean, so the module that owns this
'   integration's state also verifies its teardown
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

    Dim TargetBars As Collection
    Dim TargetCommandBar As Variant
    Dim ControlItem As CommandBarControl
    Dim ControlTag As String

    On Error GoTo Failed
    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString

    Set TargetBars = M_ContextMenu_GetCommandBars(CommandBarName)
    If TargetBars.Count = 0 Then
        ErrorNumber = vbObjectError + 2715
        ErrorDescription = "Command bar could not be resolved: " & CommandBarName
        Set TargetBars = Nothing
        Exit Function
    End If

    For Each TargetCommandBar In TargetBars
        For Each ControlItem In TargetCommandBar.Controls
            ControlTag = VBA.CStr(ControlItem.Tag)
            If VBA.StrComp(ControlTag, DP_CONTEXT_MENU_TAG, vbBinaryCompare) = 0 Then
                ErrorNumber = vbObjectError + 2716
                ErrorDescription = "DatePicker control remained on command bar: " & _
                    CommandBarName & " (index " & VBA.CStr(TargetCommandBar.Index) & ")"
                Set ControlItem = Nothing
                Set TargetCommandBar = Nothing
                Set TargetBars = Nothing
                Exit Function
            End If
        Next ControlItem
    Next TargetCommandBar

    M_ContextMenu_BarIsClean = True
    Set ControlItem = Nothing
    Set TargetCommandBar = Nothing
    Set TargetBars = Nothing
    Exit Function

Failed:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    Set ControlItem = Nothing
    Set TargetCommandBar = Nothing
    Set TargetBars = Nothing
    Err.Clear
    M_ContextMenu_BarIsClean = False

End Function

Public Function M_ContextMenu_TryRemove( _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'                REMOVE THE CONTEXT MENU, VERIFIED ON BOTH BARS
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the DatePicker context-menu entries and verifies both bars
'
' WHY THIS EXISTS
'   The component registers on the cell bar and the table bar. Verifying only one
'   would report a clean teardown while the other still carried an entry
'
' INPUTS
'   ErrorNumber, ErrorDescription
'     Receive the first bar failure observed
'
' RETURNS
'   True only when both bars verify clean
'
' BEHAVIOR
'   Calls M_ContextMenu_Remove once, then checks the Cell and List Range Popup
'   bars independently and reports the first failure
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_ContextMenu_Remove
'   M_ContextMenu_BarIsClean
'
' NOTES
'   Both bars are checked even when the first fails, so the trace records the
'   full picture rather than stopping at the first bad bar
'
'   Moved from M_DP_Core by #57, where it was named
'   M_Lifecycle_TryRemoveContextMenu, so the module that owns this
'   integration's state also verifies its teardown
'
'   Public only so the lifecycle cleanup in M_DP_Core can call it
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

    Dim LocalErrNumber As Long
    Dim LocalErrDescription As String
    Dim CellClean As Boolean
    Dim ListClean As Boolean

    On Error GoTo Failed
    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString

    M_ContextMenu_Remove

    CellClean = M_ContextMenu_BarIsClean( _
        "Cell", LocalErrNumber, LocalErrDescription)
    If Not CellClean Then
        ErrorNumber = LocalErrNumber
        ErrorDescription = LocalErrDescription
    End If

    LocalErrNumber = 0
    LocalErrDescription = VBA.vbNullString
    ListClean = M_ContextMenu_BarIsClean( _
        "List Range Popup", LocalErrNumber, LocalErrDescription)
    If Not ListClean And ErrorNumber = 0 Then
        ErrorNumber = LocalErrNumber
        ErrorDescription = LocalErrDescription
    End If

    M_ContextMenu_TryRemove = (CellClean And ListClean)
    Exit Function

Failed:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    Err.Clear
    M_ContextMenu_TryRemove = False

End Function
