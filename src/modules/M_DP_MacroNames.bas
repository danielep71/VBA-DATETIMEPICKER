Attribute VB_Name = "M_DP_MacroNames"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_DP_MacroNames
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the workbook-qualified procedure names that Excel callbacks run,
'   and caches them per workbook name
'
' WHY THIS EXISTS
'   OnAction, OnKey and OnTime resolve a bare name against every open
'   workbook. Qualifying it with this workbook's name makes each callback
'   reach this project, and the cache keeps repeated UI paths cheap (#56)
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - M_GetQualifiedMacroName, which returns 'Workbook'!Procedure
'     - M_GetQualifiedMacroName_ClearCache, which releases the cache
'
' ERROR POLICY
'   Each procedure documents its own policy
'
' DEPENDENCIES
'   ThisWorkbook
'   Scripting.Dictionary through late binding
'
' NOTES
'   The cache is keyed by ThisWorkbook.Name and rebuilt when the name
'   changes, so a Save As or rename still yields callable names
'
'   Option Private Module keeps both procedures out of other VBA projects
'   and out of the macro dialog. They are Public only so the other
'   implementation modules can call them
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mQualifiedMacroNameCache    As Object               'Scripting.Dictionary: ProcedureName -> qualified macro name
    Private mQualifiedMacroWorkbookName As String               'Workbook name used by the qualified macro-name cache


'
'------------------------------------------------------------------------------
'
'                            QUALIFIED MACRO NAMES
'
'------------------------------------------------------------------------------
'

Public Function M_GetQualifiedMacroName(ByVal ProcedureName As String) As String

'
'------------------------------------------------------------------------------
'                           GET QUALIFIED MACRO NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a workbook-qualified macro name suitable for Excel callbacks
'
' WHY THIS EXISTS
'   Shape.OnAction, CommandBarButton.OnAction, Application.OnTime, Application.
'   OnKey, and Ribbon-related callback paths should resolve back to the workbook
'   that contains this module
'
'   This helper also caches resolved callback names because it is used by repeated
'   UI integration paths such as grid-icon movement, timer scheduling, keyboard
'   shortcut registration, and context-menu creation
'
' INPUTS
'   ProcedureName
'     Public procedure name to qualify
'
' RETURNS
'   Workbook-qualified macro name
'
' BEHAVIOR
'   Trims the supplied procedure name
'   Rejects a blank procedure name
'   Initializes a late-bound dictionary cache when needed
'   Invalidates the cache automatically when ThisWorkbook.Name changes
'   Returns a cached qualified name when available
'   Otherwise builds, stores, and returns the qualified macro name
'
' ERROR POLICY
'   Raises a descriptive runtime error when ProcedureName is blank or when the
'   callback name cannot be resolved
'
' DEPENDENCIES
'   ThisWorkbook
'   Scripting.Dictionary through late binding
'   VBA.Replace
'   VBA.Trim$
'
' NOTES
'   Expected output format:
'     'WorkbookName.xlsm'!ProcedureName
'
'   The workbook name is escaped by doubling apostrophes
'
'   The cache is workbook-name aware, so Save As / rename scenarios are handled
'   without requiring a manual reset
'
'   Public only so the other implementation modules can call it. Option
'   Private Module keeps it out of other projects and the macro dialog, and
'   callers still do not depend on the cache implementation
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME                 As String = "M_GetQualifiedMacroName"

    Dim NormalizedProcedureName     As String       'Trimmed callback procedure name
    Dim CurrentWorkbookName         As String       'Current host workbook name
    Dim QualifiedMacroName          As String       'Resolved workbook-qualified macro name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the supplied procedure name once
        NormalizedProcedureName = VBA.Trim$(ProcedureName)
    'Reject an empty procedure name
        If VBA.Len(NormalizedProcedureName) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ProcedureName cannot be empty"
        End If

'------------------------------------------------------------------------------
' RESOLVE WORKBOOK NAME
'------------------------------------------------------------------------------
    'Read the current host workbook name
        CurrentWorkbookName = VBA.CStr(ThisWorkbook.Name)
    'Reject an empty workbook name
        If VBA.Len(CurrentWorkbookName) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ThisWorkbook.Name cannot be empty"
        End If

'------------------------------------------------------------------------------
' INITIALIZE OR REFRESH CACHE
'------------------------------------------------------------------------------
    'Create the cache when it does not exist
        If mQualifiedMacroNameCache Is Nothing Then
            Set mQualifiedMacroNameCache = VBA.CreateObject("Scripting.Dictionary")
            mQualifiedMacroNameCache.CompareMode = vbTextCompare
            mQualifiedMacroWorkbookName = CurrentWorkbookName
        End If
    'Recreate the cache when the workbook name changed
        If VBA.StrComp(mQualifiedMacroWorkbookName, CurrentWorkbookName, vbBinaryCompare) <> 0 Then
            Set mQualifiedMacroNameCache = VBA.CreateObject("Scripting.Dictionary")
            mQualifiedMacroNameCache.CompareMode = vbTextCompare
            mQualifiedMacroWorkbookName = CurrentWorkbookName
        End If

'------------------------------------------------------------------------------
' RETURN CACHED VALUE
'------------------------------------------------------------------------------
    'Return the cached qualified macro name when available
        If mQualifiedMacroNameCache.Exists(NormalizedProcedureName) Then
            M_GetQualifiedMacroName = VBA.CStr(mQualifiedMacroNameCache(NormalizedProcedureName))
            Exit Function
        End If

'------------------------------------------------------------------------------
' BUILD QUALIFIED NAME
'------------------------------------------------------------------------------
    'Build the workbook-qualified macro name
        QualifiedMacroName = _
            "'" & VBA.Replace(CurrentWorkbookName, "'", "''") & "'!" & NormalizedProcedureName

'------------------------------------------------------------------------------
' CACHE QUALIFIED NAME
'------------------------------------------------------------------------------
    'Store the resolved callback name for reuse
        mQualifiedMacroNameCache.Add NormalizedProcedureName, QualifiedMacroName

'------------------------------------------------------------------------------
' RETURN QUALIFIED NAME
'------------------------------------------------------------------------------
    'Return the resolved callback name
        M_GetQualifiedMacroName = QualifiedMacroName

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, _
            "Qualified macro-name resolution failed: " & Err.Description

End Function

Public Sub M_GetQualifiedMacroName_ClearCache()

'
'------------------------------------------------------------------------------
'                       CLEAR QUALIFIED MACRO NAME CACHE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears cached workbook-qualified macro names
'
' WHY THIS EXISTS
'   Callback names are session-level helper values. They should be releasable
'   during DatePicker teardown, runtime repair, or workbook lifecycle cleanup
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Releases the late-bound dictionary cache and clears the workbook-name marker
'
' ERROR POLICY
'   Best-effort cleanup
'
'   Suppresses errors because cache cleanup should never interrupt workbook
'   close, DatePicker stop, or runtime repair logic
'
' DEPENDENCIES
'   mQualifiedMacroNameCache
'   mQualifiedMacroWorkbookName
'
' NOTES
'   The cache also self-invalidates when ThisWorkbook.Name changes, so this
'   helper is cleanup-oriented rather than required for correctness
'
'   Public only so the lifecycle cleanup in M_DP_Core can call it
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' CLEAR CACHE
'------------------------------------------------------------------------------
    'Release the qualified macro-name cache
        Set mQualifiedMacroNameCache = Nothing
    'Clear the cached workbook-name marker
        mQualifiedMacroWorkbookName = vbNullString

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub
