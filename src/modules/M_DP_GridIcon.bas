Attribute VB_Name = "M_DP_GridIcon"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_DP_GridIcon
'------------------------------------------------------------------------------
' PURPOSE
'   Creates, places, hides and removes the in-grid calendar icon shape
'
' WHY THIS EXISTS
'   The icon is a worksheet shape that another workbook or a user can
'   also name DP_GridIcon. One module owns its lifecycle, placement and
'   ownership marker, so no code touches a shape it cannot prove is
'   its own (#57)
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - show, move, hide, pre-create and remove for the tracked icon
'     - ownership markers and the refusal of foreign shapes
'     - M_GridIcon_PurgeAll_Impl and M_GridIcon_TryPurgeAll, which delete
'       every owned icon in every open workbook
'
' ERROR POLICY
'   Each procedure documents its own policy
'
' DEPENDENCIES
'   Excel Shapes object model
'   M_DP_GridIconAsset (M_GridIcon_EnsureEmbeddedIconFile)
'   M_DP_MacroNames (M_GetQualifiedMacroName)
'   M_DP_Core (gDP_ShowGridIcon, DP_THEME_COLOR_HEADER,
'     M_Settings_EnsureLoaded, M_Lease_LocalOwnerId, M_Lease_NewOwnerId)
'
' NOTES
'   The icon runs the facade callback DP_Click. The image bytes live in
'   M_DP_GridIconAsset, apart from placement
'
'   DP_GRID_ICON_NAME and gDP_GridIconShape are Public so the regression
'   harness can observe the icon
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
    Private Const DP_GRID_ICON_FONT_NAME          As String = "Segoe MDL2 Assets"           'Grid icon font name
    Private Const DP_GRID_ICON_CALENDAR_CODEPOINT As Long = &HE8BF                          'Segoe MDL2 calendar glyph
    Private Const DP_GRID_ICON_FONT_SIZE          As Single = 10                            'Grid icon font size
    Private Const DP_GRID_ICON_FORE_COLOR         As Long = vbWhite                         'Grid icon foreground color
    Private Const DP_GRID_ICON_ALT_BASE           As String = "DatePicker Grid Entry Point" 'Human-readable grid icon alternative text
    Private Const DP_GRID_ICON_OWNER_TAG          As String = " | dp-owner-v1="             'Versioned ownership marker schema tag
    Private Const DP_GRID_ICON_PENDING_SUFFIX     As String = "_Pending"                    'Temporary name held during icon creation

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    Public Const DP_GRID_ICON_NAME As String = "DP_GridIcon"    'Worksheet grid icon shape name

'------------------------------------------------------------------------------
' PUBLIC STATE
'------------------------------------------------------------------------------
    Public gDP_GridIconShape As Shape   'In-grid icon shape reference

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mDP_GridIconOwnerToken    As String 'Durable token stamped into created grid icons, never used for admission
    Private mDP_GridIconRefusalKey    As String 'Diagnostic-only key of the last sheet where a foreign icon refused creation
    Private mDP_GridIconLastAnchorKey As String 'Last grid-icon target key
    Private mDP_GridIconLastLeft      As Double 'Last grid-icon left position
    Private mDP_GridIconLastTop       As Double 'Last grid-icon top position


'
'------------------------------------------------------------------------------
'
'                                  GRID ICON
'
'------------------------------------------------------------------------------
'

Private Function M_GridIcon_BuildAnchorKey(ByVal AnchorCell As Excel.Range) As String

'
'------------------------------------------------------------------------------
'                           BUILD GRID ICON ANCHOR KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Builds a stable key for the current grid-icon anchor cell
'
' WHY THIS EXISTS
'   The in-grid icon can receive repeated selection-change refreshes for the same
'   cell. A compact target key allows the high-frequency path to detect that the
'   icon is already shown at the correct target and exit early.
'
' INPUTS
'   AnchorCell
'     Normalized one-cell anchor range
'
' RETURNS
'   External cell address key, or vbNullString when the anchor cannot be resolved
'
' BEHAVIOR
'   Returns the external A1 address of the supplied anchor cell
'
' ERROR POLICY
'   Safe-default helper. Returns vbNullString on failure
'
' DEPENDENCIES
'   Excel.Range.Address
'
' NOTES
'   The caller is responsible for normalizing merged cells before calling this
'   helper
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return a safe default on failure
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit when no anchor cell is supplied
        If AnchorCell Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' RETURN ANCHOR KEY
'------------------------------------------------------------------------------
    'Return a workbook-qualified worksheet address
        M_GridIcon_BuildAnchorKey = AnchorCell.Address( _
            RowAbsolute:=False, _
            ColumnAbsolute:=False, _
            ReferenceStyle:=xlA1, _
            External:=True)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful key creation
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Return safe default
        M_GridIcon_BuildAnchorKey = vbNullString

End Function

Private Sub M_GridIcon_ClearLastTarget()

'
'------------------------------------------------------------------------------
'                         CLEAR GRID ICON LAST TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the cached last grid-icon target
'
' WHY THIS EXISTS
'   The same-target short-circuit must be invalidated when the tracked shape is
'   deleted, purged, or found stale
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the cached anchor key and cached icon position
'
' ERROR POLICY
'   Does not raise errors
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR CACHE
'------------------------------------------------------------------------------
    'Clear the cached anchor key
        mDP_GridIconLastAnchorKey = vbNullString
    'Clear the cached left position
        mDP_GridIconLastLeft = 0#
    'Clear the cached top position
        mDP_GridIconLastTop = 0#

End Sub

Private Sub M_GridIcon_RememberTarget( _
    ByVal AnchorKey As String, _
    ByVal IconLeft As Double, _
    ByVal IconTop As Double)

'
'------------------------------------------------------------------------------
'                         REMEMBER GRID ICON TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Stores the last successfully positioned grid-icon target
'
' WHY THIS EXISTS
'   The cached target allows M_GridIcon_ShowOrMove to skip unnecessary shape
'   property writes when Excel fires repeated refreshes for the same cell
'
' INPUTS
'   AnchorKey
'     Stable anchor-cell key
'
'   IconLeft
'     Icon left position in points
'
'   IconTop
'     Icon top position in points
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the supplied key and position when the key is not blank
'
' ERROR POLICY
'   Does not raise errors
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' STORE CACHE
'------------------------------------------------------------------------------
    'Clear cache when the supplied key is blank
        If VBA.LenB(AnchorKey) = 0 Then
            M_GridIcon_ClearLastTarget
            Exit Sub
        End If

    'Store the anchor key
        mDP_GridIconLastAnchorKey = AnchorKey
    'Store the icon left position
        mDP_GridIconLastLeft = IconLeft
    'Store the icon top position
        mDP_GridIconLastTop = IconTop

End Sub

Private Function M_GridIcon_IsSameVisibleTarget( _
    ByVal AnchorKey As String, _
    ByVal TargetSheet As Excel.Worksheet, _
    ByVal IconLeft As Double, _
    ByVal IconTop As Double, _
    ByVal IconSize As Double) As Boolean

'
'------------------------------------------------------------------------------
'                       GRID ICON IS SAME VISIBLE TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the tracked grid icon is already visible at the requested
'   target
'
' WHY THIS EXISTS
'   SelectionChange can fire repeatedly while the DatePicker icon is already
'   positioned correctly. In that case, moving the shape, resetting properties,
'   and bringing it to front again is unnecessary work.
'
' INPUTS
'   AnchorKey
'     Stable anchor-cell key
'
'   TargetSheet
'     Worksheet that should own the icon
'
'   IconLeft
'     Expected icon left position
'
'   IconTop
'     Expected icon top position
'
'   IconSize
'     Expected icon width and height
'
' RETURNS
'   True when the tracked icon is already valid, visible, and correctly placed
'
'   False otherwise
'
' BEHAVIOR
'   Checks the cached target key first, then validates the tracked shape, parent
'   worksheet, visibility, position, and size
'
' ERROR POLICY
'   Safe-default predicate. Returns False and clears stale tracking state on
'   failure
'
' DEPENDENCIES
'   gDP_GridIconShape
'   M_GridIcon_ClearLastTarget
'
' NOTES
'   Position comparisons use a small tolerance to avoid point-rounding noise
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const POSITION_TOLERANCE    As Double = 0.05       'Point comparison tolerance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return safe default unless all checks pass
        M_GridIcon_IsSameVisibleTarget = False
    'Enable safe-default error handling
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' VALIDATE CACHE
'------------------------------------------------------------------------------
    'Exit when the requested key is blank
        If VBA.LenB(AnchorKey) = 0 Then Exit Function
    'Exit when the target sheet is missing
        If TargetSheet Is Nothing Then Exit Function
    'Exit when the tracked icon is missing
        If Not M_GridIcon_TrackedShapeIsLive() Then Exit Function
    'Exit when the cached target key differs
        If VBA.StrComp(mDP_GridIconLastAnchorKey, AnchorKey, vbBinaryCompare) <> 0 Then Exit Function
    'Exit when the cached left position differs
        If VBA.Abs(mDP_GridIconLastLeft - IconLeft) > POSITION_TOLERANCE Then Exit Function
    'Exit when the cached top position differs
        If VBA.Abs(mDP_GridIconLastTop - IconTop) > POSITION_TOLERANCE Then Exit Function

'------------------------------------------------------------------------------
' VALIDATE TRACKED SHAPE
'------------------------------------------------------------------------------
    'Exit when the tracked icon belongs to another worksheet
        If Not (gDP_GridIconShape.Parent Is TargetSheet) Then Exit Function
    'Exit when the tracked icon is hidden
        If gDP_GridIconShape.Visible <> msoTrue Then Exit Function
    'Exit when the tracked icon left position differs
        If VBA.Abs(gDP_GridIconShape.Left - IconLeft) > POSITION_TOLERANCE Then Exit Function
    'Exit when the tracked icon top position differs
        If VBA.Abs(gDP_GridIconShape.Top - IconTop) > POSITION_TOLERANCE Then Exit Function
    'Exit when the tracked icon width differs
        If VBA.Abs(gDP_GridIconShape.Width - IconSize) > POSITION_TOLERANCE Then Exit Function
    'Exit when the tracked icon height differs
        If VBA.Abs(gDP_GridIconShape.Height - IconSize) > POSITION_TOLERANCE Then Exit Function

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Return True because the existing visible icon already matches the target
        M_GridIcon_IsSameVisibleTarget = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful validation
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Suppress stale-shape cleanup errors
        On Error Resume Next
    'Clear the tracked shape reference
        Set gDP_GridIconShape = Nothing
    'Clear the cached target
        M_GridIcon_ClearLastTarget
    'Return safe default
        M_GridIcon_IsSameVisibleTarget = False
    'Restore normal error handling
        On Error GoTo 0

End Function

Private Function M_GridIcon_SessionOwnerToken() As String

'
'==============================================================================
'                        GRID ICON SESSION OWNER TOKEN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the token stamped into grid icons created by this session
'
' WHY THIS EXISTS
'   mDP_RuntimeOwnerId is ephemeral and is empty whenever this project does not
'   hold the provider lease. An icon stamped with an empty token would be born
'   unowned, and fail-closed cleanup would then protect it from its own creator
'   permanently. This token is therefore never empty
'
' INPUTS
'   None
'
' RETURNS
'   A token valid under the v1 owner-token grammar
'
' BEHAVIOR
'   Seeds once from the live lease token when one is held, otherwise mints one
'   through the same generator the lease uses, and reuses it for the session
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_Lease_LocalOwnerId (M_DP_Core)
'   mDP_GridIconOwnerToken
'   M_Lease_NewOwnerId (M_DP_Core)
'
' NOTES
'   This is a stamping identity only. It is never consulted for admission and
'   never compared against the lease. #14 will compare stored tokens against the
'   live provider; nothing here should be read as ownership arbitration
'
'   The lease token is read through M_Lease_LocalOwnerId, so the lease state
'   stays owned by M_DP_Core (#57)
'
' UPDATED
'   2026-11-21
'==============================================================================

'------------------------------------------------------------------------------
' RESOLVE THE SESSION TOKEN
'------------------------------------------------------------------------------
    'Never let token resolution raise into a caller
        On Error Resume Next
    'Reuse the token already minted for this session
        If VBA.LenB(mDP_GridIconOwnerToken) > 0 Then
            M_GridIcon_SessionOwnerToken = mDP_GridIconOwnerToken
            Err.Clear
            Exit Function
        End If
    'Prefer the live lease token when this project holds the lease
        If VBA.LenB(M_Lease_LocalOwnerId()) > 0 Then
            mDP_GridIconOwnerToken = M_Lease_LocalOwnerId()
        Else
    'Mint a token through the same generator the lease uses
            mDP_GridIconOwnerToken = M_Lease_NewOwnerId()
        End If
    'Return the resolved token
        M_GridIcon_SessionOwnerToken = mDP_GridIconOwnerToken
    'Clear any suppressed resolution error
        Err.Clear

End Function

Private Function M_GridIcon_BuildOwnerMarker() As String

'
'==============================================================================
'                        GRID ICON BUILD OWNER MARKER
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the AlternativeText value written into every DatePicker grid icon
'
' WHY THIS EXISTS
'   The marker is the only durable, shape-scoped ownership evidence. Building it
'   in one place keeps the written form and the parsed form from drifting
'
' INPUTS
'   None
'
' RETURNS
'   The complete marker string
'
' BEHAVIOR
'   Concatenates the human-readable base, the versioned schema tag and the
'   session owner token
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   DP_GRID_ICON_ALT_BASE
'   DP_GRID_ICON_OWNER_TAG
'   M_GridIcon_SessionOwnerToken
'
' NOTES
'   The human phrase stays at the front so assistive technology announces
'   something meaningful before the machine token. Formalizing or relocating the
'   accessible representation belongs to #29
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' BUILD THE MARKER
'------------------------------------------------------------------------------
    'Never let marker construction raise into a caller
        On Error Resume Next
    'Compose the versioned marker
        M_GridIcon_BuildOwnerMarker = DP_GRID_ICON_ALT_BASE & _
            DP_GRID_ICON_OWNER_TAG & _
            M_GridIcon_SessionOwnerToken()
    'Clear any suppressed construction error
        Err.Clear

End Function

Private Function M_GridIcon_IsValidOwnerToken(ByVal TokenText As String) As Boolean

'
'==============================================================================
'                      GRID ICON IS VALID OWNER TOKEN
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a token satisfies the v1 owner-token grammar
'
' WHY THIS EXISTS
'   A bare prefix must not confer ownership. Requiring the token to parse stops
'   hand-typed or truncated alternative text from manufacturing ownership
'
' INPUTS
'   TokenText
'     Candidate token extracted from a marker
'
' RETURNS
'   True only when the token matches the v1 grammar exactly
'
' BEHAVIOR
'   Validates yyyymmddhhnnss-nnnnnnnn-HEX, which is the form produced by
'   M_Lease_NewOwnerId
'
' ERROR POLICY
'   Never raises outward. Any failure reports False
'
' DEPENDENCIES
'   None
'
' NOTES
'   The grammar is bound to schema v1. A future schema adds its own validator
'   rather than loosening this one, so shapes already marked keep their meaning
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CharIndex       As Long         'Character position being validated
    Dim CurrentChar     As String       'Character being validated
    Dim HexPart         As String       'Trailing hexadecimal component

'------------------------------------------------------------------------------
' VALIDATE THE TOKEN
'------------------------------------------------------------------------------
    'Never let validation raise into a caller
        On Error Resume Next
    'Assume the token is invalid until every rule passes
        M_GridIcon_IsValidOwnerToken = False
    'Reject a token shorter than the minimum v1 form
        If VBA.Len(TokenText) < 25 Then GoTo ExitProcedure
    'Reject a token longer than the maximum v1 form
        If VBA.Len(TokenText) > 32 Then GoTo ExitProcedure
    'Require the separators in their fixed positions
        If VBA.Mid$(TokenText, 15, 1) <> "-" Then GoTo ExitProcedure
        If VBA.Mid$(TokenText, 24, 1) <> "-" Then GoTo ExitProcedure
    'Require 14 leading timestamp digits
        For CharIndex = 1 To 14
            CurrentChar = VBA.Mid$(TokenText, CharIndex, 1)
            If CurrentChar < "0" Or CurrentChar > "9" Then GoTo ExitProcedure
        Next CharIndex
    'Require 8 fractional-second digits
        For CharIndex = 16 To 23
            CurrentChar = VBA.Mid$(TokenText, CharIndex, 1)
            If CurrentChar < "0" Or CurrentChar > "9" Then GoTo ExitProcedure
        Next CharIndex
    'Isolate the trailing hexadecimal component
        HexPart = VBA.Mid$(TokenText, 25)
    'Require a non-empty hexadecimal component
        If VBA.LenB(HexPart) = 0 Then GoTo ExitProcedure
    'Require every hexadecimal character to be an uppercase hex digit
        For CharIndex = 1 To VBA.Len(HexPart)
            CurrentChar = VBA.Mid$(HexPart, CharIndex, 1)
            If Not ((CurrentChar >= "0" And CurrentChar <= "9") Or _
                    (CurrentChar >= "A" And CurrentChar <= "F")) Then
                GoTo ExitProcedure
            End If
        Next CharIndex
    'Report a token that satisfies every rule
        M_GridIcon_IsValidOwnerToken = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Clear any suppressed validation error
        Err.Clear

End Function

Private Function M_GridIcon_ReadOwnerToken(ByVal TargetShape As Excel.Shape) As String

'
'==============================================================================
'                        GRID ICON READ OWNER TOKEN
'------------------------------------------------------------------------------
' PURPOSE
'   Extracts the owner token from a shape carrying a v1 ownership marker
'
' WHY THIS EXISTS
'   #14 must compare stored tokens against the live provider. Extraction is kept
'   beside the predicate so no other routine parses AlternativeText
'
' INPUTS
'   TargetShape
'     Shape to inspect
'
' RETURNS
'   The token when the shape carries a valid v1 marker, otherwise an empty string
'
' BEHAVIOR
'   Reads AlternativeText, verifies the base and schema tag, and returns the
'   remainder only when it parses under the v1 grammar
'
' ERROR POLICY
'   Never raises outward. Any failure returns an empty string
'
' DEPENDENCIES
'   DP_GRID_ICON_ALT_BASE
'   DP_GRID_ICON_OWNER_TAG
'   M_GridIcon_IsValidOwnerToken
'
' NOTES
'   A legacy v0 icon carries the base string with no tag. It is owned but has no
'   token, so this function returns an empty string for it. An empty result
'   therefore means "no token", never "not owned"
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MarkerText      As String       'Alternative text read from the shape
    Dim MarkerPrefix    As String       'Expected base and schema tag
    Dim TokenText       As String       'Extracted candidate token

'------------------------------------------------------------------------------
' READ THE TOKEN
'------------------------------------------------------------------------------
    'Never let a marker read raise into a caller
        On Error Resume Next
    'Start with no token
        M_GridIcon_ReadOwnerToken = VBA.vbNullString
    'Exit when no shape was supplied
        If TargetShape Is Nothing Then GoTo ExitProcedure
    'Read the marker, treating an unreadable shape as unmarked
        Err.Clear
        MarkerText = TargetShape.AlternativeText
        If Err.Number <> 0 Then GoTo ExitProcedure
    'Build the expected marker prefix
        MarkerPrefix = DP_GRID_ICON_ALT_BASE & DP_GRID_ICON_OWNER_TAG
    'Exit when the marker does not open with the expected prefix
        If VBA.Len(MarkerText) <= VBA.Len(MarkerPrefix) Then GoTo ExitProcedure
        If VBA.StrComp(VBA.Left$(MarkerText, VBA.Len(MarkerPrefix)), _
            MarkerPrefix, vbBinaryCompare) <> 0 Then GoTo ExitProcedure
    'Isolate the candidate token
        TokenText = VBA.Mid$(MarkerText, VBA.Len(MarkerPrefix) + 1)
    'Return the token only when it parses
        If M_GridIcon_IsValidOwnerToken(TokenText) Then
            M_GridIcon_ReadOwnerToken = TokenText
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Clear any suppressed read error
        Err.Clear

End Function

Private Function M_GridIcon_IsOwnedShape(ByVal TargetShape As Excel.Shape) As Boolean

'
'==============================================================================
'                          GRID ICON IS OWNED SHAPE
'------------------------------------------------------------------------------
' PURPOSE
'   The single authority on whether a shape is a DatePicker-owned grid icon
'
' WHY THIS EXISTS
'   Before this predicate the component treated the shape name alone as proof of
'   ownership. It would adopt any same-named shape, move it, resize it, bind it
'   to DP_Click, overwrite its alternative text and later delete it. Overwriting
'   the alternative text manufactured the very evidence the deletion relied on,
'   so ownership could be self-granted. Interpretation must live in exactly one
'   place or that drift returns
'
' INPUTS
'   TargetShape
'     Shape to inspect
'
' RETURNS
'   True only when ownership is proven
'
' BEHAVIOR
'   Requires the canonical or pending grid-icon name, then a marker that is
'   either a valid v1 marker or the legacy v0 alternative text
'
' ERROR POLICY
'   Never raises outward. Any failure to read reports False, so an unreadable
'   shape is treated as foreign and is never modified or deleted
'
' DEPENDENCIES
'   DP_GRID_ICON_NAME
'   DP_GRID_ICON_PENDING_SUFFIX
'   DP_GRID_ICON_ALT_BASE
'   M_GridIcon_ReadOwnerToken
'
' NOTES
'   Legacy v0 icons carry the bare base string with no schema tag. They predate
'   #53 and are recognized as product-owned so the component does not abandon
'   icons it created before the marker existed
'
'   Ownership here is product-level by design until #14, which is what allows a
'   stale icon left by a crashed provider to be reclaimed. #14 narrows the
'   predicate to token equality without changing the carrier
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ShapeName       As String       'Name read from the shape
    Dim MarkerText      As String       'Alternative text read from the shape

'------------------------------------------------------------------------------
' VERIFY OWNERSHIP
'------------------------------------------------------------------------------
    'Never let an ownership probe raise into a caller
        On Error Resume Next
    'Fail closed until ownership is proven
        M_GridIcon_IsOwnedShape = False
    'Exit when no shape was supplied
        If TargetShape Is Nothing Then GoTo ExitProcedure
    'Read the shape name, treating an unreadable shape as foreign
        Err.Clear
        ShapeName = TargetShape.Name
        If Err.Number <> 0 Then GoTo ExitProcedure
    'Require the canonical or pending grid-icon name
        If VBA.StrComp(ShapeName, DP_GRID_ICON_NAME, vbBinaryCompare) <> 0 Then
            If VBA.StrComp(ShapeName, _
                DP_GRID_ICON_NAME & DP_GRID_ICON_PENDING_SUFFIX, _
                vbBinaryCompare) <> 0 Then
                GoTo ExitProcedure
            End If
        End If
    'Accept a shape carrying a valid v1 owner token
        If VBA.LenB(M_GridIcon_ReadOwnerToken(TargetShape)) > 0 Then
            M_GridIcon_IsOwnedShape = True
            GoTo ExitProcedure
        End If
    'Read the marker again to test the legacy v0 form
        Err.Clear
        MarkerText = TargetShape.AlternativeText
        If Err.Number <> 0 Then GoTo ExitProcedure
    'Accept the legacy v0 marker, which carries no token
        If VBA.StrComp(MarkerText, DP_GRID_ICON_ALT_BASE, vbBinaryCompare) = 0 Then
            M_GridIcon_IsOwnedShape = True
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Clear any suppressed probe error
        Err.Clear

End Function

Private Function M_GridIcon_ResolveOwnedShape( _
    ByVal TargetSheet As Excel.Worksheet, _
    ByRef ForeignShapePresent As Boolean) As Excel.Shape

'
'==============================================================================
'                        GRID ICON RESOLVE OWNED SHAPE
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the canonical grid icon on a worksheet, reporting separately when a
'   same-named shape exists but is not owned
'
' WHY THIS EXISTS
'   Callers need three answers, not two: an owned icon to reuse, no shape at all,
'   or a foreign shape that must be left completely alone. Collapsing the last
'   two would let creation overwrite a user shape
'
' INPUTS
'   TargetSheet
'     Worksheet to inspect
'
'   ForeignShapePresent
'     Set True when a same-named shape exists and ownership is not proven
'
' RETURNS
'   The owned shape, or Nothing
'
' BEHAVIOR
'   Looks the shape up by canonical name and routes it through the ownership
'   predicate without touching it
'
' ERROR POLICY
'   Never raises outward. A lookup failure reports no shape and no collision
'
' DEPENDENCIES
'   DP_GRID_ICON_NAME
'   M_GridIcon_IsOwnedShape
'
' NOTES
'   This routine reads only. It must never modify the shape it inspects, because
'   the foreign case is decided by what it finds
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CandidateShape  As Excel.Shape  'Same-named shape found on the worksheet

'------------------------------------------------------------------------------
' RESOLVE THE SHAPE
'------------------------------------------------------------------------------
    'Never let resolution raise into a caller
        On Error Resume Next
    'Report no collision until one is found
        ForeignShapePresent = False
    'Exit when no worksheet was supplied
        If TargetSheet Is Nothing Then GoTo ExitProcedure
    'Look the canonical shape up without touching it
        Err.Clear
        Set CandidateShape = TargetSheet.Shapes(DP_GRID_ICON_NAME)
    'Exit when no same-named shape exists
        If Err.Number <> 0 Or CandidateShape Is Nothing Then
            Err.Clear
            Set CandidateShape = Nothing
            GoTo ExitProcedure
        End If
    'Return an owned icon, or report a foreign collision
        If M_GridIcon_IsOwnedShape(CandidateShape) Then
            Set M_GridIcon_ResolveOwnedShape = CandidateShape
        Else
            ForeignShapePresent = True
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Release the local reference without releasing the returned shape
        Set CandidateShape = Nothing
    'Clear any suppressed resolution error
        Err.Clear

End Function

Private Sub M_GridIcon_ReportForeignShape( _
    ByVal TargetSheet As Excel.Worksheet, _
    ByVal EntryPoint As String)

'
'==============================================================================
'                       GRID ICON REPORT FOREIGN SHAPE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a diagnostic once per worksheet when a foreign same-named shape
'   refuses grid-icon creation
'
' WHY THIS EXISTS
'   Should_ShowGridIcon fires on every eligible selection, so an unsuppressed
'   refusal would print on every cell move. Silence would be worse: the user
'   would see no icon and no reason
'
' INPUTS
'   TargetSheet
'     Worksheet carrying the foreign shape
'
'   EntryPoint
'     Name of the refusing routine
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds a workbook-plus-worksheet identity key and prints only when that key
'   differs from the last reported one
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   mDP_GridIconRefusalKey
'
' NOTES
'   This state is diagnostic only. It never affects the ownership decision and
'   never suppresses re-evaluation: every eligible selection still consults the
'   predicate against current reality. Only repeated printing is suppressed
'
'   The key uses the worksheet CodeName as well as its visible name, because a
'   visible name can be renamed and can collide across workbooks
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RefusalKey      As String       'Workbook and worksheet identity key

'------------------------------------------------------------------------------
' REPORT THE REFUSAL
'------------------------------------------------------------------------------
    'Never let a diagnostic raise into a caller
        On Error Resume Next
    'Exit when no worksheet was supplied
        If TargetSheet Is Nothing Then GoTo ExitProcedure
    'Build the worksheet identity key
        RefusalKey = TargetSheet.Parent.Name & "||" & _
            TargetSheet.CodeName & "||" & TargetSheet.Name
    'Exit when this worksheet was already reported
        If VBA.StrComp(RefusalKey, mDP_GridIconRefusalKey, vbBinaryCompare) = 0 Then
            GoTo ExitProcedure
        End If
    'Remember the reported worksheet
        mDP_GridIconRefusalKey = RefusalKey
    'Write the refusal diagnostic
        Debug.Print EntryPoint & _
            " | Refused | A shape named " & DP_GRID_ICON_NAME & _
            " exists on " & TargetSheet.Name & _
            " without DatePicker ownership; it was left untouched"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Clear any suppressed diagnostic error
        Err.Clear

End Sub

Private Sub M_GridIcon_ClearRefusalKey()

'
'==============================================================================
'                        GRID ICON CLEAR REFUSAL KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the diagnostic refusal key so a later collision reports again
'
' WHY THIS EXISTS
'   A refusal that has been resolved should not stay silent forever. Clearing on
'   success and on teardown keeps the diagnostic useful without making it stateful
'   in any way that matters
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Empties the stored key
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   mDP_GridIconRefusalKey
'
' NOTES
'   Called when an owned icon is created or adopted, and during teardown. It is
'   never consulted for ownership
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' CLEAR THE KEY
'------------------------------------------------------------------------------
    'Never let a diagnostic reset raise into a caller
        On Error Resume Next
    'Forget the last reported worksheet
        mDP_GridIconRefusalKey = VBA.vbNullString
    'Clear any suppressed reset error
        Err.Clear

End Sub

Public Sub M_GridIcon_ShowOrMove(Optional ByVal TargetCell As Excel.Range)

'
'==============================================================================
'                         SHOW OR MOVE GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Shows or moves the DatePicker in-grid icon by reusing an existing worksheet
'   shape whenever possible
'
' WHY THIS EXISTS
'   SelectionChange can fire very frequently
'
'   Deleting and recreating the worksheet icon on every eligible cell selection
'   creates visible latency because Excel must rebuild and repaint the drawing-
'   layer shape
'
'   This routine is the optimized high-frequency path: it moves a reusable icon
'   when possible and creates a new icon only when needed
'
' INPUTS
'   TargetCell
'     Optional anchor range
'
'     If omitted, Excel.Application.ActiveCell is used
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures settings are loaded
'   Removes the icon and exits when the grid-icon feature is disabled
'   Resolves the target cell from TargetCell or ActiveCell
'   Normalizes the target to the first physical cell
'   Normalizes merged cells to the top-left cell of the merge area
'   Resolves the target worksheet
'   Calculates the icon position
'   Reuses the tracked icon when it belongs to the target worksheet
'   Deletes the tracked icon when it belongs to another worksheet
'   Reuses a same-named target-sheet icon when available
'   Updates only changed icon properties where practical
'   Creates the icon only when no reusable target-sheet shape exists
'
' ERROR POLICY
'   Best-effort UI routine
'
'   Suppresses failures, clears stale tracked references, falls back to the cold
'   creation path when possible, and writes diagnostics to the Immediate Window
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_GridIcon_Hide
'   M_GridIcon_Create
'   M_GetQualifiedMacroName
'   DP_GRID_ICON_NAME
'   gDP_ShowGridIcon
'   gDP_GridIconShape
'   Excel Shapes object model
'
' NOTES
'   This is the preferred high-frequency SelectionChange path
'
'   M_GridIcon_Create remains the cold-path creation routine
'
'   The manager refresh path should call this routine, not M_GridIcon_Create
'   directly, otherwise the reuse / move optimization is bypassed
'
'   Target normalization is intentionally aligned with M_GridIcon_Create
'
'   When a multi-cell range is supplied, the first physical cell is used as the
'   icon anchor
'
'   The icon is still brought to front on each move because other worksheet
'   shapes may have changed z-order since the last selection event
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_GridIcon_ShowOrMove"
    Const ICON_SIZE             As Double = 24#                     'Grid icon size in points
    Const ICON_GAP              As Double = 5#                      'Gap between target cell and icon

    Dim TargetSheet             As Excel.Worksheet   'Worksheet receiving the icon
    Dim AnchorCell              As Excel.Range       'Resolved anchor cell
    Dim AnchorKey               As String            'Stable same-target cache key

    Dim CandidateShape          As Excel.Shape       'Existing reusable icon candidate
    Dim ForeignIconPresent      As Boolean           'True when a same-named shape is not DatePicker-owned
    Dim OwnerMarker             As String            'Ownership marker written into DatePicker icons
    Dim IconLeft                As Double            'Icon left position
    Dim IconTop                 As Double            'Icon top position
    Dim HasReusableIcon         As Boolean           'True when an existing shape can be moved

    Dim MergeState              As Variant           'Target merge-state snapshot
    Dim CallbackMacroName       As String            'Workbook-qualified icon callback

    Dim HandlerStep             As String            'Current handler step for diagnostics
    Dim ErrorNumber             As Long              'Captured error number
    Dim ErrorDescription        As String            'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"
    'Load settings before reading the grid-icon feature flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' BUILD OWNERSHIP MARKER
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Build ownership marker"
    'Build the marker before any shape is inspected or normalized. Normalizing
    'an adopted icon writes this value, so an empty marker here would strip the
    'ownership evidence off an icon this component legitimately owns
        OwnerMarker = M_GridIcon_BuildOwnerMarker()

'------------------------------------------------------------------------------
' FEATURE GATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Check grid-icon feature flag"
    'Remove any stale icon and exit when the feature is disabled
        If Not gDP_ShowGridIcon Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' RESOLVE ANCHOR CELL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve anchor cell"
    'Use the supplied target cell when provided
        If Not TargetCell Is Nothing Then
            Set AnchorCell = TargetCell
    'Otherwise use ActiveCell safely
        Else
            On Error Resume Next
            Set AnchorCell = Excel.Application.ActiveCell
            Err.Clear
            On Error GoTo FailSafe
        End If
    'Remove stale icon and exit when no anchor cell is available
        If AnchorCell Is Nothing Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' NORMALIZE ANCHOR CELL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize anchor cell"
    'Use the first physical cell defensively
        Set AnchorCell = AnchorCell.Cells(1, 1)
    'Capture merge state safely
        MergeState = AnchorCell.MergeCells
    'Normalize merged targets to the top-left cell of the merge area
        If Not VBA.IsError(MergeState) Then
            If Not VBA.IsNull(MergeState) Then
                If VBA.CBool(MergeState) Then
                    Set AnchorCell = AnchorCell.MergeArea.Cells(1, 1)
                End If
            End If
        End If
    'Remove stale icon and exit when the target is not a single cell
        If AnchorCell.Cells.CountLarge <> 1 Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' RESOLVE TARGET WORKSHEET
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve target worksheet"
    'Store the target worksheet
        Set TargetSheet = AnchorCell.Worksheet
    'Remove stale icon and exit when no worksheet is available
        If TargetSheet Is Nothing Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' CALCULATE ICON POSITION
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Calculate icon position"
    'Position the icon to the right of the anchor cell
        IconLeft = AnchorCell.Left + AnchorCell.Width + ICON_GAP
    'Vertically center the icon against the anchor cell
        IconTop = AnchorCell.Top + ((AnchorCell.Height - ICON_SIZE) / 2)

'------------------------------------------------------------------------------
' SHORT-CIRCUIT SAME TARGET
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Check same-target grid icon"
    'Build the stable anchor key
        AnchorKey = M_GridIcon_BuildAnchorKey(AnchorCell)
    'Exit when the tracked icon is already visible at the requested target
        If M_GridIcon_IsSameVisibleTarget( _
            AnchorKey, _
            TargetSheet, _
            IconLeft, _
            IconTop, _
            ICON_SIZE) Then
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' RESOLVE CALLBACK MACRO
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve callback macro"
    'Build the workbook-qualified icon callback once
        CallbackMacroName = M_GetQualifiedMacroName("DP_Click")

'------------------------------------------------------------------------------
' RESOLVE TRACKED ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve tracked icon"
    'Suppress stale shape-reference failures
        On Error Resume Next
    'Start from the tracked icon reference
        Set CandidateShape = gDP_GridIconShape
    'Clear invalid tracked references
        If Err.Number <> 0 Then
            Err.Clear
            Set CandidateShape = Nothing
            Set gDP_GridIconShape = Nothing
        End If
    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' VALIDATE TRACKED ICON PARENT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate tracked icon parent"
    'Validate the tracked icon parent when a candidate exists
        If Not CandidateShape Is Nothing Then
            'Suppress stale parent-reference failures
                On Error Resume Next
            'Resolve whether the tracked icon belongs to the target worksheet
                HasReusableIcon = (CandidateShape.Parent Is TargetSheet)
            'Clear invalid tracked references when parent inspection fails
                If Err.Number <> 0 Then
                    Err.Clear
                    Set CandidateShape = Nothing
                    Set gDP_GridIconShape = Nothing
                    HasReusableIcon = False
            'Delete a tracked icon from another worksheet
                ElseIf Not HasReusableIcon Then
                    CandidateShape.Delete
                    Err.Clear
                    Set CandidateShape = Nothing
                    Set gDP_GridIconShape = Nothing
                End If
            'Restore fail-safe error handling
                On Error GoTo FailSafe
        End If

'------------------------------------------------------------------------------
' RESOLVE SAME-SHEET FALLBACK ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve same-sheet icon"
    'Try to reuse a same-named shape on the target worksheet, but only when it
    'is proven DatePicker-owned. A same-named foreign shape must not be moved,
    'resized, rebound, re-marked or deleted
        If CandidateShape Is Nothing Then
            'Suppress missing-shape lookup errors
                On Error Resume Next
            'Resolve an owned same-sheet icon, reporting a foreign collision
                Set CandidateShape = M_GridIcon_ResolveOwnedShape( _
                    TargetSheet, ForeignIconPresent)
            'Resolve whether a reusable same-sheet icon was found
                HasReusableIcon = Not (CandidateShape Is Nothing)
            'Clear lookup failures
                If Err.Number <> 0 Then
                    Err.Clear
                    Set CandidateShape = Nothing
                    HasReusableIcon = False
                End If
            'Restore fail-safe error handling
                On Error GoTo FailSafe
        End If

'------------------------------------------------------------------------------
' REFUSE A FOREIGN SAME-NAME COLLISION
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refuse foreign icon collision"
    'Leave a foreign same-named shape completely untouched and show no icon
        If ForeignIconPresent Then
            M_GridIcon_ReportForeignShape TargetSheet, PROC_NAME
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' MOVE EXISTING ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Move reusable icon"
    'Move and show the existing icon when it can be reused
        If HasReusableIcon Then
            With CandidateShape
                If .Left <> IconLeft Then .Left = IconLeft
                If .Top <> IconTop Then .Top = IconTop
                If .Width <> ICON_SIZE Then .Width = ICON_SIZE
                If .Height <> ICON_SIZE Then .Height = ICON_SIZE
                If .Placement <> xlMove Then .Placement = xlMove
                If .AlternativeText <> OwnerMarker Then .AlternativeText = OwnerMarker
                If .OnAction <> CallbackMacroName Then .OnAction = CallbackMacroName
                If .Visible <> msoTrue Then .Visible = msoTrue
                .ZOrder msoBringToFront
            End With
            'Store the reusable icon reference
                Set gDP_GridIconShape = CandidateShape
            'Forget any earlier refusal now that an owned icon is in place
                M_GridIcon_ClearRefusalKey
            'Remember the successfully positioned target
                M_GridIcon_RememberTarget AnchorKey, IconLeft, IconTop
            'Exit after moving the icon
                GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' CREATE ICON ONLY WHEN NEEDED
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Create icon"
    'Create the icon only when no reusable shape exists
        M_GridIcon_Create AnchorCell
    'Remember the target when cold creation succeeded
        If M_GridIcon_TrackedShapeIsLive() Then
            M_GridIcon_RememberTarget AnchorKey, IconLeft, IconTop
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release local references
        Set CandidateShape = Nothing
    'Release local references
        Set TargetSheet = Nothing
    'Release local references
        Set AnchorCell = Nothing
    'Clear any non-fatal pending error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Suppress fallback failures
        On Error Resume Next
    'Clear invalid tracked references
        Set gDP_GridIconShape = Nothing
    'Clear the cached last target after a grid-icon failure
        M_GridIcon_ClearLastTarget
    'Fall back to cold-path creation when an anchor cell is available
        If Not AnchorCell Is Nothing Then
            M_GridIcon_Create AnchorCell
        End If
    'Write diagnostics without interrupting worksheet interaction
        Debug.Print PROC_NAME & _
            " | Step=" & HandlerStep & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Release local references
        Set CandidateShape = Nothing
    'Release local references
        Set TargetSheet = Nothing
    'Release local references
        Set AnchorCell = Nothing
    'Clear any suppressed fallback error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Sub M_GridIcon_Hide()

'
'==============================================================================
'                           HIDE CURRENT GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Hides the currently tracked DatePicker in-grid icon without deleting it
'
' WHY THIS EXISTS
'   SelectionChange can fire very frequently
'
'   When the active cell is not eligible for the DatePicker, deleting the icon
'   forces the next eligible date-cell selection to recreate the worksheet shape
'
'   Hiding the existing icon keeps the drawing-layer shape available for fast
'   reuse when the user moves back to an eligible date cell
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Hides the tracked grid icon shape when available
'   Leaves the tracked shape reference intact when hiding succeeds
'   Clears the tracked shape reference only when the reference is stale
'   Does not scan worksheets or workbooks
'   Does not delete shapes
'
' ERROR POLICY
'   Best-effort high-frequency UI routine
'   Suppresses stale-shape and visibility errors
'   Does not raise outward
'
' DEPENDENCIES
'   gDP_GridIconShape
'
' NOTES
'   Use this routine from high-frequency selection-change paths when the current
'   target is not DatePicker-eligible
'
'   Use M_GridIcon_Remove or M_GridIcon_PurgeAll for hard cleanup boundaries
'
' UPDATED
'   2026-05-06
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress high-frequency UI cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' HIDE TRACKED SHAPE
'------------------------------------------------------------------------------
    'Hide the tracked icon when it still exists
        If M_GridIcon_TrackedShapeIsLive() Then
            gDP_GridIconShape.Visible = msoFalse
            'Clear stale tracked references only when hiding failed
                If Err.Number <> 0 Then
                    Err.Clear
                    Set gDP_GridIconShape = Nothing
                End If
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Clear any suppressed visibility error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub M_GridIcon_Create(Optional ByVal TargetCell As Excel.Range)

'
'==============================================================================
'                         CREATE GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Creates or recreates the in-grid DatePicker entry-point shape
'
' WHY THIS EXISTS
'   Users can invoke the DatePicker directly from the worksheet grid without
'   using the right-click menu, keyboard shortcut, or Ribbon
'
' INPUTS
'   TargetCell
'     Optional anchor range
'
'     If omitted, Excel.Application.ActiveCell is used when available
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures settings are loaded
'   Hides the current icon and exits when the grid-icon feature is disabled
'   Resolves and normalizes the anchor cell
'   Calculates the icon position
'   Resolves the embedded PNG file as the primary icon source
'   Creates a hidden temporary worksheet shape
'   Applies the embedded PNG picture icon when available
'   Falls back to the built-in font glyph only when the PNG path fails
'   Replaces the previous stable icon only after the new icon is prepared
'   Restores Application.ScreenUpdating deterministically
'
' ERROR POLICY
'   Best-effort UI routine
'
'   Suppresses creation failures, removes partial temporary shapes before stable
'   promotion, restores Application state, clears the tracked icon reference only
'   when stable replacement started but did not finish, and writes diagnostics to
'   the Immediate Window
'
' DEPENDENCIES
'   Excel Shapes object model
'   M_Settings_EnsureLoaded
'   M_GridIcon_Hide
'   M_GridIcon_EnsureEmbeddedIconFile
'   M_GetQualifiedMacroName
'   DP_GRID_ICON_NAME
'   DP_GRID_ICON_FONT_NAME
'   DP_GRID_ICON_CALENDAR_CODEPOINT
'   DP_GRID_ICON_FONT_SIZE
'   DP_GRID_ICON_FORE_COLOR
'   DP_THEME_COLOR_HEADER
'   gDP_ShowGridIcon
'   gDP_GridIconShape
'
' NOTES
'   The embedded PNG is the primary icon source
'
'   The Segoe MDL2 glyph is only a fallback when the embedded PNG cannot be
'   resolved, decoded, written, found, or applied to the shape
'
'   The routine deliberately does not delete the old icon at the beginning of
'   the valid create path because deleting the old icon before the new icon is
'   ready creates visible flicker
'
'   The old icon is replaced only after the new temporary icon has been created,
'   formatted, and filled
'
'   M_GridIcon_ShowOrMove remains the preferred high-frequency selection-change
'   path
'
'   This routine is the cold creation / recreation path
'
' UPDATED
'   2026-08-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_GridIcon_Create"
    Const ICON_SIZE             As Double = 24#
    Const ICON_GAP              As Double = 5#

    Dim AnchorCell              As Excel.Range      'Resolved anchor cell
    Dim TargetSheet             As Excel.Worksheet  'Worksheet receiving the icon
    Dim NewIconShape            As Excel.Shape      'New temporary icon shape
    Dim ExistingIconShape       As Excel.Shape      'Owned icon already holding the canonical name
    Dim PendingShape            As Excel.Shape      'Stale pending icon left by an interrupted run
    Dim ForeignIconPresent      As Boolean          'True when a same-named shape is not DatePicker-owned
    Dim OwnerMarker             As String           'Ownership marker written into DatePicker icons
    Dim IconLeft                As Double           'Icon left position
    Dim IconTop                 As Double           'Icon top position
    Dim IconFilePath            As String           'Resolved embedded icon file path
    Dim TempShapeName           As String           'Temporary icon shape name
    Dim UsedPictureIcon         As Boolean          'True when picture fill succeeds
    Dim StableReplaceStarted    As Boolean          'True once old stable icon replacement begins
    Dim StableIconPromoted      As Boolean          'True once the new icon is named and tracked
    Dim MergeState              As Variant          'Anchor merge-state snapshot
    Dim PreviousScreenUpdating  As Boolean          'Previous ScreenUpdating state
    Dim HasScreenState          As Boolean          'True after ScreenUpdating is captured
    Dim IconErrNumber           As Long             'Embedded icon resolution error number
    Dim IconErrDescription      As String           'Embedded icon resolution error description
    Dim PictureErrNumber        As Long             'Picture-fill error number
    Dim PictureErrDescription   As String           'Picture-fill error description
    Dim HandlerStep             As String           'Current handler step for diagnostics
    Dim ErrorNumber             As Long             'Captured error number
    Dim ErrorDescription        As String           'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"
    'Load settings before reading the grid-icon feature flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' FEATURE GATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Check grid-icon feature flag"
    'Hide any reusable icon and exit when the feature is disabled
        If Not gDP_ShowGridIcon Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' RESOLVE ANCHOR CELL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve anchor cell"
    'Use the supplied target cell when provided
        If Not TargetCell Is Nothing Then
            Set AnchorCell = TargetCell
        Else
            On Error Resume Next
            Set AnchorCell = Excel.Application.ActiveCell
            Err.Clear
            On Error GoTo FailSafe
        End If

    'Hide any reusable icon and exit when no anchor cell is available
        If AnchorCell Is Nothing Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' NORMALIZE ANCHOR CELL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize anchor cell"
    'Use the first physical cell defensively
        Set AnchorCell = AnchorCell.Cells(1, 1)
    'Capture merge state safely
        MergeState = AnchorCell.MergeCells
    'Normalize merged cells to the top-left cell
        If Not VBA.IsError(MergeState) Then
            If Not VBA.IsNull(MergeState) Then
                If VBA.CBool(MergeState) Then
                    Set AnchorCell = AnchorCell.MergeArea.Cells(1, 1)
                End If
            End If
        End If
    'Hide any reusable icon and exit when the target is not a single cell
        If AnchorCell.Cells.CountLarge <> 1 Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' RESOLVE TARGET WORKSHEET
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve target worksheet"
    'Store the target worksheet
        Set TargetSheet = AnchorCell.Worksheet
    'Hide any reusable icon and exit when no worksheet is available
        If TargetSheet Is Nothing Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' CALCULATE ICON POSITION
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Calculate icon position"
    'Position the icon to the right of the anchor cell
        IconLeft = AnchorCell.Left + AnchorCell.Width + ICON_GAP
    'Vertically center the icon against the anchor cell
        IconTop = AnchorCell.Top + ((AnchorCell.Height - ICON_SIZE) / 2)
    'Build the temporary shape name
        TempShapeName = DP_GRID_ICON_NAME & DP_GRID_ICON_PENDING_SUFFIX
    'Build the ownership marker before any shape is created, so an icon that
    'survives an interrupted create path is already recognizable to cleanup
        OwnerMarker = M_GridIcon_BuildOwnerMarker()

'------------------------------------------------------------------------------
' RESOLVE EMBEDDED ICON FILE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve embedded icon file"
    'Start with no resolved picture icon
        IconFilePath = vbNullString
    'Suppress embedded-icon materialization errors and allow fallback only if needed
        On Error Resume Next
    'Resolve or create the embedded temporary icon file
        IconFilePath = M_GridIcon_EnsureEmbeddedIconFile()
    'Write diagnostics when embedded-icon resolution failed
        If Err.Number <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Error=" & VBA.CStr(Err.Number) & _
                " | " & Err.Description
            Err.Clear
        End If
    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' CAPTURE APPLICATION STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture Application state"
    'Capture the current ScreenUpdating state
        PreviousScreenUpdating = Excel.Application.ScreenUpdating
    'Mark ScreenUpdating as captured
        HasScreenState = True
    'Suppress worksheet repaint during icon creation
        Excel.Application.ScreenUpdating = False

'------------------------------------------------------------------------------
' REMOVE STALE TEMPORARY SHAPE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Remove stale temporary shape"
    'Suppress stale temporary-shape cleanup errors
        On Error Resume Next
    'Resolve a stale temporary icon left by a previous interrupted run without
    'touching it. The pending name is internal, but this shape comes back from
    'Excel rather than from this transaction, so it is proven like any other
        Set PendingShape = Nothing
        Set PendingShape = TargetSheet.Shapes(TempShapeName)
        Err.Clear
    'Delete it only when DatePicker ownership is proven
        If Not PendingShape Is Nothing Then
            If M_GridIcon_IsOwnedShape(PendingShape) Then
                PendingShape.Delete
            End If
        End If
    'Release the resolved reference
        Set PendingShape = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' CREATE TEMPORARY ICON SHAPE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Create temporary icon shape"
    'Create the new icon with a temporary name
        Set NewIconShape = TargetSheet.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            IconLeft, _
            IconTop, _
            ICON_SIZE, _
            ICON_SIZE)

'------------------------------------------------------------------------------
' APPLY BASE SHAPE SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply base shape settings"
    'Apply base icon behavior and callback settings
        With NewIconShape
            .Name = TempShapeName
            .Placement = xlMove
            .AlternativeText = OwnerMarker
            .OnAction = M_GetQualifiedMacroName("DP_Click")
            .LockAspectRatio = msoTrue
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = DP_THEME_COLOR_HEADER
            .Line.Visible = msoFalse
            .Visible = msoFalse
        End With

'------------------------------------------------------------------------------
' APPLY EMBEDDED PICTURE ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply embedded picture icon"
    'Start from fallback mode
        UsedPictureIcon = False
    'Try to apply the embedded picture icon when a usable file path exists
        If VBA.LenB(IconFilePath) <> 0 Then
            If VBA.LenB(VBA.Dir$(IconFilePath, vbNormal)) <> 0 Then
                'Suppress picture-fill errors so glyph remains an emergency fallback
                    On Error Resume Next
                'Clear any fallback text before applying the picture
                    NewIconShape.TextFrame2.TextRange.Text = vbNullString
                'Apply the embedded PNG as the icon fill
                    NewIconShape.Fill.UserPicture IconFilePath
                'Store whether picture application succeeded
                    UsedPictureIcon = (Err.Number = 0)
                'Write diagnostics when picture application failed
                    If Err.Number <> 0 Then
                        Debug.Print PROC_NAME & _
                            " | Step=" & HandlerStep & _
                            " | IconFilePath=" & IconFilePath & _
                            " | Error=" & VBA.CStr(Err.Number) & _
                            " | " & Err.Description
                        Err.Clear
                    End If
                'Restore fail-safe error handling
                    On Error GoTo FailSafe
            Else
                'Log missing materialized icon file
                    Debug.Print PROC_NAME & _
                        " | Step=" & HandlerStep & _
                        " | Embedded icon file not found: " & IconFilePath
            End If
        Else
            'Log empty icon path so fallback is explainable
                Debug.Print PROC_NAME & _
                    " | Step=" & HandlerStep & _
                    " | Embedded icon path is empty"
        End If

'------------------------------------------------------------------------------
' APPLY FALLBACK GLYPH ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply fallback glyph icon"
    'Apply the built-in fallback glyph only when no picture icon was applied
        If Not UsedPictureIcon Then
            'Apply the base fallback shape appearance
                With NewIconShape
                    .Fill.Visible = msoTrue
                    .Fill.ForeColor.RGB = DP_THEME_COLOR_HEADER
                    .Line.Visible = msoFalse
                End With
            'Suppress glyph-formatting errors because the shape itself remains usable
                On Error Resume Next
            'Clear existing text
                NewIconShape.TextFrame2.TextRange.Text = vbNullString
            'Apply the calendar glyph
                NewIconShape.TextFrame2.TextRange.Text = VBA.ChrW$(DP_GRID_ICON_CALENDAR_CODEPOINT)
            'Apply the glyph font name
                NewIconShape.TextFrame2.TextRange.Font.Name = DP_GRID_ICON_FONT_NAME
            'Apply the glyph font size
                NewIconShape.TextFrame2.TextRange.Font.Size = DP_GRID_ICON_FONT_SIZE
            'Apply the glyph foreground color
                NewIconShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = DP_GRID_ICON_FORE_COLOR
            'Remove left text margin
                NewIconShape.TextFrame2.MarginLeft = 0
            'Remove right text margin
                NewIconShape.TextFrame2.MarginRight = 0
            'Remove top text margin
                NewIconShape.TextFrame2.MarginTop = 0
            'Remove bottom text margin
                NewIconShape.TextFrame2.MarginBottom = 0
            'Vertically center the glyph
                NewIconShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
            'Horizontally center the glyph
                NewIconShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            'Clear any suppressed glyph-formatting error
                Err.Clear
            'Restore fail-safe error handling
                On Error GoTo FailSafe
        End If

'------------------------------------------------------------------------------
' ATOMIC REPLACE OLD ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Replace old icon"
    'Resolve whether the canonical name is already taken, and by whom, before
    'anything is deleted or promoted
        Set ExistingIconShape = M_GridIcon_ResolveOwnedShape( _
            TargetSheet, ForeignIconPresent)
    'Abandon this creation when a foreign shape holds the canonical name. The
    'pending shape is discarded rather than promoted over the user's shape
        If ForeignIconPresent Then
            M_GridIcon_ReportForeignShape TargetSheet, PROC_NAME
            On Error Resume Next
            NewIconShape.Delete
            Err.Clear
            Set NewIconShape = Nothing
            Set ExistingIconShape = Nothing
            On Error GoTo FailSafe
            GoTo CleanExit
        End If
    'Mark that stable icon replacement has started
        StableReplaceStarted = True
    'Suppress old-icon cleanup errors
        On Error Resume Next
    'Delete the tracked old icon when available
        If M_GridIcon_TrackedShapeIsLive() Then
            gDP_GridIconShape.Delete
        End If
    'Delete the same-named icon only when it is proven DatePicker-owned
        If Not ExistingIconShape Is Nothing Then
            ExistingIconShape.Delete
        End If
    'Release the resolved reference
        Set ExistingIconShape = Nothing
    'Clear any suppressed old-icon cleanup error
        Err.Clear
    'Restore fail-safe error handling
        On Error GoTo FailSafe
    'Promote the temporary icon to the stable DatePicker icon name
        NewIconShape.Name = DP_GRID_ICON_NAME
    'Store the new icon reference
        Set gDP_GridIconShape = NewIconShape
    'Mark the new icon as safely promoted
        StableIconPromoted = True
    'Forget any earlier refusal now that an owned icon is in place
        M_GridIcon_ClearRefusalKey
    'Remember the successfully promoted icon target
        M_GridIcon_RememberTarget _
            M_GridIcon_BuildAnchorKey(AnchorCell), _
            IconLeft, _
            IconTop

'------------------------------------------------------------------------------
' SHOW FINAL ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Show final icon"
    'Suppress final visual-order failures because the icon has already been created
        On Error Resume Next
    'Show the fully formatted icon
        gDP_GridIconShape.Visible = msoTrue
    'Bring the icon to the front
        gDP_GridIconShape.ZOrder msoBringToFront
    'Clear any suppressed final visual-order error
        Err.Clear
    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Suppress state-restore failures
        On Error Resume Next
    'Restore ScreenUpdating when it was captured
        If HasScreenState Then
            Excel.Application.ScreenUpdating = PreviousScreenUpdating
        End If
    'Release local references
        Set NewIconShape = Nothing
    'Release local references
        Set TargetSheet = Nothing
    'Release local references
        Set AnchorCell = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Suppress cleanup failures
        On Error Resume Next
    'Delete the partially created temporary icon only before stable promotion
        If Not StableIconPromoted Then
            If Not NewIconShape Is Nothing Then NewIconShape.Delete
        End If
    'Clear the tracked icon reference only when replacement started but did not finish
        If StableReplaceStarted And Not StableIconPromoted Then
            Set gDP_GridIconShape = Nothing
        End If
    'Restore ScreenUpdating when it was captured
        If HasScreenState Then
            Excel.Application.ScreenUpdating = PreviousScreenUpdating
        End If

    'Write diagnostics without interrupting worksheet interaction
        Debug.Print PROC_NAME & _
            " | Step=" & HandlerStep & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Release local references
        Set NewIconShape = Nothing
    'Release local references
        Set TargetSheet = Nothing
    'Release local references
        Set AnchorCell = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Sub M_GridIcon_PreCreateHidden(Optional ByVal TargetCell As Excel.Range)

'
'==============================================================================
'                         PRE-CREATE HIDDEN GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Pre-creates the DatePicker in-grid icon and hides it for later fast reuse
'
' WHY THIS EXISTS
'   Creating worksheet shapes during SelectionChange can create visible latency
'
'   Pre-creating the icon during startup allows the high-frequency selection path
'   to move and show an existing shape instead of creating one on demand
'
' INPUTS
'   TargetCell
'     Optional initial anchor cell
'
'     If omitted, Excel.Application.ActiveCell is used when available
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures settings are loaded
'   Exits when the grid-icon feature is disabled
'   Creates the icon through the normal cold creation path when no tracked icon
'   exists
'   Hides the icon immediately after creation
'
' ERROR POLICY
'   Best-effort startup optimization
'   Does not raise outward
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_GridIcon_Create
'   gDP_ShowGridIcon
'   gDP_GridIconShape
'
' NOTES
'   This routine is an optimization only
'
'   M_GridIcon_ShowOrMove should still keep M_GridIcon_Create as a fallback
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorCell              As Excel.Range  'Initial icon anchor cell

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress startup optimization errors
        On Error Resume Next

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure settings are available before reading the grid-icon feature flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' FEATURE GATE
'------------------------------------------------------------------------------
    'Exit when the grid-icon feature is disabled
        If Not gDP_ShowGridIcon Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' EXIT IF ICON ALREADY EXISTS
'------------------------------------------------------------------------------
    'Exit when a tracked icon already exists
        If M_GridIcon_TrackedShapeIsLive() Then GoTo HideIcon

'------------------------------------------------------------------------------
' RESOLVE INITIAL ANCHOR
'------------------------------------------------------------------------------
    'Use the supplied target cell when provided
        If Not TargetCell Is Nothing Then
            Set AnchorCell = TargetCell
        Else
            Set AnchorCell = Application.ActiveCell
        End If

    'Exit when no anchor cell is available
        If AnchorCell Is Nothing Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' CREATE ICON
'------------------------------------------------------------------------------
    'Create the icon through the normal cold creation path
        M_GridIcon_Create AnchorCell

'------------------------------------------------------------------------------
' HIDE ICON
'------------------------------------------------------------------------------
HideIcon:
    'Hide the icon after creation so startup does not display it prematurely
        If M_GridIcon_TrackedShapeIsLive() Then
            gDP_GridIconShape.Visible = msoFalse
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Release local references
        Set AnchorCell = Nothing
    'Clear any suppressed startup optimization error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Function M_GridIcon_TrackedShapeIsLive() As Boolean

'
'==============================================================================
'                      TRACKED GRID ICON IS LIVE
'==============================================================================
' PURPOSE
'   Reports whether the tracked grid-icon reference still points at a shape that
'   exists, and clears it when it does not
'
' WHY THIS EXISTS
'   "Not gDP_GridIconShape Is Nothing" tests the variable, not the object. The
'   icon is an ordinary worksheet shape and can be destroyed without going
'   through any routine that maintains the reference:
'
'     the worksheet holding it is deleted
'     M_GridIcon_PurgeAll removes it by name from another workbook
'     the user deletes it
'
'   The variable then still holds a reference to an object that no longer exists,
'   and the next dereference raises. That produced a routine 424 on every
'   teardown, suppressed and logged, which is noise that hides real failures
'
' INPUTS
'   None
'
' RETURNS
'   True when the tracked shape exists and can be used
'
'   False when nothing is tracked, or when the reference is stale. A stale
'   reference is cleared before returning
'
' BEHAVIOR
'   Probes a cheap property of the tracked shape and reports whether the probe
'   succeeded
'
' ERROR POLICY
'   Best-effort. Never raises. A stale reference is an ordinary condition here,
'   not a failure to report
'
' DEPENDENCIES
'   gDP_GridIconShape
'
' NOTES
'   .Name is used as the probe because it is cheap and any surviving shape
'   answers it. Any property access would do; the point is that the object model
'   is asked rather than the variable
'
'   Callers that need to know a deletion failed should still capture Err after
'   their own call. This function only removes the stale-reference case from that
'   signal
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ProbeName       As String       'Probed shape name, discarded

'------------------------------------------------------------------------------
' PROBE THE TRACKED SHAPE
'------------------------------------------------------------------------------
    'Never let a liveness probe raise into a caller
        On Error Resume Next
    'Set safe default result
        M_GridIcon_TrackedShapeIsLive = False
    'Exit when nothing is tracked
        If gDP_GridIconShape Is Nothing Then
            Err.Clear
            Exit Function
        End If
    'Ask the object model whether the shape still answers
        Err.Clear
        ProbeName = gDP_GridIconShape.Name
    'Drop a reference to a shape that no longer exists
        If Err.Number <> 0 Then
            Err.Clear
            Set gDP_GridIconShape = Nothing
            Exit Function
        End If
    'Drop a reference to a shape this component cannot prove it owns. A live
    'reference with a matching name is not ownership, and downstream callers
    'delete through this result
        If Not M_GridIcon_IsOwnedShape(gDP_GridIconShape) Then
            Err.Clear
            Set gDP_GridIconShape = Nothing
            Exit Function
        End If
    'Report a usable tracked shape
        M_GridIcon_TrackedShapeIsLive = True
    'Clear any suppressed probe error
        Err.Clear

End Function

Public Sub M_GridIcon_Remove()

'
'==============================================================================
'                         REMOVE CURRENT GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the currently tracked DatePicker in-grid icon
'
' WHY THIS EXISTS
'   Selection-change and worksheet-change events can fire very frequently
'
'   The normal icon-removal path must therefore be lightweight and must not scan
'   all worksheets or workbooks
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Deletes the tracked DatePicker icon shape when available
'   Clears the tracked icon reference regardless of deletion result
'   Attempts to resolve the active worksheet
'   Attempts to resolve a same-named icon from the active worksheet
'   Deletes the active-sheet fallback icon when found
'   Logs non-trivial cleanup failures to the Immediate Window
'
' ERROR POLICY
'   Best-effort cleanup
'   Suppresses shape-deletion and ActiveSheet-resolution errors
'   Logs tracked-shape deletion failures
'   Logs ActiveSheet resolution failures
'   Logs active-sheet fallback-shape deletion failures
'   Does not log the expected case where no same-named active-sheet shape exists
'   Does not raise outward
'
' DEPENDENCIES
'   gDP_GridIconShape
'   DP_GRID_ICON_NAME
'   Excel Shapes object model
'
' NOTES
'   This routine is intended for high-frequency UI refresh paths
'
'   Missing active-sheet fallback icons are normal and are intentionally not
'   logged
'
'   Use M_GridIcon_PurgeAll for hard cleanup boundaries such as save, close,
'   print, teardown, or regression reset
'
' UPDATED
'   2026-08-27
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME                 As String = "M_GridIcon_Remove" 'Current procedure name

    Dim ActiveSheetObject           As Object       'Current active sheet object
    Dim ActiveWorksheet             As Worksheet    'Current active worksheet
    Dim FallbackShape               As Shape        'Same-named active-sheet fallback shape
    Dim FallbackShapeIsForeign      As Boolean      'True when the same-named shape is not DatePicker-owned

    Dim TrackedErrNumber            As Long         'Tracked-shape deletion error number
    Dim TrackedErrDescription       As String       'Tracked-shape deletion error description
    Dim ActiveSheetErrNumber        As Long         'ActiveSheet resolution error number
    Dim ActiveSheetErrDescription   As String       'ActiveSheet resolution error description
    Dim FallbackDeleteErrNumber     As Long         'Fallback-shape deletion error number
    Dim FallbackDeleteErrDescription As String      'Fallback-shape deletion error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' DELETE TRACKED SHAPE
'------------------------------------------------------------------------------
    'Clear any pending error before deleting the tracked shape
        Err.Clear
    'Delete the tracked grid icon shape when it still exists. A stale reference is
    'cleared by the liveness check and is not a deletion failure
        If M_GridIcon_TrackedShapeIsLive() Then
            gDP_GridIconShape.Delete
            TrackedErrNumber = Err.Number
            TrackedErrDescription = Err.Description
            Err.Clear
        End If
    'Clear the tracked shape reference regardless of deletion result
        Set gDP_GridIconShape = Nothing
    'Clear the cached last target
        M_GridIcon_ClearLastTarget

'------------------------------------------------------------------------------
' RESOLVE ACTIVE WORKSHEET
'------------------------------------------------------------------------------
    'Clear any pending error before resolving ActiveSheet
        Err.Clear
    'Capture the active sheet object safely
        Set ActiveSheetObject = Excel.Application.ActiveSheet
    'Capture ActiveSheet resolution failure if any
        ActiveSheetErrNumber = Err.Number
    'Capture ActiveSheet resolution failure description if any
        ActiveSheetErrDescription = Err.Description
    'Clear any suppressed ActiveSheet error
        Err.Clear
    'Use the active sheet only when it is a worksheet
        If Not ActiveSheetObject Is Nothing Then
            If TypeOf ActiveSheetObject Is Excel.Worksheet Then
                Set ActiveWorksheet = ActiveSheetObject
            End If
        End If

'------------------------------------------------------------------------------
' RESOLVE ACTIVE-SHEET FALLBACK SHAPE
'------------------------------------------------------------------------------
    'Resolve a same-named icon from the active worksheet, but only when it is
    'proven DatePicker-owned. A foreign same-named shape is left untouched
        If Not ActiveWorksheet Is Nothing Then
            Set FallbackShape = M_GridIcon_ResolveOwnedShape( _
                ActiveWorksheet, FallbackShapeIsForeign)
            Err.Clear
        End If

'------------------------------------------------------------------------------
' DELETE ACTIVE-SHEET FALLBACK SHAPE
'------------------------------------------------------------------------------
    'Delete the fallback shape only when it was resolved
        If Not FallbackShape Is Nothing Then
            'Clear any pending error before deleting the fallback shape
                Err.Clear
            'Delete the same-named active-sheet fallback icon
                FallbackShape.Delete
            'Capture fallback-shape deletion failure if any
                FallbackDeleteErrNumber = Err.Number
            'Capture fallback-shape deletion failure description if any
                FallbackDeleteErrDescription = Err.Description
            'Clear any suppressed fallback deletion error
                Err.Clear
        End If

'------------------------------------------------------------------------------
' DIAGNOSTICS
'------------------------------------------------------------------------------
    'Write tracked-shape diagnostics only when a tracked deletion failed
        If TrackedErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=DeleteTrackedShape" & _
                " | Error=" & VBA.CStr(TrackedErrNumber) & _
                " | " & TrackedErrDescription
        End If
    'Write ActiveSheet diagnostics only when ActiveSheet resolution failed
        If ActiveSheetErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=ResolveActiveSheet" & _
                " | Error=" & VBA.CStr(ActiveSheetErrNumber) & _
                " | " & ActiveSheetErrDescription
        End If
    'Write fallback-shape diagnostics only when deletion of a resolved shape failed
        If FallbackDeleteErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=DeleteActiveSheetFallbackShape" & _
                " | Error=" & VBA.CStr(FallbackDeleteErrNumber) & _
                " | " & FallbackDeleteErrDescription
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release local object references
        Set FallbackShape = Nothing
    'Release local object references
        Set ActiveWorksheet = Nothing
    'Release local object references
        Set ActiveSheetObject = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Sub M_GridIcon_PurgeAll_Impl()

'
'------------------------------------------------------------------------------
'                         PURGE ALL GRID ICONS
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_GridIcon_PurgeAll
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_GridIcon_PurgeAll
'
' RETURNS
'   As documented on M_DatePicker.M_GridIcon_PurgeAll
'
' BEHAVIOR
'   As documented on M_DatePicker.M_GridIcon_PurgeAll
'
' ERROR POLICY
'   As documented on M_DatePicker.M_GridIcon_PurgeAll. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_GridIcon_TryPurgeAll
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

    Dim ErrorNumber As Long
    Dim ErrorDescription As String

    If Not M_GridIcon_TryPurgeAll(ErrorNumber, ErrorDescription) Then
        Debug.Print "M_GridIcon_PurgeAll | Incomplete | Error=" & _
            VBA.CStr(ErrorNumber) & " | " & ErrorDescription
    End If

End Sub

Private Sub M_GridIcon_DeleteNamedShapeAcrossWorkbook( _
    ByVal TargetWorkbook As Workbook, _
    ByVal TargetShapeName As String)

'
'------------------------------------------------------------------------------
'                    DELETE NAMED SHAPE ACROSS WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes proven DatePicker-owned in-grid icon shapes from every worksheet in a
'   target workbook
'
' WHY THIS EXISTS
'   DatePicker in-grid icons are transient worksheet shapes
'
'   During workbook activation, add-in reload, VBA reset, or interrupted UI
'   flows, a stale icon may remain on a worksheet even when the tracked shape
'   reference has already been lost
'
'   This helper provides the workbook-level cleanup pass used by the broader
'   grid-icon purge routine
'
' INPUTS
'   TargetWorkbook
'     Workbook to scan
'
'   TargetShapeName
'     Name of the worksheet shape to delete
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Suppresses cleanup errors
'   Exits when no workbook is supplied
'   Exits when the target shape name is blank
'   Loops through every worksheet in the workbook
'   Resolves the same-named shape without modifying it
'   Deletes it only when DatePicker ownership is proven
'   Restores normal error handling before exit
'
' ERROR POLICY
'   Best-effort cleanup
'   Suppresses workbook, worksheet, protection, and shape-deletion errors
'   Does not raise custom errors
'
' DEPENDENCIES
'   Excel Workbook / Worksheet / Shape object model
'
' NOTES
'   The shape name selects candidates; it never proves ownership. Every
'   candidate is routed through M_GridIcon_IsOwnedShape before deletion, and a
'   shape whose ownership cannot be proven is left completely untouched
'
'   Before #53 this routine deleted by name alone, so an unrelated shape in an
'   unrelated workbook was destroyed silently under the suppression below
'
'   This routine intentionally does not trim TargetShapeName
'   Keep this helper aligned with M_GridIcon_PurgeAll
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurWorksheet           As Worksheet    'Worksheet being scanned
    Dim CurShape               As Excel.Shape  'Same-named shape found on the worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Exit when no workbook is supplied
        If TargetWorkbook Is Nothing Then GoTo ExitProcedure
    'Exit when the target shape name is blank
        If LenB(TargetShapeName) = 0 Then GoTo ExitProcedure


'------------------------------------------------------------------------------
' DELETE NAMED SHAPES
'------------------------------------------------------------------------------
    'Loop through all worksheets in the target workbook
        For Each CurWorksheet In TargetWorkbook.Worksheets
            'Resolve the same-named shape without touching it
                Set CurShape = Nothing
                Set CurShape = CurWorksheet.Shapes(TargetShapeName)
                Err.Clear
            'Delete it only when DatePicker ownership is proven. An unrelated
            'shape that merely shares the name is never deleted
                If Not CurShape Is Nothing Then
                    If M_GridIcon_IsOwnedShape(CurShape) Then
                        CurShape.Delete
                        Err.Clear
                    End If
                End If
        Next CurWorksheet

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Release local object references
        Set CurShape = Nothing
        Set CurWorksheet = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
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

Private Function M_GridIcon_TryDeleteOwnedNameAcrossWorkbook( _
    ByVal TargetWorkbook As Excel.Workbook, _
    ByVal TargetShapeName As String, _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'               DELETE OWNED SHAPES OF ONE NAME IN ONE WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes every proven DatePicker-owned shape of a given name from one workbook
'
' WHY THIS EXISTS
'   #53 established that the shape name selects candidates and never proves
'   ownership. Teardown needs the same rule, so every candidate goes through the
'   ownership predicate instead of being deleted by name
'
' INPUTS
'   TargetWorkbook
'     Workbook to clean; a missing workbook is a no-op success
'
'   TargetShapeName
'     Canonical or pending grid-icon name; a blank name is a no-op success
'
'   ErrorNumber, ErrorDescription
'     Receive the first deletion failure observed
'
' RETURNS
'   True when every owned shape of that name was deleted
'
' BEHAVIOR
'   Walks each worksheet backwards through the Shapes collection and, for each
'   name match, deletes only when M_GridIcon_IsOwnedShape proves ownership
'
' ERROR POLICY
'   Never raises outward. A deletion that fails is recorded and the walk
'   continues
'
' DEPENDENCIES
'   M_GridIcon_IsOwnedShape
'
' NOTES
'   The index walk runs backwards because deleting a shape renumbers the
'   collection. A forward walk would skip the shape after each deletion
'
'   A shape whose ownership cannot be proven is left completely untouched, which
'   is the #53 contract: an unrelated user shape sharing the name survives
'   teardown
'
'   Moved from M_DP_Core by #57, where it was named
'   M_Lifecycle_TryDeleteOwnedShapeNameAcrossWorkbook, so the module that
'   owns this integration's state also verifies its teardown
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

    Dim CurWorksheet As Excel.Worksheet
    Dim CurShape As Excel.Shape
    Dim ShapeIndex As Long
    Dim ShapeName As String
    Dim DeleteErrNumber As Long
    Dim DeleteErrDescription As String

    On Error GoTo Failed
    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString
    M_GridIcon_TryDeleteOwnedNameAcrossWorkbook = True

    If TargetWorkbook Is Nothing Then Exit Function
    If VBA.LenB(TargetShapeName) = 0 Then Exit Function

    For Each CurWorksheet In TargetWorkbook.Worksheets
        For ShapeIndex = CurWorksheet.Shapes.Count To 1 Step -1
            Set CurShape = CurWorksheet.Shapes(ShapeIndex)
            ShapeName = VBA.CStr(CurShape.Name)
            If VBA.StrComp(ShapeName, TargetShapeName, vbBinaryCompare) = 0 Then
                If M_GridIcon_IsOwnedShape(CurShape) Then
                    On Error Resume Next
                    Err.Clear
                    CurShape.Delete
                    DeleteErrNumber = Err.Number
                    DeleteErrDescription = Err.Description
                    Err.Clear
                    On Error GoTo Failed
                    If DeleteErrNumber <> 0 Then
                        If M_GridIcon_TryDeleteOwnedNameAcrossWorkbook Then
                            ErrorNumber = DeleteErrNumber
                            ErrorDescription = DeleteErrDescription
                        End If
                        M_GridIcon_TryDeleteOwnedNameAcrossWorkbook = False
                    End If
                End If
            End If
            Set CurShape = Nothing
        Next ShapeIndex
    Next CurWorksheet

    Set CurShape = Nothing
    Set CurWorksheet = Nothing
    Exit Function

Failed:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    Set CurShape = Nothing
    Set CurWorksheet = Nothing
    Err.Clear
    M_GridIcon_TryDeleteOwnedNameAcrossWorkbook = False

End Function

Public Function M_GridIcon_TryPurgeAll( _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'               PURGE OWNED GRID ICONS FROM EVERY OPEN WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Removes every owned grid icon, canonical and pending, from all open workbooks
'
' WHY THIS EXISTS
'   Teardown cannot claim a clean shutdown while an icon it owns is still on a
'   worksheet, and the icon follows the selection, so the purge has to span
'   workbooks rather than only the host
'
' INPUTS
'   ErrorNumber, ErrorDescription
'     Receive the first failure observed across all workbooks
'
' RETURNS
'   True only when every workbook was purged without a failure
'
' BEHAVIOR
'   Walks every open workbook and deletes both grid-icon names through the
'   ownership-gated helper, keeping the first failure and continuing
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_GridIcon_TryDeleteOwnedNameAcrossWorkbook
'   DP_GRID_ICON_NAME
'   DP_GRID_ICON_PENDING_SUFFIX
'
' NOTES
'   Continuing past the first failure is deliberate. A workbook that refuses
'   deletion, such as a protected sheet, must not stop the rest being cleaned
'
'   The pending name is purged as well, so an icon left half-created by an
'   interrupted create path is not stranded
'
'   Moved from M_DP_Core by #57, where it was named
'   M_Lifecycle_TryPurgeGridIcons, so the module that owns this
'   integration's state also verifies its teardown
'
'   Public only so the lifecycle cleanup in M_DP_Core can call it
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

    Dim CurWorkbook As Excel.Workbook
    Dim LocalErrNumber As Long
    Dim LocalErrDescription As String
    Dim OverallSuccess As Boolean

    On Error GoTo Failed
    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString
    OverallSuccess = True

    For Each CurWorkbook In Excel.Application.Workbooks
        LocalErrNumber = 0
        LocalErrDescription = VBA.vbNullString
        If Not M_GridIcon_TryDeleteOwnedNameAcrossWorkbook( _
            CurWorkbook, DP_GRID_ICON_NAME, LocalErrNumber, LocalErrDescription) Then
            If OverallSuccess Then
                ErrorNumber = LocalErrNumber
                ErrorDescription = LocalErrDescription
            End If
            OverallSuccess = False
        End If

        LocalErrNumber = 0
        LocalErrDescription = VBA.vbNullString
        If Not M_GridIcon_TryDeleteOwnedNameAcrossWorkbook( _
            CurWorkbook, DP_GRID_ICON_NAME & DP_GRID_ICON_PENDING_SUFFIX, _
            LocalErrNumber, LocalErrDescription) Then
            If OverallSuccess Then
                ErrorNumber = LocalErrNumber
                ErrorDescription = LocalErrDescription
            End If
            OverallSuccess = False
        End If
    Next CurWorkbook

    If M_GridIcon_TrackedShapeIsLive() Then
        On Error Resume Next
        Err.Clear
        gDP_GridIconShape.Delete
        LocalErrNumber = Err.Number
        LocalErrDescription = Err.Description
        Err.Clear
        On Error GoTo Failed
        If LocalErrNumber <> 0 Then
            If OverallSuccess Then
                ErrorNumber = LocalErrNumber
                ErrorDescription = LocalErrDescription
            End If
            OverallSuccess = False
        End If
    End If

    If Not M_GridIcon_TrackedShapeIsLive() Then Set gDP_GridIconShape = Nothing
    M_GridIcon_ClearRefusalKey
    M_GridIcon_ClearLastTarget

    If OverallSuccess Then
        If M_GridIcon_TrackedShapeIsLive() Then
            OverallSuccess = False
            ErrorNumber = vbObjectError + 2714
            ErrorDescription = "A tracked DatePicker grid icon remained live after purge"
        End If
    End If

    M_GridIcon_TryPurgeAll = OverallSuccess
    Set CurWorkbook = Nothing
    Exit Function

Failed:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    Set CurWorkbook = Nothing
    Err.Clear
    M_GridIcon_TryPurgeAll = False

End Function
