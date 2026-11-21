Attribute VB_Name = "M_DP_GridIconAsset"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_DP_GridIconAsset
'------------------------------------------------------------------------------
' PURPOSE
'   Holds the embedded grid-icon image and writes it to a temporary file
'   that Excel can load
'
' WHY THIS EXISTS
'   The image is a large Base64 payload with its own failure modes: a
'   missing temp folder, a short write, a corrupt file. Keeping it apart
'   from placement keeps the shape code readable and testable (#57)
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides M_GridIcon_EnsureEmbeddedIconFile, which reuses a usable
'   temporary file or decodes the payload and writes a new one
'
' ERROR POLICY
'   Each procedure documents its own policy
'
' DEPENDENCIES
'   MSXML2.DOMDocument.6.0 through late binding
'   VBA binary file I/O and the Windows temporary folder
'
' NOTES
'   Only M_GridIcon_EnsureEmbeddedIconFile is Public; M_DP_GridIcon calls it
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
    Private Const DP_GRID_ICON_TEMP_FILE_NAME As String = "VBA_DatePicker_GridIcon_v1_64.png"   'Embedded icon temp-file name
    Private Const DP_GRID_ICON_TEMP_MIN_BYTES As Long = 500                                     'Minimum valid embedded icon size


'
'------------------------------------------------------------------------------
'
'                             EMBEDDED ICON ASSET
'
'------------------------------------------------------------------------------
'

Public Function M_GridIcon_EnsureEmbeddedIconFile() As String

'
'------------------------------------------------------------------------------
'                       ENSURE EMBEDDED GRID ICON FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the path of the embedded DatePicker grid-icon PNG stored in the
'   current user's temporary directory
'
' WHY THIS EXISTS
'   Shape.Fill.UserPicture requires a file path. The project should not
'   distribute a separate icon file, so the icon is embedded as Base64 text in
'   VBA code and materialized to the user's temp folder on demand
'
' INPUTS
'   None
'
' RETURNS
'   Full path of the temporary PNG file
'
' BEHAVIOR
'   Resolves the user's temp directory, builds a stable icon file path, reuses
'   the file if it already exists and appears valid, recreates the file from the
'   embedded Base64 payload when missing or invalid, validates the written file,
'   and returns the final valid file path
'
' ERROR POLICY
'   Raises a descriptive runtime error if the temp path cannot be resolved, the
'   embedded payload cannot be decoded, the icon file cannot be written, or the
'   written icon file fails validation
'
' DEPENDENCIES
'   M_GridIcon_GetTempFolder
'   M_GridIcon_CombinePath
'   M_GridIcon_IconFileIsUsable
'   M_GridIcon_EmbeddedIconBytes
'   M_GridIcon_WriteBytesToFile
'   DP_GRID_ICON_TEMP_FILE_NAME
'
' NOTES
'   The temp file is intentionally not deleted during normal cleanup
'
'   The routine returns an existing valid temp file without decoding the embedded
'   payload again
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_GridIcon_EnsureEmbeddedIconFile"

    Dim TempFolder              As String       'Resolved temp folder
    Dim IconPath                As String       'Resolved icon file path
    Dim IconBytes()             As Byte         'Decoded embedded PNG bytes
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
' RESOLVE TEMP PATH
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve temp folder"
    'Resolve the user's temp folder
        TempFolder = M_GridIcon_GetTempFolder()
    'Reject an empty temp folder
        If VBA.LenB(TempFolder) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "Temporary folder could not be resolved."
        End If

'------------------------------------------------------------------------------
' BUILD ICON PATH
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Build icon path"
    'Build the full icon path
        IconPath = M_GridIcon_CombinePath(TempFolder, DP_GRID_ICON_TEMP_FILE_NAME)

'------------------------------------------------------------------------------
' REUSE EXISTING FILE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate existing icon file"
    'Return the existing temp file when it appears usable
        If M_GridIcon_IconFileIsUsable(IconPath) Then
            M_GridIcon_EnsureEmbeddedIconFile = IconPath
            Exit Function
        End If

'------------------------------------------------------------------------------
' DECODE EMBEDDED PAYLOAD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Decode embedded icon payload"
    'Decode the embedded PNG payload
        IconBytes = M_GridIcon_EmbeddedIconBytes()

'------------------------------------------------------------------------------
' WRITE TEMP FILE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Write embedded icon file"
    'Write the decoded bytes to the temp file
        M_GridIcon_WriteBytesToFile IconPath, IconBytes

'------------------------------------------------------------------------------
' VALIDATE WRITTEN FILE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate written icon file"
    'Reject the file if it still does not appear usable
        If Not M_GridIcon_IconFileIsUsable(IconPath) Then
            Err.Raise vbObjectError + 514, PROC_NAME, "Embedded icon file was written but failed validation."
        End If

'------------------------------------------------------------------------------
' RETURN PATH
'------------------------------------------------------------------------------
    'Return the valid icon path
        M_GridIcon_EnsureEmbeddedIconFile = IconPath

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release decoded byte-array storage
        Erase IconBytes
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Release decoded byte-array storage
        Erase IconBytes
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            ErrorDescription

End Function

Private Function M_GridIcon_GetTempFolder() As String

'
'------------------------------------------------------------------------------
'                           GET TEMP FOLDER
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the current user's temporary folder
'
' WHY THIS EXISTS
'   The embedded grid icon needs a writable per-user location because
'   Shape.Fill.UserPicture requires a real file path
'
' INPUTS
'   None
'
' RETURNS
'   Normalized temp folder path without a trailing path separator
'
' BEHAVIOR
'   Reads TEMP first, falls back to TMP when TEMP is unavailable, trims the
'   resolved value, removes trailing path separators, and returns the normalized
'   folder path
'
' ERROR POLICY
'   Safe default. Returns an empty string on failure and writes diagnostics to
'   the Immediate Window only when an unexpected runtime error occurs
'
' DEPENDENCIES
'   VBA.Environ$
'
' NOTES
'   The Windows temp folder should already exist
'
'   This routine only resolves the path. It does not validate write permission
'   and does not create folders
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_GridIcon_GetTempFolder"

    Dim TempFolder              As String       'Resolved temporary folder path
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the safe default
        M_GridIcon_GetTempFolder = vbNullString
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE TEMP FOLDER
'------------------------------------------------------------------------------
    'Read TEMP first
        TempFolder = VBA.Trim$(VBA.Environ$("TEMP"))
    'Read TMP when TEMP is unavailable
        If VBA.LenB(TempFolder) = 0 Then
            TempFolder = VBA.Trim$(VBA.Environ$("TMP"))
        End If

'------------------------------------------------------------------------------
' NORMALIZE PATH
'------------------------------------------------------------------------------
    'Remove trailing path separators when present
        Do While VBA.LenB(TempFolder) > 0
            If VBA.Right$(TempFolder, 1) = "\" Or VBA.Right$(TempFolder, 1) = "/" Then
                TempFolder = VBA.Left$(TempFolder, VBA.Len(TempFolder) - 1)
            Else
                Exit Do
            End If
        Loop

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the normalized temp folder path
        M_GridIcon_GetTempFolder = TempFolder

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
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
    'Return the safe default
        M_GridIcon_GetTempFolder = vbNullString
    'Write diagnostics only when an unexpected error occurred
        If ErrorNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Error=" & VBA.CStr(ErrorNumber) & _
                " | " & ErrorDescription
        End If
    'Clear the suppressed error
        Err.Clear

End Function

Private Function M_GridIcon_CombinePath( _
    ByVal FolderPath As String, _
    ByVal FileName As String) As String

'
'------------------------------------------------------------------------------
'                           COMBINE PATH
'------------------------------------------------------------------------------
' PURPOSE
'   Combines a folder path and file name using a Windows path separator
'
' WHY THIS EXISTS
'   Repeated temp-file path construction should remain deterministic, readable,
'   and isolated in one helper
'
' INPUTS
'   FolderPath
'     Folder path
'
'   FileName
'     File name
'
' RETURNS
'   Combined file path
'
' BEHAVIOR
'   If FolderPath already ends with a backslash, appends FileName directly
'   If FolderPath does not end with a backslash, inserts one backslash separator
'   Does not trim, validate, normalize, expand, or inspect either input
'
' ERROR POLICY
'   Does not raise custom errors
'   Lets native VBA string behavior apply
'
' DEPENDENCIES
'   None
'
' NOTES
'   This helper assumes a Windows Excel environment
'   Empty FolderPath behavior is intentionally preserved from the original code
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const C_PATH_SEPARATOR As String = "\"        'Windows path separator

'------------------------------------------------------------------------------
' RETURN COMBINED PATH
'------------------------------------------------------------------------------
    'Return the direct concatenation when the folder already ends with a separator
        If Right$(FolderPath, Len(C_PATH_SEPARATOR)) = C_PATH_SEPARATOR Then
            M_GridIcon_CombinePath = FolderPath & FileName
        Else
            M_GridIcon_CombinePath = FolderPath & C_PATH_SEPARATOR & FileName
        End If

End Function

Private Function M_GridIcon_IconFileIsUsable(ByVal IconPath As String) As Boolean

'
'------------------------------------------------------------------------------
'                         ICON FILE IS USABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a candidate embedded grid-icon file appears usable
'
' WHY THIS EXISTS
'   A previous temp-file write may have failed, produced a missing file, or
'   produced a truncated file
'
'   The DatePicker should recreate the embedded grid-icon file when the existing
'   candidate file is missing, too small, unreadable, or not recognizable as PNG
'
' INPUTS
'   IconPath
'     Candidate PNG path
'
' RETURNS
'   True if the file exists, has a plausible size, and starts with the PNG
'   signature bytes
'
'   False otherwise
'
' BEHAVIOR
'   Rejects blank paths after Trim$
'   Rejects missing files
'   Rejects files smaller than DP_GRID_ICON_TEMP_MIN_BYTES
'   Opens the candidate file in binary read mode
'   Reads the first eight bytes
'   Accepts the file only when the standard PNG signature matches
'
' ERROR POLICY
'   Safe-default predicate
'   Returns False on any failure
'   Suppresses cleanup errors while attempting to close the file handle
'
' DEPENDENCIES
'   DP_GRID_ICON_TEMP_MIN_BYTES
'   Dir$
'   FileLen
'   FreeFile
'
' NOTES
'   This is a lightweight validation, not a full PNG parser
'   The routine intentionally does not validate the full PNG structure
'   The routine intentionally preserves the original path handling behavior
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FileNumber             As Integer      'Free file number
    Dim Signature(1 To 8)       As Byte         'PNG signature bytes

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable safe-default error handling
        On Error GoTo ErrorHandler
    'Default to unusable
        M_GridIcon_IconFileIsUsable = False

'------------------------------------------------------------------------------
' VALIDATE PATH
'------------------------------------------------------------------------------
    'Exit when the path is blank
        If LenB(Trim$(IconPath)) = 0 Then Exit Function
    'Exit when the file does not exist
        If LenB(Dir$(IconPath, vbNormal)) = 0 Then Exit Function
    'Exit when the file is implausibly small
        If FileLen(IconPath) < DP_GRID_ICON_TEMP_MIN_BYTES Then Exit Function

'------------------------------------------------------------------------------
' READ SIGNATURE
'------------------------------------------------------------------------------
    'Get a free file number
        FileNumber = FreeFile
    'Open the file for binary reading
        Open IconPath For Binary Access Read As #FileNumber
    'Read the first eight bytes
        Get #FileNumber, 1, Signature
    'Close the file
        Close #FileNumber

'------------------------------------------------------------------------------
' VALIDATE PNG SIGNATURE
'------------------------------------------------------------------------------
    'Return True only when the standard PNG signature matches
        M_GridIcon_IconFileIsUsable = _
            (Signature(1) = &H89 And _
             Signature(2) = &H50 And _
             Signature(3) = &H4E And _
             Signature(4) = &H47 And _
             Signature(5) = &HD And _
             Signature(6) = &HA And _
             Signature(7) = &H1A And _
             Signature(8) = &HA)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Suppress cleanup errors
        On Error Resume Next
    'Close the file number when one has been assigned
        If FileNumber <> 0 Then Close #FileNumber
    'Return safe default
        M_GridIcon_IconFileIsUsable = False
    'Restore normal error handling
        On Error GoTo 0

End Function

Private Sub M_GridIcon_WriteBytesToFile( _
    ByVal IconPath As String, _
    ByRef IconBytes() As Byte)

'
'------------------------------------------------------------------------------
'                         WRITE BYTES TO FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes decoded embedded icon bytes to a binary file
'
' WHY THIS EXISTS
'   The embedded PNG is stored as Base64 text in code, but Excel requires a
'   physical file path for Shape.Fill.UserPicture
'
'   This helper isolates the disk-write step so the grid-icon creation logic
'   remains focused on path resolution, icon validation, and shape rendering
'
' INPUTS
'   IconPath
'     Destination file path
'
'   IconBytes
'     Decoded PNG byte array
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Rejects blank destination paths after Trim$
'   Rejects empty byte payloads
'   Deletes an existing target file before writing
'   Opens the target file in binary write mode
'   Writes the full byte array starting at byte position 1
'   Closes the file handle after writing
'
' ERROR POLICY
'   Raises a descriptive runtime error if validation fails or the file cannot
'   be written
'
'   Attempts to close the file handle before re-raising the original error
'
' DEPENDENCIES
'   M_GridIcon_ByteArrayLength
'   Dir$
'   Kill
'   FreeFile
'
' NOTES
'   Existing temp files are overwritten
'   Parent-folder creation is intentionally not handled here
'   The caller is responsible for passing a valid writable destination folder
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_GridIcon_WriteBytesToFile"

    Dim FileNumber             As Integer      'Free file number
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty destination path
        If LenB(Trim$(IconPath)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "IconPath cannot be empty."
        End If
    'Reject an empty byte payload
        If M_GridIcon_ByteArrayLength(IconBytes) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "IconBytes cannot be empty."
        End If

'------------------------------------------------------------------------------
' REPLACE EXISTING FILE
'------------------------------------------------------------------------------
    'Delete any existing file before writing
        If LenB(Dir$(IconPath, vbNormal)) <> 0 Then Kill IconPath

'------------------------------------------------------------------------------
' WRITE FILE
'------------------------------------------------------------------------------
    'Get a free file number
        FileNumber = FreeFile
    'Open the target file for binary writing
        Open IconPath For Binary Access Write As #FileNumber
    'Write the byte array to disk
        Put #FileNumber, 1, IconBytes
    'Close the file
        Close #FileNumber

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error before cleanup
        ErrorNumber = Err.Number
    'Capture the original error description before cleanup
        ErrorDescription = Err.Description
    'Suppress cleanup errors
        On Error Resume Next
    'Close the file number when one has been assigned
        If FileNumber <> 0 Then Close #FileNumber
    'Restore normal error handling
        On Error GoTo 0
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Sub

Private Function M_GridIcon_ByteArrayLength(ByRef SourceBytes() As Byte) As Long

'
'------------------------------------------------------------------------------
'                         BYTE ARRAY LENGTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of bytes in a byte array
'
' WHY THIS EXISTS
'   Dynamic byte arrays may be uninitialized
'
'   A defensive length helper keeps binary-write validation safe and avoids
'   repeating fragile LBound / UBound checks around the grid-icon file writer
'
' INPUTS
'   SourceBytes
'     Byte array to inspect
'
' RETURNS
'   Number of bytes in the array
'
'   Zero when the array bounds cannot be read
'
' BEHAVIOR
'   Uses LBound and UBound to calculate the byte count
'   Returns zero when SourceBytes is uninitialized
'   Returns zero when array-bound inspection fails for any reason
'
' ERROR POLICY
'   Safe-default predicate
'   Returns zero on failure
'   Does not raise custom errors
'
' DEPENDENCIES
'   LBound
'   UBound
'
' NOTES
'   This helper assumes a one-dimensional byte array
'   The routine intentionally performs no content validation
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return zero when array bounds cannot be read
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RETURN LENGTH
'------------------------------------------------------------------------------
    'Return the byte count from the declared array bounds
        M_GridIcon_ByteArrayLength = UBound(SourceBytes) - LBound(SourceBytes) + 1

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Return the safe default
        M_GridIcon_ByteArrayLength = 0

End Function

Private Function M_GridIcon_Base64LooksValid(ByVal EncodedText As String) As Boolean

'
'==============================================================================
'                         BASE64 LOOKS VALID
'------------------------------------------------------------------------------
' PURPOSE
'   Performs lightweight validation of an embedded Base64 payload
'
' WHY THIS EXISTS
'   MSXML raises a generic parsing error when an invalid Base64 string is assigned
'   to a bin.base64 node
'
'   This helper catches obvious payload corruption before the decode step so the
'   caller can fail with a clearer diagnostic
'
' INPUTS
'   EncodedText
'     Base64 text to validate
'
' RETURNS
'   True when the text looks like a valid Base64 payload
'
'   False when the text is blank, has invalid length, contains unsupported
'   characters, or contains padding in an invalid position
'
' BEHAVIOR
'   Removes common whitespace characters
'   Checks that the remaining length is a multiple of four
'   Allows only A-Z, a-z, 0-9, plus, slash, and equals
'   Allows equals padding only in the final two characters
'
' ERROR POLICY
'   Safe-default predicate
'
'   Returns False on any unexpected validation failure
'
' DEPENDENCIES
'   None
'
' NOTES
'   This is not a full decoder
'
'   It is intentionally a fast pre-flight check before MSXML decoding
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CleanText       As String       'Whitespace-normalized Base64 text
    Dim TextLength      As Long         'Normalized text length
    Dim CharIndex       As Long         'Character loop index
    Dim CurrentChar     As String       'Current character being inspected
    Dim PaddingPosition As Long         'First equals padding position

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default result
        M_GridIcon_Base64LooksValid = False

    'Enable safe-default validation
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Trim the supplied payload
        CleanText = VBA.Trim$(EncodedText)

    'Remove carriage returns
        CleanText = VBA.Replace(CleanText, vbCr, vbNullString)

    'Remove line feeds
        CleanText = VBA.Replace(CleanText, vbLf, vbNullString)

    'Remove tab characters
        CleanText = VBA.Replace(CleanText, vbTab, vbNullString)

    'Remove spaces
        CleanText = VBA.Replace(CleanText, " ", vbNullString)

'------------------------------------------------------------------------------
' VALIDATE LENGTH
'------------------------------------------------------------------------------
    'Capture normalized text length
        TextLength = VBA.Len(CleanText)

    'Reject an empty payload
        If TextLength = 0 Then Exit Function

    'Reject payloads whose length is not divisible by four
        If TextLength Mod 4 <> 0 Then Exit Function

'------------------------------------------------------------------------------
' VALIDATE CHARACTERS
'------------------------------------------------------------------------------
    'Loop through each Base64 character
        For CharIndex = 1 To TextLength
            'Read the current character
                CurrentChar = VBA.Mid$(CleanText, CharIndex, 1)
            'Validate the current character
                Select Case CurrentChar
                    Case "A" To "Z", "a" To "z", "0" To "9", "+", "/"
                        'Valid non-padding Base64 character
                    Case "="
                        'Store the first padding position
                            If PaddingPosition = 0 Then PaddingPosition = CharIndex
                        'Reject padding before the final two characters
                            If CharIndex < TextLength - 1 Then Exit Function
                    Case Else
                        'Reject unsupported characters
                            Exit Function
                End Select
        Next CharIndex

'------------------------------------------------------------------------------
' VALIDATE PADDING ORDER
'------------------------------------------------------------------------------
    'Reject non-padding characters after padding begins
        If PaddingPosition > 0 Then
            For CharIndex = PaddingPosition To TextLength
                If VBA.Mid$(CleanText, CharIndex, 1) <> "=" Then Exit Function
            Next CharIndex
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Return successful validation
        M_GridIcon_Base64LooksValid = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful validation
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Return safe default on unexpected validation failure
        M_GridIcon_Base64LooksValid = False

End Function

Private Function M_GridIcon_EmbeddedIconBytes() As Byte()

'
'------------------------------------------------------------------------------
'                         EMBEDDED ICON BYTES
'------------------------------------------------------------------------------
' PURPOSE
'   Decodes the embedded Base64 PNG payload into a byte array
'
' WHY THIS EXISTS
'   The grid icon should use the embedded PNG as the primary visual asset
'
'   The worksheet glyph remains only a fallback when the PNG cannot be decoded,
'   written, or applied to the shape
'
' INPUTS
'   None
'
' RETURNS
'   Decoded PNG bytes
'
' BEHAVIOR
'   Loads the embedded Base64 payload
'   Normalizes the Base64 text by removing whitespace and copy/paste artifacts
'   Validates the Base64 character set and padding position before decoding
'   Decodes the payload through a late-bound MSXML Base64 node
'
' ERROR POLICY
'   Raises a compact descriptive runtime error when the embedded payload is
'   blank, structurally invalid, or cannot be decoded
'
'   Does not raise the full MSXML error text because that can contain the entire
'   Base64 payload and flood the Immediate Window or result sheet
'
' DEPENDENCIES
'   M_GridIcon_EmbeddedIconBase64
'   M_GridIcon_NormalizeBase64
'   M_GridIcon_Base64LooksValid
'   MSXML2.DOMDocument.6.0
'
' NOTES
'   This routine keeps the embedded PNG as the primary grid-icon source
'
'   If this routine fails, M_GridIcon_Create should fall back to the glyph path
'
' UPDATED
'   2026-05-08
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_GridIcon_EmbeddedIconBytes"

    Dim XmlDoc                 As Object       'MSXML document
    Dim XmlNode                As Object       'MSXML Base64 node
    Dim EncodedText            As String       'Normalized Base64 encoded PNG
    Dim PayloadLength          As Long         'Normalized Base64 payload length
    Dim PayloadRemainder       As Long         'Payload length remainder modulo 4
    Dim HandlerStep            As String       'Current handler step for diagnostics
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' LOAD AND NORMALIZE PAYLOAD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Load embedded Base64 payload"
    'Load and normalize the embedded Base64 payload
        EncodedText = M_GridIcon_NormalizeBase64(M_GridIcon_EmbeddedIconBase64())
    'Capture normalized payload length
        PayloadLength = VBA.Len(EncodedText)
    'Capture normalized payload modulo 4
        PayloadRemainder = PayloadLength Mod 4

'------------------------------------------------------------------------------
' VALIDATE BASE64 PAYLOAD
'------------------------------------------------------------------------------
    'Reject an invalid embedded payload before MSXML decoding
        If Not M_GridIcon_Base64LooksValid(EncodedText) Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Embedded icon payload is not valid Base64."
        End If

'------------------------------------------------------------------------------
' CREATE DECODER
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Create MSXML decoder"
    'Create the MSXML document
        Set XmlDoc = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    'Create a temporary Base64 node
        Set XmlNode = XmlDoc.createElement("Base64Data")

'------------------------------------------------------------------------------
' DECODE BASE64
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Decode Base64 payload"
    'Configure the node as Base64 binary data
        XmlNode.DataType = "bin.base64"
    'Assign the normalized Base64 text
        XmlNode.Text = EncodedText
    'Return the decoded byte array
        M_GridIcon_EmbeddedIconBytes = XmlNode.NodeTypedValue

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release object references
        Set XmlNode = Nothing
    'Release object references
        Set XmlDoc = Nothing
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Release object references
        Set XmlNode = Nothing
    'Release object references
        Set XmlDoc = Nothing
    'Raise a compact descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Embedded grid-icon PNG decode failed. " & _
            "Length=" & VBA.CStr(PayloadLength) & _
            "; Length Mod 4=" & VBA.CStr(PayloadRemainder) & _
            "; Native error=" & VBA.CStr(ErrorNumber) & _
            "; " & ErrorDescription

End Function

Private Function M_GridIcon_NormalizeBase64(ByVal EncodedText As String) As String

'
'------------------------------------------------------------------------------
'                         NORMALIZE BASE64 TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Removes whitespace and common copy/paste artifacts from Base64 text
'
' WHY THIS EXISTS
'   Embedded Base64 stored in VBA can be damaged by line wrapping, copied hidden
'   characters, non-breaking spaces, or accidental formatting artifacts
'
' INPUTS
'   EncodedText
'     Raw Base64 text
'
' RETURNS
'   Normalized Base64 text
'
' BEHAVIOR
'   Removes carriage returns, line feeds, tabs, normal spaces, non-breaking
'   spaces, zero-width spaces, and byte-order-mark characters
'
' ERROR POLICY
'   Does not raise custom errors
'
' DEPENDENCIES
'   VBA.Replace
'   VBA.ChrW$
'
' NOTES
'   This routine does not validate Base64 correctness
'
'   Validation belongs in M_GridIcon_Base64LooksValid
'
' UPDATED
'   2026-05-08
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultText             As String       'Normalized Base64 text

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start from the supplied text
        ResultText = EncodedText

'------------------------------------------------------------------------------
' REMOVE STANDARD WHITESPACE
'------------------------------------------------------------------------------
    'Remove carriage returns
        ResultText = VBA.Replace(ResultText, vbCr, vbNullString)
    'Remove line feeds
        ResultText = VBA.Replace(ResultText, vbLf, vbNullString)
    'Remove tabs
        ResultText = VBA.Replace(ResultText, vbTab, vbNullString)
    'Remove normal spaces
        ResultText = VBA.Replace(ResultText, " ", vbNullString)

'------------------------------------------------------------------------------
' REMOVE COMMON COPY / PASTE ARTIFACTS
'------------------------------------------------------------------------------
    'Remove non-breaking spaces
        ResultText = VBA.Replace(ResultText, VBA.ChrW$(160), vbNullString)
    'Remove zero-width spaces
        ResultText = VBA.Replace(ResultText, VBA.ChrW$(8203), vbNullString)
    'Remove byte-order marks
        ResultText = VBA.Replace(ResultText, VBA.ChrW$(65279), vbNullString)

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the normalized Base64 text
        M_GridIcon_NormalizeBase64 = ResultText

End Function

Private Function M_GridIcon_EmbeddedIconBase64() As String

'
'------------------------------------------------------------------------------
'                         EMBEDDED ICON BASE64
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the embedded Base64 PNG payload used by the DatePicker grid icon
'
' WHY THIS EXISTS
'   The DatePicker should not require a distributed external icon file, but
'   Shape.Fill.UserPicture requires a real file path
'
' INPUTS
'   None
'
' RETURNS
'   Base64 text representing a 64 x 64 transparent PNG icon
'
' BEHAVIOR
'   Builds the Base64 payload by concatenating short string chunks
'   Returns the concatenated Base64 payload as a String
'
' ERROR POLICY
'   Does not raise custom errors
'
' DEPENDENCIES
'   None
'
' NOTES
'   Matches src/ribbon/images/DP_GridIcon_64.png pixel for pixel (64 x 64,
'   8-bit RGBA) but not byte for byte. The decoded PNG is 4,671 bytes against
'   the tracked file's 5,198: it omits the sRGB, gAMA and pHYs chunks and
'   compresses the image data differently
'   Base64 length: 6228
'   Base64 length Mod 4: 0
'
' UPDATED
'   2026-11-09
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EncodedText            As String       'Base64 payload

'------------------------------------------------------------------------------
' BUILD PAYLOAD
'------------------------------------------------------------------------------
    'Append the embedded Base64 chunks
        EncodedText = "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAASBklEQVR42uVbfYwd1XX/nXNn5r19u+u1jXHBBttFoSStEtSK"
        EncodedText = EncodedText & "RoikUIgUVY3U0lYmfzVRm0aVQC1/RFHbKNLalRopqhoV8gFKE0ILIek6QAOtBUEE07pQWoiBLCFAUooB8+EPvPv2fc3Mvad/"
        EncodedText = EncodedText & "zJ07987MW5uoSKQZaeX1ezP349xzfud3fmcW+Bm/6M3eLyIgIvnanff//llnbn7P0eMrz9196+tf3bfvKi0iRETyVi64nGNx"
        EncodedText = EncodedText & "8YHo/IvUH2zbuvn8F48c/f5Hf+fyW+x3ACBvhaFoaWlJAcC+/QdvPNafiDYiqyMt37rn4B2Liw9ES0tLCiJU3v9//SMitCjC"
        EncodedText = EncodedText & "i4uL0e33HLyrP9YiInJsdSL79h/8PAAUa3Rr+MmvxcVFvmxxMcLiIvve8tdfXvr5/3ryGRGRfDAYpiKS/uC/X5bP3njbReVz"
        EncodedText = EncodedText & "b9Xpl2N/8R/uvOS5w6+JaJ3aNeQPH3o6/+Rnv7Kt9gRfdtli9GbXRNhdnLTvdiLLiYjwl7+5/+If/OhFMx6PzcrqmhmPJ/rH"
        EncodedText = EncodedText & "h1/VNy39yxUiwsvLy4mIqOJnyfu3/F1qn7V9N+X7ZUlEhL9+132/9+JrJ/R4NM5XVtdkMknN40//WF//1dsvFBEWkVgKT6yu"
        EncodedText = EncodedText & "Yk+0PgaIEGwM/9Ynb7lgLcWHc03vA/EOEYlAZLJMdz/zh7+649JfPg/9YYq5mQRPPHcEf3L9Ay+BoyFTMaaIQAAQuQ8A8qYj"
        EncodedText = EncodedText & "ghEAYsAggMndI/Zz2OfL+2FAUGyUmcx+4dortl+w6yyMxhnmZzv47qM/kr/4yiMvdJI4hRgGU6aIj8Dk/zYTp9+897o/esZu"
        EncodedText = EncodedText & "koAKp6L65peXl5Nr/u7QZ06Mu1cnc2fMJCqBFoERQDEhHY2hRUFEYIqVYpILdG/7ObOzsxBj3EYp9CL7SfEMEUNAxa0i7jzE"
        EncodedText = EncodedText & "fg+pbd7eIwCyyQi5KT4zIoAYjI0izO/Y1ZntQGsDxQAg79LZ5AP9wfE//7Wr//66v9y949OXX065b4TIbX7PHvqbhw53P/6F"
        EncodedText = EncodedText & "B/6pu+ncD8Iw1sY6z82QREDF/ILJJEVuDIsYiBFoYwAQxplIujIQd8j2frEbIw+ag1uIQBBQ4S8Q95xnRPswEUGLIGJDRETG"
        EncodedText = EncodedText & "SOFpIhADrK6NjclzlKYSQBLF0pvb2gXMn336H1+4cP/+/Vf+5iN7MuwVACQMAIt79hDt3Wu+dfN3buxs2vHBScbpytpIsjyL"
        EncodedText = EncodedText & "CFAEMJFwpBSngzdYjwcQMIwxABhmMsDo+EsUxxErAjMRM8H9EIGJwcxgxcTsfoe9H1wEgTCRMFP1ObP9nomJhOM44tEbr1E6"
        EncodedText = EncodedText & "GhRhZARCxRrStaMcxREzCVPxvMq0jlYGE8mNSuOFc37jr+569fO0d6/ZvfsqBgDevXtJ7d271/z61Tdfis7mj4xTyYfjcRLZ"
        EncodedText = EncodedText & "lfonIRCsHn0BJAbaCIwYGDHQ2qB/4mUYk5f+DhDssyUG1LKzVP8V5/5U+QjZSBIfNgg6T4s1EGDK0xdA5xn6J45ALHaU4xII"
        EncodedText = EncodedText & "RET9wShOc+SIFz7+/j+++b379u3Tu3cvKd5nbx5puVrNbJJxmiFigpBzSut/DK0zpHmONDdgAnJd7EKIkWuNdDwEMYdY54Fe"
        EncodedText = EncodedText & "yVGobo3SGOLNWDNWAYqEPB0h1xpSQCdybUAE5EaQTYYweQq2wEue3ZmZRpMcPLMgucmvAYDXX3+KGPuu0tddt78DFV8ySTUR"
        EncodedText = EncodedText & "iMtAJodHBIhBHMUgFeGOB5/CcDjCpvkOjM6w7/4nkGpBksxAjHEnCqBAeru56sTJbVjE8wPPVcqTrUBRYIxGp9PDKBN8497v"
        EncodedText = EncodedText & "AaKxab6L1dU13PnAMqK4A1axndOagAhEbAGVeJwKiUouffTRR+MHH9ybRwBwaHVlq9G8VQtAJFS5Y+WcxcXYfs4u3P3wM3j1"
        EncodedText = EncodedText & "U1/HJe/ZhceffRn/+uRL2HbueSBWEKNtKhM3hoj9TeAyRGHU0BNcRiCyC5YgRCACYoWzt+/Ezfcs4/kjJ/Hu87fh4e8fxmPP"
        EncodedText = EncodedText & "vobtu94JIQUgh+dQLvlBDGkhaI3Ne257bAHAsQgA0oGJtUAV6clDbbeGwn21ztHbeDZ2MvDk4Zfw0A8fwUyng3N3vgPdTecg"
        EncodedText = EncodedText & "z3IvdaFIHUE4VChPVEG88wIRL2NIGcRVPBPBiMHC1p2IFOHg0y/j/ideRq/bxc7zLkAyfxa0zmzadGYunjdSegFyrfHKK/2K"
        EncodedText = EncodedText & "B7DSQuDqwXLhdmFE5HK21hqdhbPxjk1nw+gUrBLkhpFlKYjZHSiDIFy6uQ9k5SalBhQe2AacgZx3kPWMPNeYPWMHLthyDsRk"
        EncodedText = EncodedText & "IE6QGUae2wOQCjQJ4rZTgUt1KlGN8sKur1iEx8zKIq9AYo2cCEQJJDUQ0WBiVBm8IDQk5P7nxqfCs+pM1VaZDrSodMFwvRBT"
        EncodedText = EncodedText & "rCnPNTIAoASiBYTckkmBFHm7wJDS+j7h8mwfBUxNvEkgpdktBlXkhIndZGXIVDFNDeLj6K0HdD4zFg8oqxRC3rFVX5eGdazR"
        EncodedText = EncodedText & "GLBdpxgTbtCBLNm0bNkl1QwwGRefMRPIA0DyFu9nJaKC3Uh5PFLl3CK+yW2yJJogDwcggLL3U3VKbj6QJariuEQ1OYPgsUx/"
        EncodedText = EncodedText & "YYp9AhDgSzkvK4KhWgjEc4qGRyc01sMA+4nIEQvnnlLGIqpTFQqYDpUsRsSRYReH1KTDzm+s2/uUuRy6DK0SZEXEM7L4SOnR"
        EncodedText = EncodedText & "h/KZimewiqAmGeZ9A1C2Nvrdi3dM3v2ud/a0zl3+dOFDFSRBPO8sF2LTXrFxcWhPHoJX35eRIMHmqzEpDHuq0qWxSF4eRsiW"
        EncodedText = EncodedText & "Ku8TrxapQrS4X0URln+o6KUXBvRYaYCNc2l6+YVnZldcsutnQgfcgFfkvrVuVQ12JgmtjTVpI9BaVycVxNOU/78drtNck4hA"
        EncodedText = EncodedText & "KYW1cQ70vRBYQ6FHKC5yFNHbfcc/+cVMxT7DNNgPqigf7X1i4YdbkGMBm4ra9VgjUhvTs6tMH9Pd26LzhmOiWuc6Y7bQgBoP"
        EncodedText = EncodedText & "qG1+kmYYjychK7P77M10EEXKDdofFFUaAUGCj5RCr9d1Y2ZZjuF4gsa+RDDT7SBJYgeYw9EYWZZb0aS6n4kwOzvjEqTWBoPh"
        EncodedText = EncodedText & "uBmtIuh0EnQ7SbA/aSdCzZPr9weYpJmluGI3TzDGwBjBpoU5gIA0zbDSH3hoTi6DiwBxHCFJYgDAYDjGYDgGc5jujDHIco0t"
        EncodedText = EncodedText & "mxeK8jbXWO0PbGUnAS4ZbRBFCjPdDgBgOJqgPxhZbhJmqEmao5NElS5hZ1xro8ItMeAkGbGEpzhcRpXhy3qbHP2sMftGVDBT"
        EncodedText = EncodedText & "JZbaJTOz5R2CqrdBUFxuvzIqmFtpNBEFHuATspoei7mGAWh628j3DqlVWaGoIaeEzDD7Sx2mPXHUqwg9M8p0iA8FRDccBXBe"
        EncodedText = EncodedText & "/43X65E1H6jYl0y11np9KWmZjZr035NUKuNXAkdAKU8xR1to+37CoRpxOsmWpvfOiFpbDtSYWqakV2rso76stvMs5TYiqmvO"
        EncodedText = EncodedText & "rdsiAkouzOs6q9TjV1p/dV4hTRMJJEDgQspuiOO2CjXBsMZIEHZuHiMNDGh6ZVho+TMaIyERavcVQieJYYyx4ORZj9ihuogg"
        EncodedText = EncodedText & "UhHiKII2uqjWvL0pFSGKlCt1O0mMLM/BzJ7wQWA26HQSB4RKMTpJXKRWQlAERZEKx4wjjBV7pa7VIw2QJHE1V5lFBPUsMN/w"
        EncodedText = EncodedText & "FRHB/FwPs7MzjdN2BYm1ilKEMzZvcCcW3mu1A2vk2V4XMzOd1thUyitfibB50wbbe6DG/MxlC03QdfyhGbLMFFaLduy5OhM0"
        EncodedText = EncodedText & "xrRTx5qA0IYB5cTM7TEduLV4SnHtTq+r5qVMbh9T2tLgFLSQ6RgQVR/yVGBrw27TwhwprEyDUtg/ZTUVeiTwolB3CIduY65t"
        EncodedText = EncodedText & "x+PfK57B1vq1EGigLRVNhzTNA7AqBZFuJ64mJWCSZshzHWQoESCKFDpJBGMEzIwT/RQPLh8LExYBeW7w/l86E2dt7Nh7C4aZ"
        EncodedText = EncodedText & "5TqQyCEAW3zwT3g8mTSMIgJ0kthS9tCwc81iqAmCK6sDjCepJ38VnxtjsGGuh4X5WQiANMtx/I2VqvkhlfbGxDhj0wawUiAi"
        EncodedText = EncodedText & "fO7bz+Gm+57HQi8qOku2El0dZrjy4u244ZqLAAi0Njhxsg9tjBM1SpIhALZsXkAniUEA+sMRTq6uFeEScH5BksTYsnkBtRKl"
        EncodedText = EncodedText & "XgvMT6+fuRQ5q1yrAWjjUWFjySpzrdFBro2u7Own+xNsno2x0IugTeUB3ZjxxtrEYUmWWYBlqhEhFD1JY5y3GSuMluJoPVRK"
        EncodedText = EncodedText & "jCiXZnKNNZsHufSAaeSmkYNdN0dO+1UrHwojRchyDW1b67kxMFKEWxyxJ7PB0/yotava6CGiWpusQ95IEebsofOpuF8borbS"
        EncodedText = EncodedText & "SzoV+pIjIdTKEkOK2qJxtk1W39qUXbQDZmCABgbIm32jLhCrW6avqClxjeOLBFloGqL74vup19KsQqWF9UdtVgne7SA6BeeH"
        EncodedText = EncodedText & "1+SkoiVWc9f6giMmRIpgpCAzzIRIGadKiUgznzpMCTfQ1IWn+RPWF0QgaHUTrXWwYSI4eTp4ccGYoB8wvUgjHFudQGuD3OZ8"
        EncodedText = EncodedText & "Vgr94QTAhuCcxXV2JFDIpcVLdEDkvHKXVesh4nQEkQ1zPYyi5gCwklh5wnEcYeOGOSuJhfVapBhxHDkAveZD5+PnFrogeC9T"
        EncodedText = EncodedText & "EUFrwYfeu831JCOlsDA/i8w2PMVrnLJiJHHsPKU30w3A09WcVhIjanbdmmmQmikwSWJ0OvFpCZhzvZmmlC4VayRbmf3C9nl8"
        EncodedText = EncodedText & "6sO/ONXorviCoNfrtguoHhMVAFHEWNgw21q9B+uUZnBEVRqkVh4g0hIi05TflnvqMZtr3cr5CyywtUfNGG2VONWelSlh16Y6"
        EncodedText = EncodedText & "iEizFsAUWbyEQ5I6ilKDXrYtgpmC12MUc6HzNbh8rZlJRX3SOiaFY9absb6zsLeuQFxq6AH1jEqEwXCM4WgUYoA1xPxcz8nN"
        EncodedText = EncodedText & "IoKTK2vIc416U0kphY0bZl0JPR4XCi7V+oIiwOxMF71e13neSn8NWZqFdb4UNcXChlkoLl7qSNMMq/1BYxvGFFL7/FwvNKSc"
        EncodedText = EncodedText & "JggOR2OkWR6qvTYLjEYTzHQSu4BC6/etXRotTTPM9rqueBmOJ5ikmSudKxS3gNbrFv2DPMdwOA4KrjKcTJZjJutgptsBETCa"
        EncodedText = EncodedText & "pHZMDgJERDAYjTFrx2y7ovVoINk3HIvTk8Dd/DIZtq1WhEH4aku9ni/FDK4tqDhMqqUw8trtCIhUXQx28r347yjBvblSx8e1"
        EncodedText = EncodedText & "UzLB2tuNJQUWaUfB8DuawuKmy6XrAay0HU4LNRafe8r0ef1ymNcVhWWdVZ2iepDT5OnSqmbTNJlyisxdb8a0P98m7FQGqEtP"
        EncodedText = EncodedText & "9VfcCC2y0/oGkilGlcYUza7OKewbZlyR9jR4Gk1tPq0JPfd2oeDevoB7iSqUsL3PxW90SMMI5SZM7RU2n4f4b402lR8JQtDR"
        EncodedText = EncodedText & "Z2mXxaVugCyNqU05TJJC9iIue3dcCA+lJGUfiaIIcRyBrApLDjgLmhxFyk3csVIal708r6lRKjyFEFNQaNcM8qTxKFKIS1kc"
        EncodedText = EncodedText & "hfStFLvx/LFLWbzWGaJOEpHLAju2bsiMER1ggQjm52Yxa1vbdd9STI6OKkXYEsji3qtWtsHq8/Zup9MaNszsxmQmbNm8YIss"
        EncodedText = EncodedText & "tGaSUhaf6XaQxPHUMSXs2EJEJslcIT9F9k/NVr52x32H09xsTCJljDGqRFSlONDjUNPVyBNKlaKakkGt8am4ev9wWk+xfCRS"
        EncodedText = EncodedText & "KvySmqKMk+4a/DrEcVbKpLnhteHohb3XXrsqIsQHDhxQAKTfH93z2vFV64GmwbNLqt+YvJYGw/tkKrTUx5Optbs395RmZzCm"
        EncodedText = EncodedText & "tI9pm27yyvEVWhuM7wWAAwcOKD5w4IABADMef+k/Di0PDcBKRWba4t8kyL4trqJ9p0xuwP/5vacGufAN1gDFSZd/EHnDN/75"
        EncodedText = EncodedText & "Y//++HMiIrmI5Hmey/+Hn3I/Bw89K1+89dsf8/fsrvKD62+9+xPfeegJ6Y9SEREjItlP+Y/pD1O576En5Eu33f2J+uapboSr"
        EncodedText = EncodedText & "rrpK/+0td37gzI2b9m7besb7ztuxPSgmqldmZXooBLIYVdzVgqP/brDUKKr7e0O0vpw29a0F/92kMv/3ByM8/+IRvHr0xMFj"
        EncodedText = EncodedText & "J04u/ulHfvu75R6njunf8Lmbbv+V3szMRUTYaYzEdeWmAJew4OEAeELCYfxvTLMBylw0SKu77H3F33fZMUw1ov3Ofg0GFWBn"
        EncodedText = EncodedText & "AGY2eW7+J9fZI9d+9MpDALB7aUnt8zY/9VpaWlKNPz39Kb5EhBoxb6//BTrnfQKAxO/pAAAAAElFTkSuQmCC"

'------------------------------------------------------------------------------
' RETURN PAYLOAD
'------------------------------------------------------------------------------
    'Return the Base64 payload
        M_GridIcon_EmbeddedIconBase64 = EncodedText

End Function
