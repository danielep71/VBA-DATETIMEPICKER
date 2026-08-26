VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_DatePicker 
   Caption         =   "DATETIME PICKER"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "UF_DatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' MODULE: UF_DatePicker
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the runtime UserForm implementation for the Excel DatePicker UI
'
' WHY THIS EXISTS
'   The DatePicker form is built primarily at runtime. This keeps layout,
'   dynamic control creation, label-event routing, calendar rendering,
'   month/year navigation, footer shortcuts, in-form settings, keyboard
'   navigation, and live-clock behavior centralized in one UserForm code-behind
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Manages:
'     - UserForm lifecycle, formatting, activation, keyboard routing, and teardown
'     - runtime creation and reuse of labels, frames, pages, MultiPage controls,
'       ComboBoxes, and CheckBoxes
'     - header month/year captions, navigation arrows, and compact settings entry
'     - weekday-row captions with Sunday/Monday and local/fixed-English modes
'     - fixed 6 x 7 calendar-grid rendering
'     - paired day-cell labels for larger hover and click targets
'     - today, selected-date, keyboard-date, outside-month, disabled, weekend,
'       and hover visual states
'     - month and year selection through a reusable overlay picker panel
'     - footer shortcuts for Today and Now
'     - in-form settings through a label-tabbed MultiPage panel
'     - settings persistence through the explicit Save action
'     - live footer-clock updates through shared timer infrastructure
'     - cleanup of timers, hook collections, cached controls, hover trackers,
'       keyboard state, and overlay state
'
' ERROR POLICY
'   Runtime build, navigation, selection, and settings routines raise descriptive
'   runtime errors when they cannot complete safely
'
'   High-frequency visual cleanup paths, timer callbacks, hover-reset routines,
'   and form teardown routines are best-effort and intentionally suppress errors
'
' DEPENDENCIES
'   MSForms.UserForm
'   MSForms.Label
'   MSForms.Frame
'   MSForms.MultiPage
'   MSForms.Page
'   MSForms.ComboBox
'   MSForms.CheckBox
'   StdFont
'
'   cDatePickerLabelHook
'
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_Settings_IsValidFirstDayOfWeek
'
'   M_FormBridge_ConsumeInitialDate
'
'   M_Caption_GetMonth
'   M_Caption_GetDate
'
'   M_Platform_ShouldUseWinAPI
'   M_Window_RemoveTitleBar
'   M_Window_MoveFormToMouse
'
'   M_Timer_ApplyClockMode
'   M_Timer_Stop
'
'   M_ContextMenu_Update
'   M_KeyboardShortcut_Update
'   M_GridIcon_Remove
'
'   M_DatePolicy_CanSelectDate
'   M_Picker_SelectDate
'   DP_Today
'   DP_Now
'   DP_Close
'
'   gDP_FirstDayOfWeek
'   gDP_UseLocalNames
'   gDP_ClockMode
'   gDP_SizeMode
'   gDP_HighlightWeekends
'   gDP_AllowOutsideMonthSelection
'   gDP_CloseAfterSelection
'   gDP_ShowRightClick
'   gDP_ShowGridIcon
'   gDP_UseWinAPI
'   gDP_EnableKeyboardShortcut
'   gDP_HasSelectedDate
'   gDP_SelectedDate
'
' NOTES
'   Runtime-created controls are not available as direct VBA member variables.
'   Use Me.Controls or cached references returned by the UF_Ensure_* helpers
'
'   The picker panel is reusable. It displays either months or years depending
'   on mPickerPanelMode
'
'   Year-panel scrolling is handled by the header year-arrow labels when the
'   year picker panel is visible
'
'   The live clock updates only Lbl_Time.Caption. It must not repaint, rebuild,
'   or reformat the footer
'
'   Date selection is delegated to M_Picker_SelectDate after the selected date
'   has been validated by M_DatePolicy_CanSelectDate
'
'   The settings panel is an in-form overlay. It replaces the older separate
'   modal settings form pattern
'
'   Compact layout may hide the footer. In compact mode, settings remain
'   accessible through the header settings icon
'
'   Developer-owned layout constants are intentionally not validated in normal
'   runtime paths. Use dedicated debug / regression routines for layout checks
'
'   This is UserForm code-behind. Paste it into UF_DatePicker
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------

    '--------------------------------FORM--------------------------------------
    Private Const DP_FORM_CAPTION                       As String = "DATETIME PICKER"    'UserForm caption
    Private Const DP_FORM_WIDTH                         As Single = 223                  'Borderless UserForm width
    Private Const DP_FORM_HEIGHT                        As Single = 269                  'Normal borderless UserForm height
    Private Const DP_FORM_HEIGHT_COMPACT                As Single = 210                  'Compact borderless UserForm height
    Private Const DP_FORM_NATIVE_WIDTH_COMPENSATION     As Single = 6                    'Additional width when native title bar is visible
    Private Const DP_FORM_NATIVE_HEIGHT_COMPENSATION    As Single = 18                   'Additional height when native title bar is visible
    Private Const DP_FORM_BACK_COLOR                    As Long = vbWhite                'UserForm background color
    Private Const DP_FORM_FORE_COLOR                    As Long = vbButtonText           'UserForm foreground color
    Private Const DP_FORM_FONT_NAME                     As String = "Segoe UI"           'UserForm font name
    Private Const DP_FORM_FONT_SIZE                     As Single = 9                    'UserForm font size
    Private Const DP_FORM_ZOOM                          As Single = 100                  'UserForm zoom percentage
    Private Const DP_FORM_STARTUP_POSITION              As Long = 0                      'Manual startup position

    '-----------------------------POSITIONING----------------------------------
    Private Const DP_FORM_MOUSE_OFFSET_XPX              As Long = 10                     'Mouse-position X offset in pixels
    Private Const DP_FORM_MOUSE_OFFSET_YPX              As Long = 0                      'Mouse-position Y offset in pixels
    Private Const DP_FORM_CENTER_ON_MOUSE               As Boolean = False               'True to center form on mouse

    '-----------------------------KEYBOARD-------------------------------------
    Private Const DP_KEY_CTRL_MASK                      As Integer = 2                   'Ctrl key mask

    '-------------------------------HEADER-------------------------------------
    Private Const DP_HEADER_TOP                         As Single = 0                    'Header banner top
    Private Const DP_HEADER_HEIGHT                      As Single = 52                   'Header banner height
    Private Const DP_HEADER_BACK_COLOR                  As Long = 9985057                'Header background color
    Private Const DP_HEADER_FORE_COLOR                  As Long = vbWhite                'Header foreground color
    Private Const DP_HEADER_HOVER_BACK_COLOR            As Long = &HAE7546               'Header hover background color

    Private Const DP_HEADER_MONTH_REL_LEFT              As Single = 24                   'Header month relative left
    Private Const DP_HEADER_MONTH_TOP                   As Single = 8                    'Header month top
    Private Const DP_HEADER_MONTH_WIDTH                 As Single = 84                   'Header month width
    Private Const DP_HEADER_MONTH_HEIGHT                As Single = 20                   'Header month height

    Private Const DP_HEADER_YEAR_REL_LEFT               As Single = 152                  'Header year relative left
    Private Const DP_HEADER_YEAR_TOP                    As Single = 9                    'Header year top
    Private Const DP_HEADER_YEAR_WIDTH                  As Single = 42                   'Header year width
    Private Const DP_HEADER_YEAR_HEIGHT                 As Single = 20                   'Header year height

    Private Const DP_HEADER_PREVMONTH_ARROW_LEFT        As Single = 4                    'Previous-month arrow relative left
    Private Const DP_HEADER_NEXTMONTH_ARROW_LEFT        As Single = 114                  'Next-month arrow relative left
    Private Const DP_HEADER_MONTH_ARROW_TOP             As Single = 10                   'Month-arrow top
    Private Const DP_HEADER_MONTH_ARROW_WIDTH           As Single = 14                   'Month-arrow width
    Private Const DP_HEADER_MONTH_ARROW_HEIGHT          As Single = 18                   'Month-arrow height

    Private Const DP_HEADER_YEAR_ARROW_LEFT             As Single = 195                  'Year-arrow relative left
    Private Const DP_HEADER_YEAR_ARROW_TOP              As Single = 7                    'Year-arrow top
    Private Const DP_HEADER_YEAR_ARROW_WIDTH            As Single = 11                   'Year-arrow width
    Private Const DP_HEADER_YEAR_ARROW_HEIGHT           As Single = 11                   'Year-arrow height

    '--------------------------HEADER SETTINGS--------------------------------
    Private Const DP_HEADER_SETTINGS_ICON_GAP           As Single = 7                    'Gap after next-month label
    Private Const DP_HEADER_SETTINGS_ICON_TOP_OFFSET    As Single = 0                    'Vertical offset against next-month label
    Private Const DP_HEADER_SETTINGS_ICON_WIDTH         As Single = 16                   'Compact settings icon width
    Private Const DP_HEADER_SETTINGS_ICON_HEIGHT        As Single = 16                   'Compact settings icon height
    Private Const DP_HEADER_SETTINGS_ICON_FONT_NAME     As String = "Segoe MDL2 Assets"  'Compact settings icon font
    Private Const DP_HEADER_SETTINGS_ICON_FONT_SIZE     As Single = 10                   'Compact settings icon font size
    Private Const DP_HEADER_SETTINGS_ICON_CODEPOINT     As Long = &HE713                 'Segoe MDL2 settings glyph
    Private Const DP_HEADER_SETTINGS_ICON_TOOLTIP       As String = "Settings"           'Compact settings tooltip

    '-----------------------------DAY OF WEEK----------------------------------
    Private Const DP_DOW_LABEL_WIDTH                    As Single = 30                   'Weekday label width
    Private Const DP_DOW_LABEL_HEIGHT                   As Single = 16                   'Weekday label height
    Private Const DP_DOW_LABEL_HORIZONTAL_STEP          As Single = 30                   'Weekday label horizontal step

    '--------------------------------DAYS--------------------------------------
    Private Const DP_DAY_GRID_START_LEFT                As Single = 10                   'Calendar grid left
    Private Const DP_DAY_GRID_START_TOP                 As Single = 60                   'Calendar grid first-row top
    Private Const DP_DAY_GRID_ROWS                      As Long = 6                      'Calendar grid rows
    Private Const DP_DAY_LABELS_PER_ROW                 As Long = 7                      'Calendar grid columns
    Private Const DP_DAY_LABEL_COUNT                    As Long = 42                     'Calendar grid cell count

    Private Const DP_DAY_LABEL_WIDTH                    As Single = 18                   'Day text label width
    Private Const DP_DAY_LABEL_HEIGHT                   As Single = 15                   'Day text label height
    Private Const DP_DAY_LABEL_HORIZONTAL_STEP          As Single = 30                   'Day text label horizontal step
    Private Const DP_DAY_LABEL_VERTICAL_STEP            As Single = 25                   'Day text label vertical step

    Private Const DP_DAY_CELL_WIDTH                     As Single = 24                   'Day background cell width
    Private Const DP_DAY_CELL_HEIGHT                    As Single = 22                   'Day background cell height

    Private Const DP_DAY_NORMAL_BACK_COLOR              As Long = vbWhite                'Normal day background
    Private Const DP_DAY_NORMAL_FORE_COLOR              As Long = vbButtonText           'Normal day foreground
    Private Const DP_DAY_CURRENT_MONTH_FORE_COLOR       As Long = vbButtonText           'Current-month day foreground
    Private Const DP_DAY_OUTSIDE_MONTH_FORE_COLOR       As Long = &H808080               'Outside-month day foreground

    Private Const DP_DAY_HOVER_BACK_COLOR               As Long = &HF2F2F2               'Day hover background
    Private Const DP_DAY_HOVER_BORDER_COLOR             As Long = &HC8C8C8               'Day hover border

    Private Const DP_DAY_TODAY_BACK_COLOR               As Long = &HF2F7FC               'Today background
    Private Const DP_DAY_TODAY_BORDER_COLOR             As Long = 9985057                'Today border
    Private Const DP_DAY_SELECTED_BACK_COLOR            As Long = 9985057                'Selected day background
    Private Const DP_DAY_SELECTED_FORE_COLOR            As Long = vbWhite                'Selected day foreground

    '-------------------------------DIVIDER------------------------------------
    Private Const DP_DIVIDER_HEIGHT                     As Single = 2                    'Divider height
    Private Const DP_DIVIDER_SIDE_MARGIN                As Single = 0                    'Divider side margin
    Private Const DP_DIVIDER_COLOR                      As Long = &HE6E6E6               'Divider color
    Private Const DP_DIVIDER_BACK_COLOR                 As Long = vbWhite                'Divider background color

    '-------------------------------FOOTER-------------------------------------
    Private Const DP_FOOTER_TOP_GAP                     As Single = 10                   'Gap between grid and footer
    Private Const DP_FOOTER_HEIGHT                      As Single = 50                   'Footer height
    Private Const DP_FOOTER_SIDE_PADDING                As Single = 8                    'Footer side padding
    Private Const DP_FOOTER_BACK_COLOR                  As Long = &HF7F7F7               'Footer background color
    Private Const DP_FOOTER_FORE_COLOR                  As Long = vbButtonText           'Footer foreground color

    '--------------------------FOOTER GROUPS-----------------------------------
    Private Const DP_FOOTER_GROUP_WIDTH                 As Single = 100                  'Footer shortcut group width
    Private Const DP_FOOTER_GROUP_HEIGHT                As Single = 38                   'Footer shortcut group height
    Private Const DP_FOOTER_GROUP_LEFT_REL_LEFT         As Single = 4                    'Left shortcut group relative left
    Private Const DP_FOOTER_GROUP_RIGHT_REL_LEFT        As Single = 94                   'Right shortcut group relative left
    Private Const DP_FOOTER_GROUP_ICON_OFFSET_LEFT      As Single = 11                   'Footer group icon offset left
    Private Const DP_FOOTER_GROUP_TEXT_OFFSET_LEFT      As Single = 22                   'Footer group text offset left
    Private Const DP_FOOTER_GROUP_TEXT_WIDTH            As Single = 76                   'Footer group text width

    '--------------------------FOOTER LEFT: TODAY------------------------------
    Private Const DP_FOOTER_TODAY_ICON_REL_LEFT         As Single = DP_FOOTER_GROUP_LEFT_REL_LEFT + DP_FOOTER_GROUP_ICON_OFFSET_LEFT       'Today icon relative left
    Private Const DP_FOOTER_TODAY_CAPTION_REL_LEFT      As Single = DP_FOOTER_GROUP_LEFT_REL_LEFT + DP_FOOTER_GROUP_TEXT_OFFSET_LEFT       'Today caption relative left
    Private Const DP_FOOTER_TODAY_VALUE_REL_LEFT        As Single = DP_FOOTER_GROUP_LEFT_REL_LEFT + DP_FOOTER_GROUP_TEXT_OFFSET_LEFT       'Today value relative left
    Private Const DP_FOOTER_TODAY_CAPTION_TOP           As Single = 10                   'Today caption top
    Private Const DP_FOOTER_TODAY_CAPTION_WIDTH         As Single = DP_FOOTER_GROUP_TEXT_WIDTH 'Today caption width
    Private Const DP_FOOTER_TODAY_CAPTION_HEIGHT        As Single = 16                   'Today caption height
    Private Const DP_FOOTER_TODAY_VALUE_TOP             As Single = 30                   'Today value top
    Private Const DP_FOOTER_TODAY_VALUE_WIDTH           As Single = DP_FOOTER_GROUP_TEXT_WIDTH 'Today value width
    Private Const DP_FOOTER_TODAY_VALUE_HEIGHT          As Single = 16                   'Today value height
    Private Const DP_FOOTER_TODAY_HALO_REL_LEFT         As Single = DP_FOOTER_GROUP_LEFT_REL_LEFT 'Today halo relative left
    Private Const DP_FOOTER_TODAY_HALO_WIDTH            As Single = DP_FOOTER_GROUP_WIDTH 'Today halo width

    '--------------------------FOOTER RIGHT: TIME------------------------------
    Private Const DP_FOOTER_TIME_ICON_REL_LEFT          As Single = DP_FOOTER_GROUP_RIGHT_REL_LEFT + DP_FOOTER_GROUP_ICON_OFFSET_LEFT      'Time icon relative left
    Private Const DP_FOOTER_TIME_CAPTION_REL_LEFT       As Single = DP_FOOTER_GROUP_RIGHT_REL_LEFT + DP_FOOTER_GROUP_TEXT_OFFSET_LEFT      'Time caption relative left
    Private Const DP_FOOTER_TIME_VALUE_REL_LEFT         As Single = DP_FOOTER_GROUP_RIGHT_REL_LEFT + DP_FOOTER_GROUP_TEXT_OFFSET_LEFT      'Time value relative left
    Private Const DP_FOOTER_TIME_CAPTION_TOP            As Single = 10                   'Time caption top
    Private Const DP_FOOTER_TIME_CAPTION_WIDTH          As Single = 66                   'Time caption width
    Private Const DP_FOOTER_TIME_CAPTION_HEIGHT         As Single = 16                   'Time caption height
    Private Const DP_FOOTER_TIME_VALUE_TOP              As Single = 30                   'Time value top
    Private Const DP_FOOTER_TIME_VALUE_WIDTH            As Single = 66                   'Time value width
    Private Const DP_FOOTER_TIME_VALUE_HEIGHT           As Single = 16                   'Time value height
    Private Const DP_FOOTER_TIME_HALO_REL_LEFT          As Single = DP_FOOTER_GROUP_RIGHT_REL_LEFT 'Time halo relative left
    Private Const DP_FOOTER_TIME_HALO_WIDTH             As Single = 88                   'Time halo width

    '-----------------------------FOOTER HOVER---------------------------------
    Private Const DP_FOOTER_HALO_TOP                    As Single = 7                    'Footer halo top
    Private Const DP_FOOTER_HALO_HEIGHT                 As Single = DP_FOOTER_GROUP_HEIGHT 'Footer halo height
    Private Const DP_FOOTER_HALO_BACK_COLOR             As Long = DP_HEADER_HOVER_BACK_COLOR 'Footer halo background
    Private Const DP_FOOTER_HOVER_FORE_COLOR            As Long = vbWhite                'Footer hover foreground
    
    '-----------------------------FOOTER ICONS---------------------------------
    Private Const DP_FOOTER_ICON_FONT_NAME              As String = "Segoe MDL2 Assets"  'Footer icon font
    Private Const DP_FOOTER_ICON_FORE_COLOR             As Long = &H808000               'Footer icon foreground
    Private Const DP_FOOTER_ICON_TOP                    As Single = 17                   'Footer icon top
    Private Const DP_FOOTER_ICON_WIDTH                  As Single = 18                   'Footer icon width
    Private Const DP_FOOTER_ICON_HEIGHT                 As Single = 18                   'Footer icon height
    Private Const DP_FOOTER_ICON_FONT_SIZE              As Single = 14                   'Footer icon font size
    Private Const DP_FOOTER_ICON_CALENDAR_CODEPOINT     As Long = &HE8BF                 'Segoe MDL2 calendar glyph
    Private Const DP_FOOTER_ICON_CLOCK_CODEPOINT        As Long = &HE121                 'Segoe MDL2 clock glyph

    '--------------------------FOOTER SETTINGS--------------------------------
    Private Const DP_FOOTER_SEPARATOR_REL_LEFT          As Single = 180                  'Settings separator relative left
    Private Const DP_FOOTER_SEPARATOR_TOP               As Single = 9                    'Settings separator top
    Private Const DP_FOOTER_SEPARATOR_WIDTH             As Single = 1                    'Settings separator width
    Private Const DP_FOOTER_SEPARATOR_HEIGHT            As Single = 32                   'Settings separator height
    Private Const DP_FOOTER_SEPARATOR_COLOR             As Long = &HE0E0E0               'Settings separator color

    Private Const DP_FOOTER_SETTINGS_ICON_REL_LEFT      As Single = 190                  'Settings icon relative left
    Private Const DP_FOOTER_SETTINGS_ICON_TOP           As Single = 17                   'Settings icon top
    Private Const DP_FOOTER_SETTINGS_ICON_WIDTH         As Single = 18                   'Settings icon width
    Private Const DP_FOOTER_SETTINGS_ICON_HEIGHT        As Single = 18                   'Settings icon height
    Private Const DP_FOOTER_SETTINGS_ICON_FONT_NAME     As String = "Segoe MDL2 Assets"  'Settings icon font
    Private Const DP_FOOTER_SETTINGS_ICON_FONT_SIZE     As Single = 14                   'Settings icon font size
    Private Const DP_FOOTER_SETTINGS_ICON_CODEPOINT     As Long = &HE713                 'Segoe MDL2 settings glyph
    Private Const DP_FOOTER_SETTINGS_ICON_CAPTION       As String = "Settings"           'Settings tooltip

    Private Const DP_FOOTER_SETTINGS_HALO_REL_LEFT      As Single = 184                  'Settings halo relative left
    Private Const DP_FOOTER_SETTINGS_HALO_WIDTH         As Single = 30                   'Settings halo width
    Private Const DP_FOOTER_SETTINGS_HALO_TOP           As Single = DP_FOOTER_HALO_TOP    'Settings halo top
    Private Const DP_FOOTER_SETTINGS_HALO_HEIGHT        As Single = DP_FOOTER_HALO_HEIGHT 'Settings halo height

    '-----------------------------PICKER PANEL---------------------------------
    Private Const DP_PICKER_PANEL_NAME                  As String = "Fra_PickerPanel"    'Picker panel frame name
    Private Const DP_PICKER_PANEL_WIDTH                 As Single = 218                  'Picker panel width
    Private Const DP_PICKER_PANEL_HEIGHT                As Single = 154                  'Picker panel height
    Private Const DP_PICKER_PANEL_BACK_COLOR            As Long = vbWhite                'Picker panel background

    Private Const DP_PICKER_PANEL_MODE_NONE             As Long = 0                      'No picker panel mode
    Private Const DP_PICKER_PANEL_MODE_MONTHS           As Long = 1                      'Month picker panel mode
    Private Const DP_PICKER_PANEL_MODE_YEARS            As Long = 2                      'Year picker panel mode

    Private Const DP_PICKER_ITEM_COUNT                  As Long = 12                     'Picker item count
    Private Const DP_PICKER_ITEMS_PER_ROW               As Long = 3                      'Picker items per row
    Private Const DP_PICKER_ITEM_TOP                    As Single = 9                    'Picker item first top
    Private Const DP_PICKER_ITEM_WIDTH                  As Single = 58                   'Picker item width
    Private Const DP_PICKER_ITEM_HEIGHT                 As Single = 28                   'Picker item height
    Private Const DP_PICKER_ITEM_HORIZONTAL_STEP        As Single = 66                   'Picker item horizontal step
    Private Const DP_PICKER_ITEM_VERTICAL_STEP          As Single = 34                   'Picker item vertical step
    Private Const DP_PICKER_ITEM_TEXT_TOP_OFFSET        As Single = 9                    'Picker item text top offset
    Private Const DP_PICKER_ITEM_TEXT_HEIGHT            As Single = 16                   'Picker item text height
    Private Const DP_PICKER_ITEM_BACK_COLOR             As Long = vbWhite                'Picker item background
    Private Const DP_PICKER_ITEM_BORDER_COLOR           As Long = &HE6E6E6               'Picker item border
    Private Const DP_PICKER_ITEM_FORE_COLOR             As Long = vbButtonText           'Picker item foreground
    Private Const DP_PICKER_ITEM_HOVER_BACK_COLOR       As Long = &HF2F2F2               'Picker item hover background
    Private Const DP_PICKER_ITEM_HOVER_BORDER_COLOR     As Long = &HC8C8C8               'Picker item hover border

    '----------------------------SETTINGS PANEL--------------------------------
    Private Const DP_SETTINGS_PANEL_NAME                As String = "Fra_Settings"       'Settings panel frame name
    Private Const DP_SETTINGS_PANEL_BACK_COLOR          As Long = vbWhite                'Settings panel background
    Private Const DP_SETTINGS_TITLE_CAPTION             As String = "SETTINGS"           'Settings title caption
    Private Const DP_SETTINGS_TITLE_TOP                 As Single = 10                   'Settings title top
    Private Const DP_SETTINGS_TITLE_LEFT                As Single = 0                    'Settings title left
    Private Const DP_SETTINGS_TITLE_WIDTH               As Single = DP_PICKER_PANEL_WIDTH 'Settings title width
    Private Const DP_SETTINGS_TITLE_HEIGHT              As Single = 18                   'Settings title height

    '-------------------------SETTINGS HEADER ICONS----------------------------
    Private Const DP_SETTINGS_HEADER_ICON_FONT_NAME     As String = "Segoe MDL2 Assets"  'Settings header icon font
    Private Const DP_SETTINGS_HEADER_ICON_TOP           As Single = 7                    'Settings header icon top
    Private Const DP_SETTINGS_HEADER_ICON_WIDTH         As Single = 16                   'Settings header icon width
    Private Const DP_SETTINGS_HEADER_ICON_HEIGHT        As Single = 16                   'Settings header icon height
    Private Const DP_SETTINGS_HEADER_ICON_FORE_COLOR    As Long = &H808080               'Settings header icon foreground
    Private Const DP_SETTINGS_HEADER_ICON_HOVER_FORE_COLOR As Long = DP_HEADER_BACK_COLOR 'Settings header hover foreground
    Private Const DP_SETTINGS_HEADER_ICON_HOVER_BACK_COLOR As Long = &HF2F2F2             'Settings header hover background
    Private Const DP_SETTINGS_HEADER_ICON_HOVER_BORDER_COLOR As Long = &HD9D9D9           'Settings header hover border

    Private Const DP_SETTINGS_SAVE_LEFT                 As Single = 176                  'Save icon left
    Private Const DP_SETTINGS_SAVE_WIDTH                As Single = 18                   'Save icon width
    Private Const DP_SETTINGS_SAVE_HEIGHT               As Single = 16                   'Save icon height
    Private Const DP_SETTINGS_SAVE_FONT_SIZE            As Single = 11                   'Save icon font size
    Private Const DP_SETTINGS_SAVE_CODEPOINT            As Long = &HE74E                 'Segoe MDL2 save glyph
    Private Const DP_SETTINGS_SAVE_TOOLTIP              As String = "Save settings"      'Save tooltip

    Private Const DP_SETTINGS_CLOSE_LEFT                As Single = 194                  'Close icon left
    Private Const DP_SETTINGS_CLOSE_FONT_SIZE           As Single = 9                    'Close icon font size
    Private Const DP_SETTINGS_CLOSE_CODEPOINT           As Long = &HE8BB                 'Segoe MDL2 close glyph
    Private Const DP_SETTINGS_CLOSE_TOOLTIP             As String = "Close settings"     'Close tooltip

    '--------------------------SETTINGS MULTIPAGE------------------------------
    Private Const DP_SETTINGS_MULTIPAGE_NAME            As String = "Mp_Settings"        'Settings MultiPage name
    Private Const DP_SETTINGS_MULTIPAGE_LEFT            As Single = 7                    'Settings MultiPage left
    Private Const DP_SETTINGS_MULTIPAGE_TOP             As Single = 50                   'Settings MultiPage top
    Private Const DP_SETTINGS_MULTIPAGE_WIDTH           As Single = 204                  'Settings MultiPage width
    Private Const DP_SETTINGS_MULTIPAGE_HEIGHT          As Single = 96                   'Settings MultiPage height

    Private Const DP_SETTINGS_PAGE_BACK_LEFT            As Single = 0                    'Settings page background left
    Private Const DP_SETTINGS_PAGE_BACK_TOP             As Single = 0                    'Settings page background top
    Private Const DP_SETTINGS_PAGE_BACK_WIDTH           As Single = 206                  'Settings page background width
    Private Const DP_SETTINGS_PAGE_BACK_HEIGHT          As Single = 96                   'Settings page background height
    Private Const DP_SETTINGS_PAGE_BACK_COLOR           As Long = vbWhite                'Settings page background color

    Private Const DP_SETTINGS_PAGE_DISPLAY_INDEX        As Long = 0                      'Display settings page index
    Private Const DP_SETTINGS_PAGE_BEHAVIOR_INDEX       As Long = 1                      'Behavior settings page index
    Private Const DP_SETTINGS_PAGE_INTEGRATION_INDEX    As Long = 2                      'Integration settings page index
    Private Const DP_SETTINGS_PAGE_COUNT                As Long = 3                      'Settings page count

    Private Const DP_SETTINGS_PAGE_DISPLAY_CAPTION      As String = "Display Settings"   'Display settings page caption
    Private Const DP_SETTINGS_PAGE_BEHAVIOR_CAPTION     As String = "Behavior Settings"  'Behavior settings page caption
    Private Const DP_SETTINGS_PAGE_INTEGRATION_CAPTION  As String = "Integration Settings" 'Integration settings page caption

    '--------------------------SETTINGS FAKE TABS------------------------------
    Private Const DP_SETTINGS_TAB_DISPLAY_NAME          As String = "Lbl_SettingsTabDisplay" 'Display tab label name
    Private Const DP_SETTINGS_TAB_BEHAVIOR_NAME         As String = "Lbl_SettingsTabBehavior" 'Behavior tab label name
    Private Const DP_SETTINGS_TAB_INTEGRATION_NAME      As String = "Lbl_SettingsTabIntegration" 'Integration tab label name
    Private Const DP_SETTINGS_TAB_LEFT                  As Single = 7                    'Settings tab left
    Private Const DP_SETTINGS_TAB_TOP                   As Single = 32                   'Settings tab top
    Private Const DP_SETTINGS_TAB_HEIGHT                As Single = 18                   'Settings tab height
    Private Const DP_SETTINGS_TAB_GAP                   As Single = 1                    'Settings tab gap
    Private Const DP_SETTINGS_TAB_DISPLAY_WIDTH         As Single = 60                   'Display tab width
    Private Const DP_SETTINGS_TAB_BEHAVIOR_WIDTH        As Single = 66                   'Behavior tab width
    Private Const DP_SETTINGS_TAB_INTEGRATION_WIDTH     As Single = 76                   'Integration tab width
    Private Const DP_SETTINGS_TAB_DISPLAY_CAPTION       As String = "Display"            'Display tab caption
    Private Const DP_SETTINGS_TAB_BEHAVIOR_CAPTION      As String = "Behavior"           'Behavior tab caption
    Private Const DP_SETTINGS_TAB_INTEGRATION_CAPTION   As String = "Integration"        'Integration tab caption
    Private Const DP_SETTINGS_TAB_NORMAL_BACK_COLOR     As Long = &HF2F2F2               'Inactive tab background
    Private Const DP_SETTINGS_TAB_SELECTED_BACK_COLOR   As Long = vbWhite                'Active tab background
    Private Const DP_SETTINGS_TAB_BORDER_COLOR          As Long = &HD9D9D9               'Tab border
    Private Const DP_SETTINGS_TAB_NORMAL_FORE_COLOR     As Long = vbButtonText           'Inactive tab foreground
    Private Const DP_SETTINGS_TAB_SELECTED_FORE_COLOR   As Long = DP_HEADER_BACK_COLOR   'Active tab foreground

    '----------------------SETTINGS DISPLAY PAGE CONTROLS----------------------
    Private Const DP_SETTINGS_FIRSTDAY_LABEL_LEFT       As Single = 8                    'First-day label left
    Private Const DP_SETTINGS_FIRSTDAY_LABEL_TOP        As Single = 6                    'First-day label top
    Private Const DP_SETTINGS_FIRSTDAY_LABEL_WIDTH      As Single = 82                   'First-day label width
    Private Const DP_SETTINGS_FIRSTDAY_LABEL_HEIGHT     As Single = 14                   'First-day label height
    Private Const DP_SETTINGS_FIRSTDAY_LABEL_CAPTION    As String = "First day of the week" 'First-day label caption
    Private Const DP_SETTINGS_FIRSTDAY_COMBO_LEFT       As Single = 96                   'First-day ComboBox left
    Private Const DP_SETTINGS_FIRSTDAY_COMBO_TOP        As Single = 3                    'First-day ComboBox top
    Private Const DP_SETTINGS_FIRSTDAY_COMBO_WIDTH      As Single = 84                   'First-day ComboBox width
    Private Const DP_SETTINGS_FIRSTDAY_COMBO_HEIGHT     As Single = 16                   'First-day ComboBox height

    Private Const DP_SETTINGS_CHECKBOX_LEFT             As Single = 8                    'Settings CheckBox left
    Private Const DP_SETTINGS_CHECKBOX_WIDTH            As Single = 178                  'Settings CheckBox width
    Private Const DP_SETTINGS_CHECKBOX_HEIGHT           As Single = 16                   'Settings CheckBox height
    Private Const DP_SETTINGS_CHECKBOX_FONT_SIZE        As Single = 8.25                 'Settings CheckBox font size
    Private Const DP_SETTINGS_CHECKBOX_TOP_1            As Single = 24                   'Settings first CheckBox top
    Private Const DP_SETTINGS_CHECKBOX_VERTICAL_STEP    As Single = 17                   'Settings CheckBox vertical step

    Private Const DP_SETTINGS_CHECK_LOCAL_NAMES_CAPTION As String = "Use local names"    'Local names CheckBox caption
    Private Const DP_SETTINGS_CHECK_LIVE_CLOCK_CAPTION  As String = "Live clock"         'Live clock CheckBox caption
    Private Const DP_SETTINGS_CHECK_COMPACT_CAPTION     As String = "Compact layout"     'Compact layout CheckBox caption
    Private Const DP_SETTINGS_CHECK_WEEKENDS_CAPTION    As String = "Highlight weekends" 'Weekend highlight CheckBox caption

    '--------------------------SETTINGS BEHAVIOR PAGE--------------------------
    Private Const DP_SETTINGS_CHECK_ALLOW_OUTSIDE_CAPTION As String = "Allow outside-month selection" 'Outside-month CheckBox caption
    Private Const DP_SETTINGS_CHECK_CLOSE_AFTER_CAPTION As String = "Close after selection" 'Close-after-selection CheckBox caption

    '-------------------------SETTINGS INTEGRATION PAGE------------------------
    Private Const DP_SETTINGS_CHECK_RIGHT_CLICK_CAPTION As String = "Right-click menu"   'Right-click CheckBox caption
    Private Const DP_SETTINGS_CHECK_IN_GRID_ICON_CAPTION As String = "In-grid icon"      'In-grid icon CheckBox caption
    Private Const DP_SETTINGS_CHECK_WINAPI_STYLE_CAPTION As String = "Use WinAPI styling" 'WinAPI styling CheckBox caption

    '-----------------------------DATE BOUNDS----------------------------------
    Private Const DP_YEAR_PANEL_MAX_START               As Long = 9988                   'DP_MAX_YEAR - DP_PICKER_ITEM_COUNT + 1

'------------------------------------------------------------------------------
' PRIVATE VARIABLES
'------------------------------------------------------------------------------

    '-----------------------------EVENT HOOKS----------------------------------
    Private mDayLabelHooks                              As Collection                   'Runtime day-label event hooks
    Private mHeaderLabelHooks                           As Collection                   'Runtime header-label event hooks
    Private mPickerPanelHooks                           As Collection                   'Runtime picker-panel event hooks
    Private mFooterLabelHooks                           As Collection                   'Runtime footer-label event hooks
    Private mSettingsPanelHooks                         As Collection                   'Runtime settings-panel event hooks

    '-----------------------------HOVER STATE----------------------------------
    Private mHoveredDayCellIndex                        As Long                         'Currently hovered day-cell index
    Private mHoveredHeaderLabelName                     As String                       'Currently hovered header label name
    Private mHoveredPickerItemIndex                     As Long                         'Currently hovered picker item index
    Private mHoveredFooterActionName                    As String                       'Currently hovered footer action
    Private mHoveredSettingsPanelLabelName              As String                       'Currently hovered settings-panel label name
    
    Private mDayCellIndexMap                            As Object                       'Scripting.Dictionary: LabelName -> DayIndex

    '-----------------------------CONTROL CACHE--------------------------------
    Private mDayTextLabels(1 To DP_DAY_LABEL_COUNT)         As MSForms.Label            'Cached day text labels
    Private mDayBackLabels(1 To DP_DAY_LABEL_COUNT)         As MSForms.Label            'Cached day background labels
    Private mPickerTextLabels(1 To DP_PICKER_ITEM_COUNT)    As MSForms.Label            'Cached picker text labels
    Private mPickerBackLabels(1 To DP_PICKER_ITEM_COUNT)    As MSForms.Label            'Cached picker background labels
    Private mLbl_Time                                       As MSForms.Label            'Cached footer Time value label

    '-----------------------------DATE CACHE-----------------------------------
    Private mDayCellDates(1 To DP_DAY_LABEL_COUNT)      As Date                         'Cached date for each day cell
    Private mDayCellHasDate(1 To DP_DAY_LABEL_COUNT)    As Boolean                      'True when cached day-cell date is available

    '-----------------------------DISPLAY STATE--------------------------------
    Private mDisplayYear                                As Long                         'Currently displayed year
    Private mDisplayMonth                               As Long                         'Currently displayed month

    '-----------------------------OVERLAY STATE--------------------------------
    Private mPickerPanelMode                            As Long                         'Current picker panel mode
    Private mYearPanelStart                             As Long                         'First visible year in year panel

    '-----------------------------LIFECYCLE STATE------------------------------
    Private mHasActivated                               As Boolean                      'True after one-time activation logic has run

    '-----------------------------KEYBOARD STATE-------------------------------
    Private mKeyboardDate                               As Date                         'Current keyboard navigation date
    Private mHasKeyboardDate                            As Boolean                      'True when keyboard date is initialized

    '-----------------------------FONT CACHE-----------------------------------
    Private mDayFontNormal                              As Object                       'Cached normal day font
    Private mDayFontWeekend                             As Object                       'Cached weekend day font


'
'------------------------------------------------------------------------------
'
'                               FORM LIFECYCLE
'
'------------------------------------------------------------------------------
'

Private Sub UserForm_Initialize()

'
'------------------------------------------------------------------------------
'                           INITIALIZE USERFORM
'------------------------------------------------------------------------------
' PURPOSE
'   Initializes the DatePicker UserForm runtime layout and controls
'
' WHY THIS EXISTS
'   The DatePicker form is mostly built at runtime. Each new UserForm instance
'   must resolve the current settings, format the shell, initialize the displayed
'   month/year, create runtime controls, populate the visible calendar grid, and
'   apply optional Windows-specific styling
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures DatePicker settings are available, formats the UserForm shell,
'   resolves the initial display date, initializes keyboard navigation state,
'   builds all runtime UI sections, populates the calendar grid, and applies
'   borderless window styling when enabled and supported
'
' ERROR POLICY
'   Raises a descriptive runtime error to the caller if initialization fails
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   UF_Form_Format
'   M_FormBridge_ConsumeInitialDate
'   UF_DisplayPeriod_Initialize
'   UF_KeyboardDate_Initialize
'   UF_Header_Build
'   UF_WeekdayRow_Build
'   UF_DayGrid_Build
'   UF_PickerPanel_Build
'   UF_Footer_Build
'   UF_DayGrid_Populate
'   M_Platform_ShouldUseWinAPI
'   M_Window_RemoveTitleBar
'
' NOTES
'   The display period must be initialized before header labels are created
'
'   Runtime controls are rebuilt for each form instance, while settings are
'   loaded only when not already available
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_DatePicker.UserForm_Initialize"

    Dim InitialDate     As Date                     'Initial date consumed from the form bridge or system date
    Dim WindowResult    As DP_WindowStyleResult     'Structured outcome of the native styling attempt

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Reset the one-time activation guard
        mHasActivated = False
    'Use the configured startup position before the first paint
        Me.StartUpPosition = DP_FORM_STARTUP_POSITION
    'Ensure saved DatePicker settings are available
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' FORMAT USERFORM
'------------------------------------------------------------------------------
    'Format the existing DatePicker UserForm instance
        UF_Form_Format

'------------------------------------------------------------------------------
' RESOLVE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Use the bridge-provided initial date when available
        If Not M_FormBridge_ConsumeInitialDate(InitialDate) Then InitialDate = VBA.Date
    'Initialize the displayed month and year
        UF_DisplayPeriod_Initialize InitialDate
    'Initialize keyboard navigation date
        UF_KeyboardDate_Initialize InitialDate

'------------------------------------------------------------------------------
' CREATE RUNTIME UI
'------------------------------------------------------------------------------
    'Create the header banner and its labels
        UF_Header_Build
    'Create the locale-dependent or fixed weekday header row
        UF_WeekdayRow_Build
    'Create the 6 x 7 day-label grid
        UF_DayGrid_Build
    'Create the hidden month/year picker panel
        UF_PickerPanel_Build
    'Create the footer banner and its labels
        UF_Footer_Build

'------------------------------------------------------------------------------
' POPULATE RUNTIME UI
'------------------------------------------------------------------------------
    'Populate the day grid for the initialized display period
        UF_DayGrid_Populate mDisplayYear, mDisplayMonth

'------------------------------------------------------------------------------
' APPLY WINDOW STYLE
'------------------------------------------------------------------------------
    'Apply borderless window styling only when enabled and supported, and capture
    'the structured outcome instead of discarding it
        If M_Platform_ShouldUseWinAPI Then
            WindowResult = M_Window_RemoveTitleBar(Me)
        End If
    'Act on the outcome. Applied, RolledBack and a non-attempt are all known
    'states and continue normally: the form is either borderless as requested or
    'still wearing its original native chrome. RecoveryRequired is the one state
    'that is neither, so the load fails rather than presenting a window whose
    'native style could not be applied or restored
        If WindowResult.RecoveryRequired Then
            'Record the actionable detail before unwinding
                Debug.Print PROC_NAME & _
                    " | Window recovery required | FailedStep=" & WindowResult.FailedStep & _
                    " | LastApiError=" & VBA.CStr(WindowResult.LastApiError)
            'Fail the load. Initialize cannot safely unload the instance it is
            'still constructing, so aborting here is what keeps the unknown-state
            'window from ever being shown
                Err.Raise vbObjectError + 640, PROC_NAME, _
                    "DatePicker window styling left the native window in no known " & _
                    "good state and the form was not loaded. FailedStep=" & _
                    WindowResult.FailedStep & "; LastApiError=" & _
                    VBA.CStr(WindowResult.LastApiError)
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
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub

Private Sub UserForm_Activate()

'
'------------------------------------------------------------------------------
'                           ACTIVATE USERFORM
'------------------------------------------------------------------------------
' PURPOSE
'   Runs one-time post-show initialization for the DatePicker UserForm
'
' WHY THIS EXISTS
'   Some UI operations are more reliable after the UserForm window has been
'   created and activated. The activation event provides the correct point for
'   final optional window styling, mouse-based positioning, and runtime behavior
'   that depends on the displayed form
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits immediately if post-show initialization has already run, applies
'   optional borderless window styling when enabled and supported, positions the
'   form close to the mouse when platform support is available, applies the
'   configured static/live clock mode, and marks post-show initialization as
'   complete
'
' ERROR POLICY
'   Raises a descriptive runtime error if post-show initialization fails
'
' DEPENDENCIES
'   M_Platform_ShouldUseWinAPI
'   M_Window_RemoveTitleBar
'   M_Window_MoveFormToMouse
'   M_Timer_ApplyClockMode
'
' NOTES
'   UserForm_Activate can fire more than once during a form lifetime. The
'   mHasActivated guard prevents repeated positioning, timer initialization, and
'   post-show work
'
'   Borderless styling is controlled by M_Platform_ShouldUseWinAPI
'
'   Mouse positioning is intentionally attempted separately. Disabling WinAPI
'   styling should not disable mouse-based positioning when native WinAPI calls
'   are otherwise available
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_DatePicker.UserForm_Activate"

    Dim WindowResult    As DP_WindowStyleResult     'Structured outcome of the native styling attempt

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Exit if one-time post-show initialization has already run
        If mHasActivated Then Exit Sub

'------------------------------------------------------------------------------
' APPLY OPTIONAL WINDOW STYLE
'------------------------------------------------------------------------------
    'Retry title-bar removal only when borderless styling is enabled and
    'supported, and capture the structured outcome instead of discarding it
        If M_Platform_ShouldUseWinAPI Then
            WindowResult = M_Window_RemoveTitleBar(Me)
        End If
    'Act on the outcome. Applied, RolledBack and a non-attempt are known states
    'and activation continues. RecoveryRequired means the native style is neither
    'applied nor restored, so the form is torn down rather than left visible and
    'interactive in a state nothing can describe
        If WindowResult.RecoveryRequired Then
            'Record the actionable detail before unwinding
                Debug.Print PROC_NAME & _
                    " | Window recovery required | FailedStep=" & WindowResult.FailedStep & _
                    " | LastApiError=" & VBA.CStr(WindowResult.LastApiError)
            'Consume the one-time activation guard first. A re-entrant Activate
            'raised during teardown exits at the guard above, so the unload below
            'cannot drive a recursive activate/unload loop
                mHasActivated = True
            'Suppress teardown errors: the window is already in an unknown state
            'and a failed unload must not replace that diagnostic with its own
                On Error Resume Next
            'Remove the form rather than continuing post-show initialization
                Unload Me
            'Clear any suppressed teardown error
                Err.Clear
            'Exit because this instance is terminating
                Exit Sub
        End If

'------------------------------------------------------------------------------
' POSITION USERFORM
'------------------------------------------------------------------------------
    'Move the form close to the current mouse position when platform support allows it
        M_Window_MoveFormToMouse _
            Me, _
            DP_FORM_MOUSE_OFFSET_XPX, _
            DP_FORM_MOUSE_OFFSET_YPX, _
            DP_FORM_CENTER_ON_MOUSE

'------------------------------------------------------------------------------
' APPLY CLOCK MODE
'------------------------------------------------------------------------------
    'Apply the configured static or live clock mode
        M_Timer_ApplyClockMode

'------------------------------------------------------------------------------
' FINALIZE ACTIVATION
'------------------------------------------------------------------------------
    'Mark one-time post-show initialization as complete
        mHasActivated = True

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

Private Sub UserForm_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer)

'
'------------------------------------------------------------------------------
'                           USERFORM KEY DOWN
'------------------------------------------------------------------------------
' PURPOSE
'   Handles DatePicker keyboard navigation and shortcuts
'
' WHY THIS EXISTS
'   Users should be able to navigate, select, and invoke common DatePicker
'   actions without using the mouse
'
' INPUTS
'   KeyCode
'     Pressed key code
'
'   Shift
'     Standard MSForms keyboard modifier state
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Supports arrow-key navigation, month/year navigation, first/last day of
'   displayed month, keyboard date selection, Today, Now, month/year picker
'   access, and controlled closing through Esc
'
'   When an overlay panel is visible, Esc hides the overlay first instead of
'   closing the full DatePicker form
'
' ERROR POLICY
'   Raises a descriptive runtime error if keyboard routing fails
'
' DEPENDENCIES
'   UF_KeyboardDate_MoveDays
'   UF_KeyboardDate_MoveMonths
'   UF_KeyboardDate_Set
'   M_DatePolicy_CanSelectDate
'   M_Picker_SelectDate
'   DP_Today
'   DP_Now
'   DP_Close
'   UF_PickerPanel_ShowMonths
'   UF_PickerPanel_ShowYears
'
' NOTES
'   Letter shortcuts are ignored when Ctrl or Alt is pressed, so the form does
'   not accidentally steal broader Excel or add-in shortcuts
'
'   If your runtime settings or picker frame names differ, update the local
'   SETTINGS_PANEL_NAME and PICKER_PANEL_NAME constants
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UserForm_KeyDown"
    
    Const SETTINGS_PANEL_NAME   As String = "Fra_Settings"              'Settings panel frame name
    Const PICKER_PANEL_NAME     As String = "Fra_PickerPanel"           'Month/year picker panel frame name
    Const MSFORMS_ALT_MASK      As Integer = 4                          'MSForms Alt modifier mask

    Dim CtrlIsPressed           As Boolean                              'True when Ctrl modifier is pressed
    Dim AltIsPressed            As Boolean                              'True when Alt modifier is pressed
    
    Dim LetterShortcutAllowed   As Boolean                              'True when letter shortcuts may be routed
    Dim ActionHandled           As Boolean                              'True when the pressed key was handled
    
    Dim SettingsPanel           As MSForms.control                      'Settings panel control reference
    Dim PickerPanel             As MSForms.control                      'Picker panel control reference
    Dim SettingsPanelVisible    As Boolean                              'True when settings panel is visible
    Dim PickerPanelVisible      As Boolean                              'True when picker panel is visible

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Resolve Ctrl key state
        CtrlIsPressed = ((Shift And DP_KEY_CTRL_MASK) <> 0)
    'Resolve Alt key state
        AltIsPressed = ((Shift And MSFORMS_ALT_MASK) <> 0)
    'Allow letter shortcuts only when Ctrl and Alt are not pressed
        LetterShortcutAllowed = Not CtrlIsPressed And Not AltIsPressed

'------------------------------------------------------------------------------
' RESOLVE OVERLAY PANELS
'------------------------------------------------------------------------------
    'Suppress missing-control errors while probing optional runtime panels
        On Error Resume Next
    'Resolve the settings panel reference
        Set SettingsPanel = Me.Controls(SETTINGS_PANEL_NAME)
    'Resolve the settings panel visibility state
        If Not SettingsPanel Is Nothing Then
            SettingsPanelVisible = CBool(SettingsPanel.Visible)
        End If
    'Resolve the picker panel reference
        Set PickerPanel = Me.Controls(PICKER_PANEL_NAME)
    'Resolve the picker panel visibility state
        If Not PickerPanel Is Nothing Then
            PickerPanelVisible = CBool(PickerPanel.Visible)
        End If
    'Clear any suppressed probing error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' ROUTE SETTINGS PANEL KEYS
'------------------------------------------------------------------------------
    'Let the settings panel own keyboard behavior while visible
        If SettingsPanelVisible Then
            'Route settings-panel keys
                Select Case KeyCode
                    Case vbKeyEscape
                        'Hide the settings panel and mark the key as handled
                            SettingsPanel.Visible = False: ActionHandled = True
                End Select
            'Consume the key when it was handled
                If ActionHandled Then KeyCode = 0
            'Exit because normal calendar shortcuts should not fire over settings
                Exit Sub
        End If

'------------------------------------------------------------------------------
' ROUTE PICKER PANEL KEYS
'------------------------------------------------------------------------------
    'Let the month/year picker panel own keyboard behavior while visible
        If PickerPanelVisible Then
            'Route picker-panel keys
                Select Case KeyCode
                    Case vbKeyEscape
                        'Hide the picker panel and mark the key as handled
                            PickerPanel.Visible = False: ActionHandled = True
                End Select
            'Consume the key when it was handled
                If ActionHandled Then KeyCode = 0
            'Exit because normal calendar shortcuts should not fire over the picker panel
                Exit Sub
        End If

'------------------------------------------------------------------------------
' ROUTE CALENDAR KEYS
'------------------------------------------------------------------------------
    'Route the pressed key
        Select Case KeyCode
        
            Case vbKeyEscape
                'Close the DatePicker
                    DP_Close
                    ActionHandled = True

            Case vbKeyReturn, vbKeySpace
                'Select the keyboard-highlighted date when available and selectable
                    If mHasKeyboardDate Then
                        If M_DatePolicy_CanSelectDate(mKeyboardDate, mDisplayYear, mDisplayMonth) Then
                            M_Picker_SelectDate mKeyboardDate
                        End If
                    End If
                'Mark the key as handled
                    ActionHandled = True

            Case vbKeyLeft
                'Move one day backward when Alt is not pressed
                    If Not AltIsPressed Then
                        UF_KeyboardDate_MoveDays -1
                        ActionHandled = True
                    End If

            Case vbKeyRight
                'Move one day forward when Alt is not pressed
                    If Not AltIsPressed Then
                        UF_KeyboardDate_MoveDays 1
                        ActionHandled = True
                    End If

            Case vbKeyUp
                'Move one week backward when Alt is not pressed
                    If Not AltIsPressed Then
                        UF_KeyboardDate_MoveDays -7
                        ActionHandled = True
                    End If

            Case vbKeyDown
                'Move one week forward when Alt is not pressed
                    If Not AltIsPressed Then
                        UF_KeyboardDate_MoveDays 7
                        ActionHandled = True
                    End If

            Case vbKeyPageUp
                'Move one year or one month backward when Alt is not pressed
                    If Not AltIsPressed Then
                        If CtrlIsPressed Then
                            UF_KeyboardDate_MoveMonths -12
                        Else
                            UF_KeyboardDate_MoveMonths -1
                        End If

                        ActionHandled = True
                    End If

            Case vbKeyPageDown
                'Move one year or one month forward when Alt is not pressed
                    If Not AltIsPressed Then
                        If CtrlIsPressed Then
                            UF_KeyboardDate_MoveMonths 12
                        Else
                            UF_KeyboardDate_MoveMonths 1
                        End If

                        ActionHandled = True
                    End If

            Case vbKeyHome
                'Move to the first day of the displayed month when Alt is not pressed
                    If Not AltIsPressed Then
                        UF_KeyboardDate_Set VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
                        ActionHandled = True
                    End If

            Case vbKeyEnd
                'Move to the last day of the displayed month when Alt is not pressed
                    If Not AltIsPressed Then
                        UF_KeyboardDate_Set VBA.DateSerial(mDisplayYear, mDisplayMonth + 1, 0)
                        ActionHandled = True
                    End If

            Case vbKeyM
                'Show the month picker panel when no conflicting modifier is pressed
                    If LetterShortcutAllowed Then
                        UF_PickerPanel_ShowMonths
                        ActionHandled = True
                    End If

            Case vbKeyY
                'Show the year picker panel when no conflicting modifier is pressed
                    If LetterShortcutAllowed Then
                        UF_PickerPanel_ShowYears
                        ActionHandled = True
                    End If

            Case vbKeyT
                'Write today without time when no conflicting modifier is pressed
                    If LetterShortcutAllowed Then
                        DP_Today
                        ActionHandled = True
                    End If

            Case vbKeyN
                'Write today with the current system time when no conflicting modifier is pressed
                    If LetterShortcutAllowed Then
                        DP_Now
                        ActionHandled = True
                    End If

        End Select

'------------------------------------------------------------------------------
' CONSUME HANDLED KEY
'------------------------------------------------------------------------------
    'Consume the key only when this routine handled it
        If ActionHandled Then KeyCode = 0

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

Private Sub UserForm_MouseMove( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

'
'------------------------------------------------------------------------------
'                           USERFORM MOUSE MOVE
'------------------------------------------------------------------------------
' PURPOSE
'   Removes active hover formatting when the mouse moves over the UserForm
'   background
'
' WHY THIS EXISTS
'   Runtime-created labels receive their own MouseMove event when the pointer is
'   over them. The UserForm receives MouseMove when the pointer leaves those
'   labels and moves over the form surface
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
'   Exits immediately when no hover state is active. Otherwise, resets only the
'   hover areas that currently have active hover state
'
' ERROR POLICY
'   Best-effort UI cleanup. Suppresses hover-reset failures because MouseMove is
'   a high-frequency visual event and should never interrupt the user workflow
'
' DEPENDENCIES
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_PickerPanel_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'
' NOTES
'   Guarding the reset calls avoids repeated no-op reset work while the mouse is
'   moving across the UserForm background
'
'   Button, Shift, X, and Y are required by the MSForms event signature and are
'   intentionally not used by this routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UserForm_MouseMove"

    Dim AnyHoverStateActive     As Boolean  'True when at least one hover state is active

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe

    'Resolve whether any hover state is active
        AnyHoverStateActive = _
            (mHoveredDayCellIndex <> 0) Or _
            (Len(mHoveredHeaderLabelName) <> 0) Or _
            (mHoveredPickerItemIndex <> 0) Or _
            (Len(mHoveredFooterActionName) <> 0) Or _
            (Len(mHoveredSettingsPanelLabelName) <> 0)

    'Exit immediately when there is no hover state to reset
        If Not AnyHoverStateActive Then Exit Sub

'------------------------------------------------------------------------------
' RESET HOVER STATE
'------------------------------------------------------------------------------
    'Reset day-cell hover only when a day cell is currently hovered
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Reset header hover only when a header label is currently hovered
        If Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Reset picker-panel hover only when a picker-panel item is currently hovered
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Reset footer hover only when a footer action is currently hovered
        If Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset
    'Reset settings-panel hover only when a settings label is currently hovered
        If Len(mHoveredSettingsPanelLabelName) <> 0 Then UF_SettingsPanel_HoverReset

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the fail-safe handler
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Suppress any secondary cleanup errors
        On Error Resume Next
    'Clear the day-cell hover tracker
        mHoveredDayCellIndex = 0
    'Clear the header hover tracker
        mHoveredHeaderLabelName = vbNullString
    'Clear the picker-panel hover tracker
        mHoveredPickerItemIndex = 0
    'Clear the footer hover tracker
        mHoveredFooterActionName = vbNullString
    'Clear the settings-panel hover tracker
        mHoveredSettingsPanelLabelName = vbNullString
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub UserForm_Terminate()

'
'------------------------------------------------------------------------------
'                           TERMINATE USERFORM
'------------------------------------------------------------------------------
' PURPOSE
'   Releases runtime references held by the DatePicker UserForm
'
' WHY THIS EXISTS
'   Runtime label event hooks, cached control references, cached font objects,
'   timer callbacks, hover trackers, and keyboard state are form-lifetime
'   resources. They should be released when the form is destroyed so that the
'   next DatePicker instance starts from a clean state
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stops the live clock timer, releases runtime event-hook collections, clears
'   cached control references, releases cached font objects, and resets transient
'   form state
'
' ERROR POLICY
'   Best-effort teardown. Suppresses cleanup errors because the UserForm is
'   already terminating and shutdown should never be interrupted
'
' DEPENDENCIES
'   M_Timer_Stop
'   UF_DayGrid_ClearCache
'   UF_PickerPanel_ClearCache
'
' NOTES
'   Visual hover reset is not required during termination because the UserForm is
'   being destroyed
'
'   This routine should not call DP_Close or unload the form again
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors during form teardown
        On Error Resume Next

'------------------------------------------------------------------------------
' STOP TIMER
'------------------------------------------------------------------------------
    'Stop the live clock timer before releasing form-level references
        M_Timer_Stop

'------------------------------------------------------------------------------
' RELEASE EVENT HOOKS
'------------------------------------------------------------------------------
    'Release day-label event hooks
        Set mDayLabelHooks = Nothing
    'Release header-label event hooks
        Set mHeaderLabelHooks = Nothing
    'Release picker-panel event hooks
        Set mPickerPanelHooks = Nothing
    'Release footer-label event hooks
        Set mFooterLabelHooks = Nothing
    'Release settings-panel event hooks
        Set mSettingsPanelHooks = Nothing

'------------------------------------------------------------------------------
' CLEAR CONTROL CACHES
'------------------------------------------------------------------------------
    'Clear cached day-label references
        UF_DayGrid_ClearCache
    'Clear cached picker-panel label references
        UF_PickerPanel_ClearCache

'------------------------------------------------------------------------------
' RELEASE CACHED OBJECTS
'------------------------------------------------------------------------------
    'Release cached normal day font
        Set mDayFontNormal = Nothing
    'Release cached weekend day font
        Set mDayFontWeekend = Nothing
    'Release cached footer Time value label
        Set mLbl_Time = Nothing
    'Release the day-cell index map
        Set mDayCellIndexMap = Nothing

'------------------------------------------------------------------------------
' CLEAR HOVER STATE
'------------------------------------------------------------------------------
    'Clear the current day-cell hover index
        mHoveredDayCellIndex = 0
    'Clear the current header-label hover state
        mHoveredHeaderLabelName = vbNullString
    'Clear the current picker-panel hover state
        mHoveredPickerItemIndex = 0
    'Clear the current footer hover state
        mHoveredFooterActionName = vbNullString
    'Clear the current settings-panel hover state
        mHoveredSettingsPanelLabelName = vbNullString

'------------------------------------------------------------------------------
' CLEAR PANEL STATE
'------------------------------------------------------------------------------
    'Clear the picker-panel mode
        mPickerPanelMode = 0
    'Clear the year-panel start
        mYearPanelStart = 0

'------------------------------------------------------------------------------
' CLEAR KEYBOARD STATE
'------------------------------------------------------------------------------
    'Clear keyboard navigation date
        mKeyboardDate = 0
    'Clear keyboard navigation availability
        mHasKeyboardDate = False

'------------------------------------------------------------------------------
' CLEAR ACTIVATION STATE
'------------------------------------------------------------------------------
    'Clear the activation guard
        mHasActivated = False

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                               USERFORM SHELL
'
'------------------------------------------------------------------------------
'

Private Sub UF_Form_Format()

'
'------------------------------------------------------------------------------
'                           FORMAT DATEPICKER FORM
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the standard visual formatting to the existing UF_DatePicker
'   UserForm instance
'
' WHY THIS EXISTS
'   The DatePicker form shell should be formatted consistently before runtime
'   controls are created or refreshed. Shell formatting is separated from final
'   sizing because borderless and native-title-bar modes require different outer
'   UserForm dimensions
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Applies standard shell properties and font settings, then applies the
'   effective form size according to the configured layout mode and WinAPI
'   styling setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if form formatting fails
'
' DEPENDENCIES
'   UF_Form_ApplyEffectiveSize
'   gDP_SizeMode
'
' NOTES
'   Runtime labels inherit the UserForm font if created after this routine
'
'   Constant validation is intentionally not repeated here. Form constants are
'   developer-owned configuration and should be checked by a dedicated debug or
'   regression routine, not by every UserForm initialization
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_DatePicker.UF_Form_Format"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' FORMAT USERFORM SHELL
'------------------------------------------------------------------------------
    'Apply the standard DatePicker shell properties
        With Me
            .Caption = DP_FORM_CAPTION
            .BackColor = DP_FORM_BACK_COLOR
            .ForeColor = DP_FORM_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .ScrollBars = fmScrollBarsNone
            .KeepScrollBarsVisible = fmScrollBarsNone
            .StartUpPosition = DP_FORM_STARTUP_POSITION
            .Cycle = fmCycleAllForms
            .MousePointer = fmMousePointerDefault
            .PictureAlignment = fmPictureAlignmentCenter
            .PictureSizeMode = fmPictureSizeModeClip
            .PictureTiling = False
            .RightToLeft = False
            .Zoom = 100
        End With

'------------------------------------------------------------------------------
' FORMAT USERFORM FONT
'------------------------------------------------------------------------------
    'Apply the standard DatePicker UserForm font
        With Me.Font
            .Name = DP_FORM_FONT_NAME
            .Size = DP_FORM_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

'------------------------------------------------------------------------------
' APPLY EFFECTIVE FORM SIZE
'------------------------------------------------------------------------------
    'Apply compact effective size when compact mode is enabled
        If gDP_SizeMode = DP_SizeMode_Compact Then
            UF_Form_ApplyEffectiveSize DP_FORM_HEIGHT_COMPACT
        Else
            UF_Form_ApplyEffectiveSize DP_FORM_HEIGHT
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
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub

Public Sub UF_Form_BeginDrag()

'
'------------------------------------------------------------------------------
'                           BEGIN FORM DRAG
'------------------------------------------------------------------------------
' PURPOSE
'   Starts UserForm movement from a custom DatePicker drag surface
'
' WHY THIS EXISTS
'   The DatePicker can run without a native title bar. A custom header surface
'   therefore needs a public owner-form method that can be called by the runtime
'   label hook when the user presses the left mouse button on a drag-enabled
'   label
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates native drag behavior to M_Window_BeginUserFormDrag
'
' ERROR POLICY
'   Best-effort UI behavior. Suppresses drag failures so form interaction is not
'   interrupted if the window handle cannot be resolved
'
' DEPENDENCIES
'   M_Window_BeginUserFormDrag
'
' NOTES
'   This method is intentionally Public because cDatePickerLabelHook routes to it
'   through CallByName
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress drag failures
        On Error Resume Next

'------------------------------------------------------------------------------
' BEGIN DRAG
'------------------------------------------------------------------------------
    'Start native UserForm drag movement
        M_Window_BeginUserFormDrag Me

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Clear any suppressed drag error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub UF_Form_ApplyEffectiveSize(ByVal TargetFormHeight As Single)

'
'------------------------------------------------------------------------------
'                         APPLY EFFECTIVE FORM SIZE
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the DatePicker UserForm outer size for the selected layout and
'   styling mode
'
' WHY THIS EXISTS
'   Borderless WinAPI mode and native-title-bar mode do not require exactly the
'   same outer UserForm dimensions. When the native title bar remains visible,
'   a small deterministic compensation is needed to preserve the intended visual
'   layout without relying on unstable InsideWidth / InsideHeight measurements
'
' INPUTS
'   TargetFormHeight
'     Intended DatePicker form height for the selected layout mode
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Applies the standard configured form width and requested height when WinAPI
'   styling is active. Applies a small native-chrome compensation when WinAPI
'   styling is disabled
'
' ERROR POLICY
'   Raises a descriptive runtime error if the requested target height is invalid
'
' DEPENDENCIES
'   M_Platform_ShouldUseWinAPI
'   DP_FORM_WIDTH
'   DP_FORM_NATIVE_WIDTH_COMPENSATION
'   DP_FORM_NATIVE_HEIGHT_COMPENSATION
'
' NOTES
'   This routine intentionally avoids dynamic InsideWidth / InsideHeight
'   compensation because those values can be unstable while the UserForm window
'   is being created or restyled
'
'   If native-title-bar mode feels too wide or too tall, tune only
'   DP_FORM_NATIVE_WIDTH_COMPENSATION and DP_FORM_NATIVE_HEIGHT_COMPENSATION
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_DatePicker.UF_Form_ApplyEffectiveSize"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject invalid requested height
        If TargetFormHeight <= 0 Then
            Err.Raise vbObjectError + 601, PROC_NAME, "TargetFormHeight must be greater than zero"
        End If

'------------------------------------------------------------------------------
' APPLY BORDERLESS SIZE
'------------------------------------------------------------------------------
    'Apply the standard borderless size when WinAPI styling is enabled and supported
        If M_Platform_ShouldUseWinAPI Then
            With Me
                .Width = DP_FORM_WIDTH
                .Height = TargetFormHeight
            End With
            'Exit because no native-title-bar compensation is required
                Exit Sub
        End If

'------------------------------------------------------------------------------
' APPLY NATIVE-CHROME SIZE
'------------------------------------------------------------------------------
    'Apply compensated size when the native title bar remains visible
        With Me
            .Width = DP_FORM_WIDTH + DP_FORM_NATIVE_WIDTH_COMPENSATION
            .Height = TargetFormHeight + DP_FORM_NATIVE_HEIGHT_COMPENSATION
        End With

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

'
'------------------------------------------------------------------------------
'
'                           DISPLAY AND KEYBOARD STATE
'
'------------------------------------------------------------------------------
'


Private Sub UF_DisplayPeriod_Initialize( _
    Optional ByVal InitialDate As Date = 0)

'
'------------------------------------------------------------------------------
'                           INITIALIZE DISPLAY PERIOD
'------------------------------------------------------------------------------
' PURPOSE
'   Initializes the DatePicker displayed month and year
'
' WHY THIS EXISTS
'   Header labels, keyboard navigation, picker-panel state, and the day grid
'   depend on mDisplayMonth and mDisplayYear being initialized before runtime UI
'   controls are built or refreshed
'
' INPUTS
'   InitialDate
'     Initial date used to set the displayed month and year
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Uses the current system date when InitialDate is zero, normalizes the
'   resolved date to a date-only value, validates the resulting display year and
'   renderable calendar boundary, then stores mDisplayYear and mDisplayMonth
'
' ERROR POLICY
'   Raises a descriptive runtime error if the resolved display period is outside
'   the supported DatePicker range or cannot render a complete calendar grid
'
' DEPENDENCIES
'   VBA.Date
'   VBA.DateValue
'   DP_MIN_YEAR
'   DP_MAX_YEAR
'
' NOTES
'   A zero date is treated as no explicit initial date
'
'   January 100 and December 9999 are rejected because the fixed calendar grid
'   cannot render the required adjacent-month dates outside the VBA Date
'   supported range
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_DisplayPeriod_Initialize"
    
    Const VBA_DATE_MIN_YEAR     As Long = 100           'Minimum year supported by VBA Date
    Const VBA_DATE_MAX_YEAR     As Long = 9999          'Maximum year supported by VBA Date

    Dim EffectiveDate           As Date                 'Resolved initial display date
    Dim EffectiveYear           As Long                 'Resolved display year
    Dim EffectiveMonth          As Long                 'Resolved display month

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE INITIAL DATE
'------------------------------------------------------------------------------
    'Use the current system date when no initial date is supplied
        If InitialDate = 0 Then
            EffectiveDate = VBA.Date
        Else
            EffectiveDate = VBA.DateValue(InitialDate)
        End If
    'Resolve the effective display year
        EffectiveYear = VBA.Year(EffectiveDate)
    'Resolve the effective display month
        EffectiveMonth = VBA.Month(EffectiveDate)

'------------------------------------------------------------------------------
' VALIDATE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Reject years outside the supported display range
        If EffectiveYear < DP_MIN_YEAR Or EffectiveYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "InitialDate year must be between " & VBA.CStr(DP_MIN_YEAR) & _
                " and " & VBA.CStr(DP_MAX_YEAR)
        End If
    'Reject the lower VBA Date rendering boundary
        If EffectiveYear = VBA_DATE_MIN_YEAR And EffectiveMonth = 1 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "January " & VBA.CStr(VBA_DATE_MIN_YEAR) & _
                " cannot render a complete DatePicker grid"
        End If
    'Reject the upper VBA Date rendering boundary
        If EffectiveYear = VBA_DATE_MAX_YEAR And EffectiveMonth = 12 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "December " & VBA.CStr(VBA_DATE_MAX_YEAR) & _
                " cannot render a complete DatePicker grid"
        End If

'------------------------------------------------------------------------------
' STORE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Store the displayed year
        mDisplayYear = EffectiveYear
    'Store the displayed month
        mDisplayMonth = EffectiveMonth

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
        Err.Raise Err.Number, PROC_NAME, "Display period initialization failed: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' KEYBOARD NAVIGATION STATE
'------------------------------------------------------------------------------

Private Sub UF_KeyboardDate_Initialize(ByVal InitialDate As Date)

'
'------------------------------------------------------------------------------
'                           INITIALIZE KEYBOARD DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Initializes the DatePicker keyboard-navigation date
'
' WHY THIS EXISTS
'   Keyboard navigation needs one internal date representing the currently
'   highlighted day before the user confirms the selection
'
' INPUTS
'   InitialDate
'     Initial date resolved when the form opens
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Uses the current selected date when available, otherwise uses InitialDate.
'   If InitialDate is zero, falls back to the current system date
'
' ERROR POLICY
'   Raises a descriptive runtime error if the keyboard date cannot be initialized
'
' DEPENDENCIES
'   gDP_HasSelectedDate
'   gDP_SelectedDate
'
' NOTES
'   This routine does not write to Excel
'
'   The keyboard date is stored as a date-only value
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_KeyboardDate_Initialize"

    Dim EffectiveDate       As Date 'Resolved keyboard navigation date

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE KEYBOARD DATE
'------------------------------------------------------------------------------
    'Use the selected date when one exists
        If gDP_HasSelectedDate Then
            'Resolve the selected date as the keyboard navigation date
                EffectiveDate = VBA.DateValue(gDP_SelectedDate)
        Else
            'Use the current system date when no explicit initial date is supplied
                If InitialDate = 0 Then
                    EffectiveDate = VBA.Date
                Else
                    EffectiveDate = VBA.DateValue(InitialDate)
                End If
        End If

'------------------------------------------------------------------------------
' STORE KEYBOARD STATE
'------------------------------------------------------------------------------
    'Store the keyboard navigation date
        mKeyboardDate = EffectiveDate
    'Mark keyboard navigation as initialized
        mHasKeyboardDate = True

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
        Err.Raise Err.Number, PROC_NAME, "Keyboard date initialization failed: " & Err.Description

End Sub

Private Sub UF_KeyboardDate_MoveDays(ByVal DayOffset As Long)

'
'------------------------------------------------------------------------------
'                           MOVE KEYBOARD DATE BY DAYS
'------------------------------------------------------------------------------
' PURPOSE
'   Moves the keyboard-selected DatePicker date by a number of days
'
' WHY THIS EXISTS
'   Arrow-key navigation moves within the calendar grid by day or by week
'
' INPUTS
'   DayOffset
'     Number of days to move
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures the keyboard date is initialized, applies the requested day offset,
'   and delegates display synchronization to UF_KeyboardDate_Set
'
' ERROR POLICY
'   Raises a descriptive runtime error if the movement cannot be applied
'
' DEPENDENCIES
'   UF_DisplayPeriod_Initialize
'   UF_KeyboardDate_Set
'
' NOTES
'   This routine does not write to Excel
'
'   A zero DayOffset is treated as a safe no-op
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_KeyboardDate_MoveDays"

    Dim EffectiveBaseDate   As Date     'Resolved starting keyboard date
    Dim NewKeyboardDate     As Date     'Resolved moved keyboard date

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' EXIT IF NO MOVEMENT IS REQUESTED
'------------------------------------------------------------------------------
    'Exit when no date movement is requested
        If DayOffset = 0 Then Exit Sub

'------------------------------------------------------------------------------
' ENSURE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Initialize the display period when it is not yet usable
        If mDisplayYear = 0 Or mDisplayMonth < 1 Or mDisplayMonth > 12 Then
            UF_DisplayPeriod_Initialize VBA.Date
        End If

'------------------------------------------------------------------------------
' ENSURE KEYBOARD DATE
'------------------------------------------------------------------------------
    'Initialize keyboard date to the first displayed day when missing
        If Not mHasKeyboardDate Then
            'Resolve the first day of the displayed month
                EffectiveBaseDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
            'Store the keyboard navigation date
                mKeyboardDate = EffectiveBaseDate
            'Mark keyboard navigation as initialized
                mHasKeyboardDate = True
        Else
            'Use the current keyboard date as the movement base
                EffectiveBaseDate = VBA.DateValue(mKeyboardDate)
        End If

'------------------------------------------------------------------------------
' MOVE DATE
'------------------------------------------------------------------------------
    'Resolve the moved keyboard date
        NewKeyboardDate = VBA.DateAdd("d", DayOffset, EffectiveBaseDate)
    'Apply the moved keyboard date and refresh visible state
        UF_KeyboardDate_Set NewKeyboardDate

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
        Err.Raise Err.Number, PROC_NAME, "Keyboard day movement failed: " & Err.Description

End Sub

Private Sub UF_KeyboardDate_MoveMonths(ByVal MonthOffset As Long)

'
'------------------------------------------------------------------------------
'                           MOVE KEYBOARD DATE BY MONTHS
'------------------------------------------------------------------------------
' PURPOSE
'   Moves the keyboard-selected DatePicker date by a number of months
'
' WHY THIS EXISTS
'   PageUp / PageDown keyboard navigation should move across months and years
'
' INPUTS
'   MonthOffset
'     Number of months to move
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures the keyboard date is initialized, applies the requested month offset,
'   and delegates display synchronization to UF_KeyboardDate_Set
'
' ERROR POLICY
'   Raises a descriptive runtime error if the movement cannot be applied
'
' DEPENDENCIES
'   UF_DisplayPeriod_Initialize
'   UF_KeyboardDate_Set
'
' NOTES
'   This routine does not write to Excel
'
'   VBA DateAdd handles month-end adjustment automatically
'
'   A zero MonthOffset is treated as a safe no-op
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_KeyboardDate_MoveMonths"

    Dim EffectiveBaseDate   As Date     'Resolved starting keyboard date
    Dim NewKeyboardDate     As Date     'Resolved moved keyboard date

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' EXIT IF NO MOVEMENT IS REQUESTED
'------------------------------------------------------------------------------
    'Exit when no date movement is requested
        If MonthOffset = 0 Then Exit Sub

'------------------------------------------------------------------------------
' ENSURE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Initialize the display period when it is not yet usable
        If mDisplayYear = 0 Or mDisplayMonth < 1 Or mDisplayMonth > 12 Then
            UF_DisplayPeriod_Initialize VBA.Date
        End If

'------------------------------------------------------------------------------
' ENSURE KEYBOARD DATE
'------------------------------------------------------------------------------
    'Initialize keyboard date to the first displayed day when missing
        If Not mHasKeyboardDate Then
            'Resolve the first day of the displayed month
                EffectiveBaseDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
            'Store the keyboard navigation date
                mKeyboardDate = EffectiveBaseDate
            'Mark keyboard navigation as initialized
                mHasKeyboardDate = True
        Else
            'Use the current keyboard date as the movement base
                EffectiveBaseDate = VBA.DateValue(mKeyboardDate)
        End If

'------------------------------------------------------------------------------
' MOVE DATE
'------------------------------------------------------------------------------
    'Resolve the moved keyboard date
        NewKeyboardDate = VBA.DateAdd("m", MonthOffset, EffectiveBaseDate)
    'Apply the moved keyboard date and refresh visible state
        UF_KeyboardDate_Set NewKeyboardDate

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
        Err.Raise Err.Number, PROC_NAME, "Keyboard month movement failed: " & Err.Description

End Sub

Private Sub UF_KeyboardDate_Set(ByVal NewKeyboardDate As Date)

'
'------------------------------------------------------------------------------
'                           SET KEYBOARD DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Sets the keyboard-selected DatePicker date and refreshes the minimum
'   necessary visual state
'
' WHY THIS EXISTS
'   Keyboard navigation can fire repeatedly. Rebuilding all day cells on every
'   arrow-key movement is unnecessary when the new keyboard date remains inside
'   the currently displayed month
'
' INPUTS
'   NewKeyboardDate
'     New date selected by keyboard navigation
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Captures the previous keyboard date, stores the new keyboard date, hides the
'   picker panel when present, applies a full grid refresh when the new date
'   moves to another displayed month, and otherwise refreshes only the previous
'   and new visible day cells
'
' ERROR POLICY
'   Raises a descriptive runtime error if NewKeyboardDate cannot be displayed or
'   the visual state cannot be refreshed
'
' DEPENDENCIES
'   UF_PickerPanel_HoverReset
'   UF_DayCell_RefreshVisibleDate
'   UF_DayGrid_Populate
'
' NOTES
'   This routine preserves the full-refresh behavior when navigation crosses a
'   month boundary
'
'   January 100 and December 9999 are rejected because the fixed calendar grid
'   cannot render the required adjacent-month dates outside the VBA Date
'   supported range
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_KeyboardDate_Set"
    
    Const VBA_DATE_MIN_YEAR     As Long = 100               'Minimum year supported by VBA Date
    Const VBA_DATE_MAX_YEAR     As Long = 9999              'Maximum year supported by VBA Date

    Dim ExistingControl         As MSForms.control          'Control resolved by picker-panel name
    Dim Fra_PickerPanel         As MSForms.Frame            'Reusable picker panel
    Dim PreviousDate            As Date                     'Previous keyboard-selected date
    Dim HasPreviousDate         As Boolean                  'True when previous keyboard date exists
    Dim NewDateOnly             As Date                     'New keyboard date without time
    Dim NewDisplayYear          As Long                     'New keyboard-date year
    Dim NewDisplayMonth         As Long                     'New keyboard-date month
    Dim NeedsFullRefresh        As Boolean                  'True when the displayed month must change
    Dim HandlerStep             As String                   'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' CAPTURE PREVIOUS STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture previous keyboard state"
    'Capture the previous keyboard date when available
        If mHasKeyboardDate Then
            PreviousDate = VBA.DateValue(mKeyboardDate)
            HasPreviousDate = True
        End If

'------------------------------------------------------------------------------
' NORMALIZE NEW KEYBOARD DATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize new keyboard date"

    'Normalize the new keyboard date
        NewDateOnly = VBA.DateValue(NewKeyboardDate)
    'Resolve the new display year
        NewDisplayYear = VBA.Year(NewDateOnly)
    'Resolve the new display month
        NewDisplayMonth = VBA.Month(NewDateOnly)

'------------------------------------------------------------------------------
' VALIDATE NEW KEYBOARD DATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate new keyboard date"

    'Reject years outside the supported DatePicker range
        If NewDisplayYear < DP_MIN_YEAR Or NewDisplayYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "NewKeyboardDate year must be between " & VBA.CStr(DP_MIN_YEAR) & _
                " and " & VBA.CStr(DP_MAX_YEAR)
        End If
    'Reject the lower VBA Date rendering boundary
        If NewDisplayYear = VBA_DATE_MIN_YEAR And NewDisplayMonth = 1 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "January " & VBA.CStr(VBA_DATE_MIN_YEAR) & _
                " cannot render a complete DatePicker grid"
        End If
    'Reject the upper VBA Date rendering boundary
        If NewDisplayYear = VBA_DATE_MAX_YEAR And NewDisplayMonth = 12 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "December " & VBA.CStr(VBA_DATE_MAX_YEAR) & _
                " cannot render a complete DatePicker grid"
        End If

'------------------------------------------------------------------------------
' RESOLVE REFRESH STRATEGY
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve refresh strategy"

    'Force a full refresh when the current display state is not usable
        If mDisplayYear = 0 Or mDisplayMonth < 1 Or mDisplayMonth > 12 Then
            NeedsFullRefresh = True
        Else
            'Resolve whether the keyboard date moved outside the displayed month
                NeedsFullRefresh = _
                    (NewDisplayYear <> mDisplayYear Or _
                     NewDisplayMonth <> mDisplayMonth)
        End If

'------------------------------------------------------------------------------
' STORE NEW KEYBOARD STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Store keyboard state"

    'Store the new keyboard date
        mKeyboardDate = NewDateOnly
    'Mark keyboard navigation as initialized
        mHasKeyboardDate = True

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Hide picker panel"

    'Clear active picker hover before hiding the picker panel
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Suppress lookup errors when the picker panel is not available
        On Error Resume Next
    'Retrieve the picker panel control when available
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Hide the picker panel when it exists and is a frame
        If Not ExistingControl Is Nothing Then
            If VBA.TypeName(ExistingControl) = "Frame" Then
                Set Fra_PickerPanel = ExistingControl
                Fra_PickerPanel.Visible = False
            End If
        End If
    'Clear picker-panel mode after hiding the panel
        mPickerPanelMode = 0
    'Clear picker-panel hover state after hiding the panel
        mHoveredPickerItemIndex = 0

'------------------------------------------------------------------------------
' APPLY FULL REFRESH WHEN MONTH CHANGES
'------------------------------------------------------------------------------
    'Apply a full grid refresh when the new keyboard date changes displayed month
        If NeedsFullRefresh Then
            'Track the current handler step
                HandlerStep = "Apply full grid refresh"
            'Store the displayed year
                mDisplayYear = NewDisplayYear
            'Store the displayed month
                mDisplayMonth = NewDisplayMonth
            'Refresh the full calendar grid
                UF_DayGrid_Populate mDisplayYear, mDisplayMonth
            'Exit after full refresh
                Exit Sub
        End If

'------------------------------------------------------------------------------
' APPLY MINIMAL REFRESH WHEN MONTH IS UNCHANGED
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply minimal day-cell refresh"

    'Refresh the previously highlighted visible day when it changed
        If HasPreviousDate Then
            If PreviousDate <> NewDateOnly Then
                UF_DayCell_RefreshVisibleDate PreviousDate
            End If
        End If
    'Refresh the newly highlighted visible day
        UF_DayCell_RefreshVisibleDate NewDateOnly

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
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Keyboard date update failed: " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                            CENTRAL ACTION ROUTING
'
'------------------------------------------------------------------------------
'

Public Sub UF_PickerPanel_HandleAction(ByVal ActionName As String)

'
'------------------------------------------------------------------------------
'                           HANDLE LABEL ACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Handles click actions raised by DatePicker header labels, day labels, footer
'   labels, settings labels, and picker-panel item labels
'
' WHY THIS EXISTS
'   Runtime-created labels route their Click events through
'   cDatePickerLabelHook. A central action router keeps dynamic UI controls
'   loosely coupled from the DatePicker business logic
'
' INPUTS
'   ActionName
'     Action name routed by cDatePickerLabelHook
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Routes header navigation, picker-panel display, footer shortcuts, settings
'   panel actions, day selection, month selection, and year selection
'
' ERROR POLICY
'   Raises a descriptive runtime error if ActionName is blank, unsupported, or
'   contains an invalid day-label / picker-item index
'
' DEPENDENCIES
'   UF_Header_MoveMonth
'   UF_Header_MoveYear
'   UF_PickerPanel_ShowMonths
'   UF_PickerPanel_ShowYears
'   UF_DayGrid_Populate
'   UF_PickerPanel_EnsureCache
'   M_DatePolicy_CanSelectDate
'   M_Picker_SelectDate
'   DP_Today
'   DP_Now
'   DP_Close
'   UF_Settings_Show
'   UF_SettingsPanel_Hide
'   UF_SettingsPanel_Save
'   UF_SettingsPanel_SelectPage
'
' NOTES
'   This routine is the central action router for runtime-created DatePicker
'   labels
'
'   The current procedure name is kept for compatibility with existing
'   cDatePickerLabelHook calls. A future semantic rename to
'   UF_LabelAction_Handle would be clearer
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_PickerPanel_HandleAction"
    
    Const PICKER_ITEM_PREFIX    As String = "PICKER_ITEM_"          'Picker item action prefix
    Const DAY_PICKED_PREFIX     As String = "DAY_PICKED_"           'Day label click action prefix
    Const PICKER_MODE_NONE      As Long = 0                         'No picker-panel mode
    Const PICKER_MODE_MONTHS    As Long = 1                         'Month picker-panel mode
    Const PICKER_MODE_YEARS     As Long = 2                         'Year picker-panel mode

    Dim RoutedAction            As String                           'Normalized routed action name
    Dim ItemIndex               As Long                             'Clicked item index
    Dim SelectedValue           As Long                             'Selected month or year value
    Dim RawItemIndex            As String                           'Raw item index text
    Dim RawTagValue             As String                           'Raw label Tag value
    Dim ExistingControl         As MSForms.control                  'Existing control using the picker-panel name
    Dim Fra_PickerPanel         As MSForms.Frame                    'Reusable picker panel
    Dim SelectedDate            As Date                             'Selected date from clicked day label
    Dim Lbl_SelectedItem        As MSForms.Label                    'Selected picker item text label
    Dim HandlerStep             As String                           'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Track the current handler step
        HandlerStep = "Normalize action"
    'Normalize the routed action name
        RoutedAction = VBA.UCase$(VBA.Trim$(ActionName))
    'Reject an empty action name
        If VBA.Len(RoutedAction) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ActionName cannot be empty"
        End If

'------------------------------------------------------------------------------
' ROUTE DIRECT ACTIONS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Route direct action"
    'Route the requested direct action
        Select Case RoutedAction

            Case "PREV_MONTH"
                'Track the current handler step
                    HandlerStep = "Move previous month"
                'Move to the previous displayed month
                    UF_Header_MoveMonth -1
                'Exit after routing the action
                    Exit Sub

            Case "NEXT_MONTH"
                'Track the current handler step
                    HandlerStep = "Move next month"
                'Move to the next displayed month
                    UF_Header_MoveMonth 1
                'Exit after routing the action
                    Exit Sub

            Case "PREV_YEAR"
                'Track the current handler step
                    HandlerStep = "Move previous year"
                'Move or scroll to the previous year range
                    UF_Header_MoveYear -1
                'Exit after routing the action
                    Exit Sub

            Case "NEXT_YEAR"
                'Track the current handler step
                    HandlerStep = "Move next year"
                'Move or scroll to the next year range
                    UF_Header_MoveYear 1
                'Exit after routing the action
                    Exit Sub

            Case "SHOW_MONTH_PANEL"
                'Track the current handler step
                    HandlerStep = "Show month picker panel"
                'Show the month picker panel
                    UF_PickerPanel_ShowMonths
                'Exit after routing the action
                    Exit Sub

            Case "SHOW_YEAR_PANEL"
                'Track the current handler step
                    HandlerStep = "Show year picker panel"
                'Show the year picker panel
                    UF_PickerPanel_ShowYears
                'Exit after routing the action
                    Exit Sub

            Case "WRITE_NOW"
                'Track the current handler step
                    HandlerStep = "Write now"
                'Write today with the current system time
                    DP_Now
                'Exit after routing the action
                    Exit Sub

            Case "WRITE_TODAY"
                'Track the current handler step
                    HandlerStep = "Write today"
                'Write today without time
                    DP_Today
                'Exit after routing the action
                    Exit Sub

            Case "SHOW_SETTINGS"
                'Track the current handler step
                    HandlerStep = "Show settings panel"
                'Show the settings panel
                    UF_Settings_Show
                'Exit after routing the action
                    Exit Sub

            Case "HIDE_SETTINGS"
                'Track the current handler step
                    HandlerStep = "Hide settings panel"
                'Hide the settings panel
                    UF_SettingsPanel_Hide
                'Exit after routing the action
                    Exit Sub

            Case "SAVE_SETTINGS"
                'Track the current handler step
                    HandlerStep = "Save settings"
                'Save settings from the settings panel
                    UF_SettingsPanel_Save
                'Exit after routing the action
                    Exit Sub

            Case "SHOW_SETTINGS_DISPLAY"
                'Track the current handler step
                    HandlerStep = "Show Display Settings page"
                'Show the Display Settings page
                    UF_SettingsPanel_SelectPage 0
                'Exit after routing the action
                    Exit Sub

            Case "SHOW_SETTINGS_BEHAVIOR"
                'Track the current handler step
                    HandlerStep = "Show Behavior Settings page"
                'Show the Behavior Settings page
                    UF_SettingsPanel_SelectPage 1
                'Exit after routing the action
                    Exit Sub

            Case "SHOW_SETTINGS_INTEGRATION"
                'Track the current handler step
                    HandlerStep = "Show Integration Settings page"
                'Show the Integration Settings page
                    UF_SettingsPanel_SelectPage 2
                'Exit after routing the action
                    Exit Sub

            Case "CLOSE_PICKER", "CLOSE"
                'Track the current handler step
                    HandlerStep = "Close picker"
                'Close the DatePicker
                    DP_Close
                'Exit after routing the action
                    Exit Sub

        End Select

'------------------------------------------------------------------------------
' HANDLE DAY LABEL CLICK
'------------------------------------------------------------------------------
    'Handle day-label click actions
        If VBA.Left$(RoutedAction, VBA.Len(DAY_PICKED_PREFIX)) = DAY_PICKED_PREFIX Then
            'Track the current handler step
                HandlerStep = "Extract day index"
            'Extract the raw day-cell index
                RawItemIndex = VBA.Mid$(RoutedAction, VBA.Len(DAY_PICKED_PREFIX) + 1)
            'Reject empty or non-digit day-cell indexes
                If VBA.Len(RawItemIndex) = 0 Or RawItemIndex Like "*[!0-9]*" Then
                    Err.Raise vbObjectError + 514, PROC_NAME, "Day cell index must be numeric"
                End If
            'Track the current handler step
                HandlerStep = "Parse day index"
            'Parse the clicked day-cell index
                ItemIndex = VBA.CLng(RawItemIndex)
            'Reject invalid day-cell indexes
                If ItemIndex < 1 Or ItemIndex > DP_DAY_LABEL_COUNT Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Day cell index must be between 1 and " & VBA.CStr(DP_DAY_LABEL_COUNT)
                End If
            'Track the current handler step
                HandlerStep = "Validate cached day-cell date"
            'Reject day cells with no cached date
                If Not mDayCellHasDate(ItemIndex) Then
                    Err.Raise vbObjectError + 516, PROC_NAME, _
                        "Clicked day cell does not have a cached date"
                End If
            'Resolve the selected date from the cached day-cell date
                SelectedDate = VBA.DateValue(mDayCellDates(ItemIndex))
            'Track the current handler step
                HandlerStep = "Validate selectable date"
            'Accept the selected date only if it is selectable
                If M_DatePolicy_CanSelectDate(SelectedDate, mDisplayYear, mDisplayMonth) Then
                    'Track the current handler step
                        HandlerStep = "Select date"
                    'Delegate write-back to the companion module
                        M_Picker_SelectDate SelectedDate
                End If
            'Exit after handling the day click
                Exit Sub
        End If

'------------------------------------------------------------------------------
' HANDLE PICKER ITEM CLICK
'------------------------------------------------------------------------------
    'Reject unsupported non-picker actions
        If VBA.Left$(RoutedAction, VBA.Len(PICKER_ITEM_PREFIX)) <> PICKER_ITEM_PREFIX Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Unsupported DatePicker action: " & RoutedAction
        End If
    'Track the current handler step
        HandlerStep = "Extract picker item index"
    'Extract the raw picker-item index
        RawItemIndex = VBA.Mid$(RoutedAction, VBA.Len(PICKER_ITEM_PREFIX) + 1)
    'Reject empty or non-digit picker-item indexes
        If VBA.Len(RawItemIndex) = 0 Or RawItemIndex Like "*[!0-9]*" Then
            Err.Raise vbObjectError + 518, PROC_NAME, "Picker item index must be numeric"
        End If
    'Track the current handler step
        HandlerStep = "Parse picker item index"
    'Parse the picker-item index
        ItemIndex = VBA.CLng(RawItemIndex)
    'Reject invalid picker item indexes
        If ItemIndex < 1 Or ItemIndex > DP_PICKER_ITEM_COUNT Then
            Err.Raise vbObjectError + 519, PROC_NAME, _
                "Picker item index must be between 1 and " & VBA.CStr(DP_PICKER_ITEM_COUNT)
        End If

'------------------------------------------------------------------------------
' RETRIEVE PICKER PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve picker panel"
    
    'Suppress lookup errors while resolving the picker panel
        On Error Resume Next
    'Retrieve the existing picker panel control
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject missing picker panel
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 520, PROC_NAME, _
                "Unable to resolve expected picker panel " & DP_PICKER_PANEL_NAME
        End If
    'Reject a name collision with a non-frame control
        If VBA.TypeName(ExistingControl) <> "Frame" Then
            Err.Raise vbObjectError + 521, PROC_NAME, _
                "Control '" & DP_PICKER_PANEL_NAME & "' exists but is not an MSForms.Frame"
        End If
    'Use the resolved picker panel frame
        Set Fra_PickerPanel = ExistingControl

'------------------------------------------------------------------------------
' ROUTE PICKER MODE
'------------------------------------------------------------------------------
    'Route the picker item according to the active picker mode
        Select Case mPickerPanelMode

            Case PICKER_MODE_MONTHS

'------------------------------------------------------------------------------
' APPLY MONTH SELECTION
'------------------------------------------------------------------------------
                'Track the current handler step
                    HandlerStep = "Apply month selection"
                'Store the selected month
                    SelectedValue = ItemIndex
                'Reject unsupported selected months
                    If SelectedValue < 1 Or SelectedValue > 12 Then
                        Err.Raise vbObjectError + 522, PROC_NAME, _
                            "Selected month must be between 1 and 12"
                    End If
                'Reject invalid display year state
                    If mDisplayYear < DP_MIN_YEAR Or mDisplayYear > DP_MAX_YEAR Then
                        Err.Raise vbObjectError + 523, PROC_NAME, _
                            "mDisplayYear is outside the supported DatePicker range"
                    End If
                'Store the selected month
                    mDisplayMonth = SelectedValue
                'Initialize keyboard date to the first day of the selected display month
                    mKeyboardDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
                'Mark keyboard navigation date as available
                    mHasKeyboardDate = True
                'Hide the picker panel
                    Fra_PickerPanel.Visible = False
                'Clear picker-panel mode
                    mPickerPanelMode = PICKER_MODE_NONE
                'Clear picker-panel hover state
                    mHoveredPickerItemIndex = 0
                'Refresh the calendar grid
                    UF_DayGrid_Populate mDisplayYear, mDisplayMonth
                'Exit after applying month selection
                    Exit Sub

            Case PICKER_MODE_YEARS

'------------------------------------------------------------------------------
' APPLY YEAR SELECTION
'------------------------------------------------------------------------------
                'Track the current handler step
                    HandlerStep = "Ensure picker cache"
                'Ensure cached picker-panel references are available
                    UF_PickerPanel_EnsureCache
                'Retrieve the selected picker text label
                    Set Lbl_SelectedItem = mPickerTextLabels(ItemIndex)
                'Reject a missing selected picker item label
                    If Lbl_SelectedItem Is Nothing Then
                        Err.Raise vbObjectError + 524, PROC_NAME, _
                            "Cached picker text label is missing for item " & VBA.CStr(ItemIndex)
                    End If
                'Track the current handler step
                    HandlerStep = "Read selected year tag"
                'Read the selected year value from the cached picker item Tag
                    RawTagValue = VBA.Trim$(VBA.CStr(Lbl_SelectedItem.Tag))
                'Reject an empty year picker Tag
                    If VBA.Len(RawTagValue) = 0 Then
                        Err.Raise vbObjectError + 525, PROC_NAME, _
                            "Selected year picker item does not contain a value"
                    End If
                'Reject non-digit year picker Tag values
                    If RawTagValue Like "*[!0-9]*" Then
                        Err.Raise vbObjectError + 526, PROC_NAME, _
                            "Selected year picker item Tag must contain a numeric year"
                    End If
                'Track the current handler step
                    HandlerStep = "Parse selected year"
                'Parse the selected year
                    SelectedValue = VBA.CLng(RawTagValue)
                'Reject unsupported selected years
                    If SelectedValue < DP_MIN_YEAR Or SelectedValue > DP_MAX_YEAR Then
                        Err.Raise vbObjectError + 527, PROC_NAME, _
                            "Selected year must be between " & VBA.CStr(DP_MIN_YEAR) & " and " & VBA.CStr(DP_MAX_YEAR)
                    End If
                'Reject invalid display month state
                    If mDisplayMonth < 1 Or mDisplayMonth > 12 Then
                        Err.Raise vbObjectError + 528, PROC_NAME, _
                            "mDisplayMonth must be between 1 and 12 before applying year selection"
                    End If
                'Track the current handler step
                    HandlerStep = "Apply year selection"
                'Store the selected year
                    mDisplayYear = SelectedValue
                'Initialize keyboard date to the first day of the selected display period
                    mKeyboardDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
                'Mark keyboard navigation date as available
                    mHasKeyboardDate = True
                'Hide the picker panel
                    Fra_PickerPanel.Visible = False
                'Clear picker-panel mode
                    mPickerPanelMode = PICKER_MODE_NONE
                'Clear picker-panel hover state
                    mHoveredPickerItemIndex = 0
                'Refresh the calendar grid
                    UF_DayGrid_Populate mDisplayYear, mDisplayMonth
                'Exit after applying year selection
                    Exit Sub

            Case Else
                'Reject unsupported picker panel modes
                    Err.Raise vbObjectError + 529, PROC_NAME, _
                        "mPickerPanelMode must be 1 for months or 2 for years"

        End Select

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
        Err.Raise Err.Number, _
            PROC_NAME & " | Action=" & RoutedAction & " | Step=" & HandlerStep, _
            "DatePicker action routing failed: " & Err.Description

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HEADER
'
'------------------------------------------------------------------------------
'

Private Sub UF_Header_Build()

'
'------------------------------------------------------------------------------
'                           CREATE HEADER SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Creates the DatePicker header banner and related header controls
'
' WHY THIS EXISTS
'   The DatePicker header groups the controls used to display and navigate the
'   current month and year. In compact layout, it also exposes the settings
'   entry point when the footer is not visible
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds or refreshes the header banner first, then builds or refreshes the
'   clickable header labels
'
' ERROR POLICY
'   Raises a descriptive runtime error if header creation fails
'
' DEPENDENCIES
'   UF_Header_BuildBanner
'   UF_Header_BuildLabels
'
' NOTES
'   This routine coordinates header creation only
'
'   The display period must already be initialized before header labels are
'   built, because month and year captions depend on mDisplayMonth and
'   mDisplayYear
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_DatePicker.UF_Header_Build"
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CREATE HEADER BANNER
'------------------------------------------------------------------------------
    'Create or format the header background banner
        UF_Header_BuildBanner

'------------------------------------------------------------------------------
' CREATE HEADER LABELS
'------------------------------------------------------------------------------
    'Create or format the header labels
        UF_Header_BuildLabels

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
        Err.Raise Err.Number, PROC_NAME, "Header creation failed. " & Err.Description

End Sub

Private Sub UF_Header_BuildBanner()

'
'------------------------------------------------------------------------------
'                           CREATE HEADER BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Creates and formats the DatePicker header background banner
'
' WHY THIS EXISTS
'   The header banner visually groups the month and year display controls into a
'   distinct top section and provides the visual background for the header
'   navigation labels
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses Lbl_HeaderBanner, formats it as the top header band, and
'   sends it behind the clickable header labels
'
' ERROR POLICY
'   Raises a descriptive runtime error if the banner cannot be created or
'   formatted
'
' DEPENDENCIES
'   UF_Ensure_Label
'
' NOTES
'   The banner is sized to the DatePicker designed canvas width. Native-title-bar
'   compensation is handled by UF_Form_ApplyEffectiveSize and should not be
'   duplicated here
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Header_BuildBanner"

    Dim Lbl_HeaderBanner    As MSForms.Label        'Header banner control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CREATE / FORMAT BANNER
'------------------------------------------------------------------------------
    'Create or retrieve the header banner label
        Set Lbl_HeaderBanner = UF_Ensure_Label("Lbl_HeaderBanner")

    'Apply layout and visual properties
        With Lbl_HeaderBanner
            .Caption = vbNullString
            .Left = 0
            .Top = DP_HEADER_TOP
            .Width = DP_FORM_WIDTH
            .Height = DP_HEADER_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_HEADER_BACK_COLOR
            .ForeColor = DP_HEADER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = True
            .Visible = True
        End With

    'Send the banner behind all clickable header labels
        Lbl_HeaderBanner.ZOrder 1

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
        Err.Raise Err.Number, PROC_NAME, "Header banner creation failed: " & Err.Description

End Sub

Private Sub UF_Header_BuildLabels()

'
'------------------------------------------------------------------------------
'                           CREATE HEADER LABELS
'------------------------------------------------------------------------------
' PURPOSE
'   Creates and formats the clickable labels displayed on the DatePicker header
'   banner
'
' WHY THIS EXISTS
'   The DatePicker header exposes the visible month/year state and the related
'   navigation controls used to move across months and years, open the reusable
'   picker panel, and expose settings access in compact layout
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses:
'     - Lbl_HeaderMonth
'     - Lbl_HeaderYear
'     - Lbl_PrevMonth
'     - Lbl_NextMonth
'     - Lbl_PrevYear
'     - Lbl_NextYear
'     - Lbl_HeaderSettings
'
'   Formats each label and registers its click / hover hook in
'   mHeaderLabelHooks
'
'   Assigns clean StdFont objects to all header labels so reused runtime labels
'   do not retain stale Italic, Underline, or Strikethrough states
'
' ERROR POLICY
'   Raises a descriptive runtime error if one or more labels cannot be created,
'   formatted, or hooked
'
' DEPENDENCIES
'   UF_Ensure_Label
'   cDatePickerLabelHook
'   M_Caption_GetMonth
'   UF_DisplayPeriod_Initialize
'   mDisplayMonth
'   mDisplayYear
'   gDP_UseLocalNames
'   gDP_SizeMode
'
' NOTES
'   mHeaderLabelHooks is reset each time the header labels are refreshed so that
'   runtime label events remain connected to the current controls
'
'   Header label actions are routed by cDatePickerLabelHook through the form
'   action router
'
'   Runtime MSForms labels may retain old font states when reused. This routine
'   therefore assigns clean StdFont objects instead of changing only selected
'   Font properties
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DatePicker.UF_Header_BuildLabels"   'Current procedure name

    Dim LabelHook              As cDatePickerLabelHook                            'Runtime click / hover hook

    Dim HeaderFont             As Object                                          'Clean header-label font object
    Dim HeaderIconFont         As Object                                          'Clean header icon font object

    Dim BannerLeft             As Single                                          'Header banner left position

    Dim Lbl_HeaderMonth        As MSForms.Label                                   'Header month label
    Dim Lbl_HeaderYear         As MSForms.Label                                   'Header year label
    Dim Lbl_PrevMonth          As MSForms.Label                                   'Header previous-month label
    Dim Lbl_NextMonth          As MSForms.Label                                   'Header next-month label
    Dim Lbl_PrevYear           As MSForms.Label                                   'Header previous-year label
    Dim Lbl_NextYear           As MSForms.Label                                   'Header next-year label
    Dim Lbl_HeaderSettings     As MSForms.Label                                   'Compact header settings label
    Dim Lbl_HeaderBanner       As MSForms.Label                                   'Header banner drag surface

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Reset header label click / hover hooks
        Set mHeaderLabelHooks = New Collection
    'Clear any stale header hover tracker before rebuilding labels
        mHoveredHeaderLabelName = vbNullString

'------------------------------------------------------------------------------
' REGISTER HEADER BANNER DRAG SURFACE
'------------------------------------------------------------------------------
    'Retrieve the header banner created by UF_Header_BuildBanner
        Set Lbl_HeaderBanner = Me.Controls("Lbl_HeaderBanner")
    'Create the header banner drag hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the header banner to form drag behavior only
        LabelHook.Initialize _
            OwnerForm:=Me, _
            TargetLabel:=Lbl_HeaderBanner, _
            ActionName:=vbNullString, _
            HoverMode:="NONE", _
            DragEnabled:=True
    'Store the hook so that the MouseDown event remains alive
        mHeaderLabelHooks.Add LabelHook, Lbl_HeaderBanner.Name

'------------------------------------------------------------------------------
' ENSURE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Initialize the display period if it is not already usable
        If mDisplayYear = 0 Or mDisplayMonth < 1 Or mDisplayMonth > 12 Then
            UF_DisplayPeriod_Initialize VBA.Date
        End If

'------------------------------------------------------------------------------
' CALCULATE LAYOUT
'------------------------------------------------------------------------------
    'Calculate the header banner left position
        BannerLeft = DP_DAY_GRID_START_LEFT + _
            ((DP_DAY_LABEL_WIDTH - DP_DOW_LABEL_WIDTH) / 2)

'------------------------------------------------------------------------------
' CREATE / FORMAT MONTH LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the header month label
        Set Lbl_HeaderMonth = UF_Ensure_Label("Lbl_HeaderMonth")
    'Apply layout and visual properties
        With Lbl_HeaderMonth
            .Caption = M_Caption_GetMonth(mDisplayMonth, gDP_UseLocalNames)
            .Left = BannerLeft + DP_HEADER_MONTH_REL_LEFT
            .Top = DP_HEADER_TOP + DP_HEADER_MONTH_TOP
            .Width = DP_HEADER_MONTH_WIDTH
            .Height = DP_HEADER_MONTH_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_HEADER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Open Month Picker"
        End With
    'Create a clean month-label font object
        Set HeaderFont = CreateObject("StdFont")
    'Configure the month-label font explicitly
        With HeaderFont
            .Name = DP_FORM_FONT_NAME
            .Size = 12
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With
    'Assign the clean font to the month label
        Set Lbl_HeaderMonth.Font = HeaderFont
    'Move the month label to the front
        Lbl_HeaderMonth.ZOrder 0
    'Create the month label click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the month label to the month-panel action
        LabelHook.Initialize Me, Lbl_HeaderMonth, "SHOW_MONTH_PANEL", "HEADER"
    'Store the hook so that the click / hover events remain alive
        mHeaderLabelHooks.Add LabelHook, Lbl_HeaderMonth.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT YEAR LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the header year label
        Set Lbl_HeaderYear = UF_Ensure_Label("Lbl_HeaderYear")

    'Apply layout and visual properties
        With Lbl_HeaderYear
            .Caption = CStr(mDisplayYear)
            .Left = BannerLeft + DP_HEADER_YEAR_REL_LEFT
            .Top = DP_HEADER_TOP + DP_HEADER_YEAR_TOP
            .Width = DP_HEADER_YEAR_WIDTH
            .Height = DP_HEADER_YEAR_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_HEADER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Open Year Picker"
        End With

    'Create a clean year-label font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the year-label font explicitly
        With HeaderFont
            .Name = DP_FORM_FONT_NAME
            .Size = 12
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean font to the year label
        Set Lbl_HeaderYear.Font = HeaderFont

    'Move the year label to the front
        Lbl_HeaderYear.ZOrder 0

    'Create the year label click / hover hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the year label to the year-panel action
        LabelHook.Initialize Me, Lbl_HeaderYear, "SHOW_YEAR_PANEL", "HEADER"

    'Store the hook so that the click / hover events remain alive
        mHeaderLabelHooks.Add LabelHook, Lbl_HeaderYear.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT PREVIOUS YEAR LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the previous-year label
        Set Lbl_PrevYear = UF_Ensure_Label("Lbl_PrevYear")

    'Apply layout and visual properties
        With Lbl_PrevYear
            .Caption = ChrW$(&H25B2)
            .Left = BannerLeft + DP_HEADER_YEAR_ARROW_LEFT
            .Top = DP_HEADER_TOP + DP_HEADER_YEAR_ARROW_TOP
            .Width = DP_HEADER_YEAR_ARROW_WIDTH
            .Height = DP_HEADER_YEAR_ARROW_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_HEADER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Previous Year"
        End With

    'Create a clean previous-year font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the previous-year font explicitly
        With HeaderFont
            .Name = DP_FORM_FONT_NAME
            .Size = 7
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean font to the previous-year label
        Set Lbl_PrevYear.Font = HeaderFont

    'Move the previous-year label to the front
        Lbl_PrevYear.ZOrder 0

    'Create the previous-year click / hover hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the previous-year label to the previous-year action
        LabelHook.Initialize Me, Lbl_PrevYear, "PREV_YEAR", "HEADER"

    'Store the hook so that the click / hover events remain alive
        mHeaderLabelHooks.Add LabelHook, Lbl_PrevYear.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT NEXT YEAR LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the next-year label
        Set Lbl_NextYear = UF_Ensure_Label("Lbl_NextYear")

    'Apply layout and visual properties
        With Lbl_NextYear
            .Caption = ChrW$(&H25BC)
            .Left = BannerLeft + DP_HEADER_YEAR_ARROW_LEFT
            .Top = DP_HEADER_TOP + DP_HEADER_YEAR_ARROW_TOP + 12
            .Width = DP_HEADER_YEAR_ARROW_WIDTH
            .Height = DP_HEADER_YEAR_ARROW_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_HEADER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Next Year"
        End With

    'Create a clean next-year font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the next-year font explicitly
        With HeaderFont
            .Name = DP_FORM_FONT_NAME
            .Size = 7
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean font to the next-year label
        Set Lbl_NextYear.Font = HeaderFont

    'Move the next-year label to the front
        Lbl_NextYear.ZOrder 0

    'Create the next-year click / hover hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the next-year label to the next-year action
        LabelHook.Initialize Me, Lbl_NextYear, "NEXT_YEAR", "HEADER"

    'Store the hook so that the click / hover events remain alive
        mHeaderLabelHooks.Add LabelHook, Lbl_NextYear.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT PREVIOUS MONTH LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the previous-month label
        Set Lbl_PrevMonth = UF_Ensure_Label("Lbl_PrevMonth")

    'Apply layout and visual properties
        With Lbl_PrevMonth
            .Caption = ChrW$(&H25C0)
            .Left = BannerLeft + DP_HEADER_PREVMONTH_ARROW_LEFT
            .Top = DP_HEADER_TOP + DP_HEADER_MONTH_ARROW_TOP + 2
            .Width = DP_HEADER_MONTH_ARROW_WIDTH
            .Height = DP_HEADER_MONTH_ARROW_HEIGHT - 5
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_HEADER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Previous Month"
        End With

    'Create a clean previous-month font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the previous-month font explicitly
        With HeaderFont
            .Name = DP_FORM_FONT_NAME
            .Size = 8
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean font to the previous-month label
        Set Lbl_PrevMonth.Font = HeaderFont

    'Move the previous-month label to the front
        Lbl_PrevMonth.ZOrder 0

    'Create the previous-month click / hover hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the previous-month label to the previous-month action
        LabelHook.Initialize Me, Lbl_PrevMonth, "PREV_MONTH", "HEADER"

    'Store the hook so that the click / hover events remain alive
        mHeaderLabelHooks.Add LabelHook, Lbl_PrevMonth.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT NEXT MONTH LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the next-month label
        Set Lbl_NextMonth = UF_Ensure_Label("Lbl_NextMonth")

    'Apply layout and visual properties
        With Lbl_NextMonth
            .Caption = ChrW$(&H25B6)
            .Left = BannerLeft + DP_HEADER_NEXTMONTH_ARROW_LEFT
            .Top = DP_HEADER_TOP + DP_HEADER_MONTH_ARROW_TOP + 2
            .Width = DP_HEADER_MONTH_ARROW_WIDTH
            .Height = DP_HEADER_MONTH_ARROW_HEIGHT - 5
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_HEADER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Next Month"
        End With

    'Create a clean next-month font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the next-month font explicitly
        With HeaderFont
            .Name = DP_FORM_FONT_NAME
            .Size = 8
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean font to the next-month label
        Set Lbl_NextMonth.Font = HeaderFont

    'Move the next-month label to the front
        Lbl_NextMonth.ZOrder 0

    'Create the next-month click / hover hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the next-month label to the next-month action
        LabelHook.Initialize Me, Lbl_NextMonth, "NEXT_MONTH", "HEADER"

    'Store the hook so that the click / hover events remain alive
        mHeaderLabelHooks.Add LabelHook, Lbl_NextMonth.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT COMPACT SETTINGS LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the compact header settings label
        Set Lbl_HeaderSettings = UF_Ensure_Label("Lbl_HeaderSettings")

    'Apply layout and visual properties
        With Lbl_HeaderSettings
            .Caption = vbNullString
            .Left = Lbl_NextMonth.Left + Lbl_NextMonth.Width + DP_HEADER_SETTINGS_ICON_GAP
            .Top = Lbl_NextMonth.Top + DP_HEADER_SETTINGS_ICON_TOP_OFFSET
            .Width = DP_HEADER_SETTINGS_ICON_WIDTH
            .Height = DP_HEADER_SETTINGS_ICON_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_HEADER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = (gDP_SizeMode = DP_SizeMode_Compact)
            .ControlTipText = DP_HEADER_SETTINGS_ICON_TOOLTIP
        End With

'------------------------------------------------------------------------------
' APPLY COMPACT SETTINGS ICON FONT
'------------------------------------------------------------------------------
    'Create a clean compact settings icon font object
        Set HeaderIconFont = CreateObject("StdFont")

    'Configure the compact settings icon font explicitly
        With HeaderIconFont
            .Name = DP_HEADER_SETTINGS_ICON_FONT_NAME
            .Size = DP_HEADER_SETTINGS_ICON_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean font to the compact settings icon label
        Set Lbl_HeaderSettings.Font = HeaderIconFont

    'Apply the Settings glyph after assigning the clean font
        Lbl_HeaderSettings.Caption = ChrW$(DP_HEADER_SETTINGS_ICON_CODEPOINT)

    'Move the compact settings icon to the front
        Lbl_HeaderSettings.ZOrder 0

'------------------------------------------------------------------------------
' REGISTER COMPACT SETTINGS LABEL HOOK
'------------------------------------------------------------------------------
    'Create the compact Settings icon click / hover hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the compact Settings icon to the settings action
        LabelHook.Initialize Me, Lbl_HeaderSettings, "SHOW_SETTINGS", "HEADER"

    'Store the hook so that the click / hover events remain alive
        mHeaderLabelHooks.Add LabelHook, Lbl_HeaderSettings.Name

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
        Err.Raise Err.Number, PROC_NAME, "Header label creation failed: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' NAVIGATION AND DISPLAY STATE
'------------------------------------------------------------------------------

Private Sub UF_Header_MoveMonth(ByVal MonthOffset As Long)

'
'------------------------------------------------------------------------------
'                           MOVE HEADER MONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Moves the displayed DatePicker month backward or forward
'
' WHY THIS EXISTS
'   The header month arrows should allow direct navigation across months while
'   keeping the header captions, keyboard date, overlay panels, and day grid
'   synchronized
'
' INPUTS
'   MonthOffset
'     Number of months to move
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the current display period, resolves the new display period,
'   rejects non-renderable calendar boundaries, stores the new display state,
'   resets transient overlay / hover state, hides overlay panels, initializes
'   keyboard navigation to the first day of the new month, and refreshes the day
'   grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if MonthOffset is unsupported, if the
'   current display period is invalid, if the resulting display period cannot be
'   rendered, or if the day grid cannot be refreshed
'
' DEPENDENCIES
'   UF_PickerPanel_HoverReset
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'   UF_SettingsPanel_Hide
'   UF_DayGrid_Populate
'
' NOTES
'   DateAdd is used so month transitions across year boundaries are handled by
'   VBA date arithmetic
'
'   January 100 and December 9999 are rejected because the fixed calendar grid
'   cannot render the required adjacent-month dates outside the VBA Date
'   supported range
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_Header_MoveMonth"
    
    Const VBA_DATE_MIN_YEAR     As Long = 100               'Minimum year supported by VBA Date
    Const VBA_DATE_MAX_YEAR     As Long = 9999              'Maximum year supported by VBA Date

    Dim ExistingControl         As MSForms.control          'Control resolved by picker-panel name
    Dim Fra_PickerPanel         As MSForms.Frame            'Reusable picker panel
    Dim NewDisplayDate          As Date                     'New first day of displayed month
    Dim NewDisplayYear          As Long                     'New displayed year
    Dim NewDisplayMonth         As Long                     'New displayed month
    Dim HandlerStep             As String                   'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate MonthOffset"

    'Reject unsupported month movement
        If MonthOffset <> -1 And MonthOffset <> 1 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "MonthOffset must be -1 or 1"
        End If

'------------------------------------------------------------------------------
' VALIDATE CURRENT DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate current display period"

    'Reject invalid current display year
        If mDisplayYear < DP_MIN_YEAR Or mDisplayYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "mDisplayYear must be between " & VBA.CStr(DP_MIN_YEAR) & _
                " and " & VBA.CStr(DP_MAX_YEAR)
        End If
    'Reject invalid current display month
        If mDisplayMonth < 1 Or mDisplayMonth > 12 Then
            Err.Raise vbObjectError + 515, PROC_NAME, "mDisplayMonth must be between 1 and 12"
        End If
    'Reject movement before the lower supported DatePicker boundary
        If mDisplayYear = DP_MIN_YEAR And mDisplayMonth = 1 And MonthOffset < 0 Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Displayed month cannot move before January " & VBA.CStr(DP_MIN_YEAR)
        End If
    'Reject movement after the upper supported DatePicker boundary
        If mDisplayYear = DP_MAX_YEAR And mDisplayMonth = 12 And MonthOffset > 0 Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Displayed month cannot move after December " & VBA.CStr(DP_MAX_YEAR)
        End If

'------------------------------------------------------------------------------
' RESOLVE NEW DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve new display period"

    'Resolve the new display date
        NewDisplayDate = VBA.DateAdd("m", MonthOffset, _
            VBA.DateSerial(mDisplayYear, mDisplayMonth, 1))
    'Resolve the new display year
        NewDisplayYear = VBA.Year(NewDisplayDate)
    'Resolve the new display month
        NewDisplayMonth = VBA.Month(NewDisplayDate)

'------------------------------------------------------------------------------
' VALIDATE NEW DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate new display period"

    'Reject display years outside the supported DatePicker range
        If NewDisplayYear < DP_MIN_YEAR Or NewDisplayYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "New display year must be between " & VBA.CStr(DP_MIN_YEAR) & _
                " and " & VBA.CStr(DP_MAX_YEAR)
        End If
    'Reject the lower VBA Date rendering boundary
        If NewDisplayYear = VBA_DATE_MIN_YEAR And NewDisplayMonth = 1 Then
            Err.Raise vbObjectError + 519, PROC_NAME, _
                "January " & VBA.CStr(VBA_DATE_MIN_YEAR) & _
                " cannot render a complete DatePicker grid"
        End If
    'Reject the upper VBA Date rendering boundary
        If NewDisplayYear = VBA_DATE_MAX_YEAR And NewDisplayMonth = 12 Then
            Err.Raise vbObjectError + 520, PROC_NAME, _
                "December " & VBA.CStr(VBA_DATE_MAX_YEAR) & _
                " cannot render a complete DatePicker grid"
        End If

'------------------------------------------------------------------------------
' STORE DISPLAY STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Store display state"

    'Store the new display year
        mDisplayYear = NewDisplayYear
    'Store the new display month
        mDisplayMonth = NewDisplayMonth
    'Initialize keyboard date to the first day of the new display month
        mKeyboardDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
    'Mark keyboard navigation as initialized
        mHasKeyboardDate = True

'------------------------------------------------------------------------------
' RESET TRANSIENT HOVER STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Reset transient hover state"

    'Clear picker-panel hover before hiding overlays
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Clear day-cell hover before repopulating the grid
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Clear header hover before refreshing captions
        If VBA.Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Clear footer hover before hiding overlays
        If VBA.Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset
    'Clear settings-panel hover before hiding overlays
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then UF_SettingsPanel_HoverReset

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Hide picker panel"

    'Suppress lookup errors when the picker panel is not available
        On Error Resume Next
    'Retrieve the picker panel control when available
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Hide the picker panel when it exists and is a frame
        If Not ExistingControl Is Nothing Then
            If VBA.TypeName(ExistingControl) = "Frame" Then
                Set Fra_PickerPanel = ExistingControl
                Fra_PickerPanel.Visible = False
            End If
        End If
    'Clear picker-panel mode after month navigation
        mPickerPanelMode = 0
    'Clear picker-panel hover state after month navigation
        mHoveredPickerItemIndex = 0

'------------------------------------------------------------------------------
' HIDE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Hide settings panel"

    'Hide the settings panel after month navigation
        UF_SettingsPanel_Hide

'------------------------------------------------------------------------------
' REFRESH CALENDAR
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh day grid"

    'Refresh the day grid
        UF_DayGrid_Populate mDisplayYear, mDisplayMonth

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
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Header month navigation failed: " & Err.Description

End Sub

Private Sub UF_Header_MoveYear(ByVal YearOffset As Long)

'
'------------------------------------------------------------------------------
'                           MOVE HEADER YEAR
'------------------------------------------------------------------------------
' PURPOSE
'   Handles header year-arrow navigation
'
' WHY THIS EXISTS
'   The year arrows scroll the visible year picker range when the year panel is
'   visible, otherwise they move the displayed calendar year
'
' INPUTS
'   YearOffset
'     Direction of movement
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Scrolls the 12-year picker range when the year panel is visible in year mode.
'   Otherwise validates the current display period, resolves the new display
'   year, rejects non-renderable calendar boundaries, stores the new display
'   state, hides overlay panels, initializes keyboard navigation to the first day
'   of the displayed month, and refreshes the day grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if YearOffset is unsupported, if the
'   current display period is invalid, if the resulting display period cannot be
'   rendered, or if the panel / grid cannot be refreshed
'
' DEPENDENCIES
'   UF_PickerPanel_HoverReset
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'   UF_SettingsPanel_Hide
'   UF_PickerPanel_PopulateYears
'   UF_DayGrid_Populate
'
' NOTES
'   The year panel start is clamped to the valid 12-year display range
'
'   January 100 and December 9999 are rejected because the fixed calendar grid
'   cannot render the required adjacent-month dates outside the VBA Date
'   supported range
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_Header_MoveYear"
    
    Const VBA_DATE_MIN_YEAR     As Long = 100               'Minimum year supported by VBA Date
    Const VBA_DATE_MAX_YEAR     As Long = 9999              'Maximum year supported by VBA Date
    Const PICKER_MODE_NONE      As Long = 0                 'No picker-panel mode
    Const PICKER_MODE_YEARS     As Long = 2                 'Year picker-panel mode

    Dim ExistingControl         As MSForms.control          'Control resolved by picker-panel name
    Dim Fra_PickerPanel         As MSForms.Frame            'Reusable picker panel
    
    Dim NewDisplayYear          As Long                     'New displayed year
    Dim NewYearPanelStart       As Long                     'New first year shown in panel
    Dim HandlerStep             As String                   'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate YearOffset"

    'Reject unsupported year movement
        If YearOffset <> -1 And YearOffset <> 1 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "YearOffset must be -1 or 1"
        End If

'------------------------------------------------------------------------------
' RETRIEVE PICKER PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve picker panel"

    'Suppress lookup errors when the picker panel is not available
        On Error Resume Next
    'Retrieve the picker panel control when available
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Use the picker panel when it exists and is a frame
        If Not ExistingControl Is Nothing Then
            'Reject a name collision with a non-frame control
                If VBA.TypeName(ExistingControl) <> "Frame" Then
                    Err.Raise vbObjectError + 514, PROC_NAME, _
                        "Control '" & DP_PICKER_PANEL_NAME & "' exists but is not an MSForms.Frame"
                End If
            'Store the picker panel reference
                Set Fra_PickerPanel = ExistingControl
        End If

'------------------------------------------------------------------------------
' SCROLL YEAR PANEL IF VISIBLE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Evaluate year-panel scrolling"

    'Scroll the 12-year panel when it is visible in year-selection mode
        If Not Fra_PickerPanel Is Nothing Then
            If Fra_PickerPanel.Visible Then
                If mPickerPanelMode = PICKER_MODE_YEARS Then
                    'Track the current handler step
                        HandlerStep = "Scroll year picker panel"
                    'Clear active picker hover before repopulating the year panel
                        UF_PickerPanel_HoverReset
                    'Calculate the new first visible year
                        NewYearPanelStart = mYearPanelStart + _
                            (YearOffset * DP_PICKER_ITEM_COUNT)
                    'Clamp the first visible year to the lower supported range
                        If NewYearPanelStart < DP_MIN_YEAR Then
                            NewYearPanelStart = DP_MIN_YEAR
                        End If
                    'Clamp the first visible year to the upper supported range
                        If NewYearPanelStart > DP_YEAR_PANEL_MAX_START Then
                            NewYearPanelStart = DP_YEAR_PANEL_MAX_START
                        End If
                    'Exit if the year-panel range did not change
                        If NewYearPanelStart = mYearPanelStart Then Exit Sub
                    'Store the new first visible year
                        mYearPanelStart = NewYearPanelStart
                    'Refresh the year picker panel
                        UF_PickerPanel_PopulateYears
                    'Exit after scrolling the panel
                        Exit Sub
                End If
            End If
        End If

'------------------------------------------------------------------------------
' VALIDATE CURRENT DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate current display period"

    'Reject invalid current displayed year
        If mDisplayYear < DP_MIN_YEAR Or mDisplayYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "mDisplayYear must be between " & VBA.CStr(DP_MIN_YEAR) & _
                " and " & VBA.CStr(DP_MAX_YEAR)
        End If
    'Reject invalid current displayed month
        If mDisplayMonth < 1 Or mDisplayMonth > 12 Then
            Err.Raise vbObjectError + 516, PROC_NAME, "mDisplayMonth must be between 1 and 12"
        End If

'------------------------------------------------------------------------------
' RESOLVE NEW DISPLAY YEAR
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve new display year"

    'Resolve the new displayed year
        NewDisplayYear = mDisplayYear + YearOffset
    'Reject years outside the supported DatePicker range
        If NewDisplayYear < DP_MIN_YEAR Or NewDisplayYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Displayed year must remain between " & VBA.CStr(DP_MIN_YEAR) & _
                " and " & VBA.CStr(DP_MAX_YEAR)
        End If

'------------------------------------------------------------------------------
' VALIDATE NEW DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate new display period"

    'Reject the lower VBA Date rendering boundary
        If NewDisplayYear = VBA_DATE_MIN_YEAR And mDisplayMonth = 1 Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "January " & VBA.CStr(VBA_DATE_MIN_YEAR) & _
                " cannot render a complete DatePicker grid"
        End If
    'Reject the upper VBA Date rendering boundary
        If NewDisplayYear = VBA_DATE_MAX_YEAR And mDisplayMonth = 12 Then
            Err.Raise vbObjectError + 519, PROC_NAME, _
                "December " & VBA.CStr(VBA_DATE_MAX_YEAR) & _
                " cannot render a complete DatePicker grid"
        End If

'------------------------------------------------------------------------------
' STORE DISPLAY STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Store display state"

    'Store the new displayed year
        mDisplayYear = NewDisplayYear
    'Initialize keyboard date to the first day of the current display month in the new year
        mKeyboardDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
    'Mark keyboard navigation as initialized
        mHasKeyboardDate = True

'------------------------------------------------------------------------------
' RESET TRANSIENT HOVER STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Reset transient hover state"

    'Clear picker-panel hover before hiding overlays
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Clear day-cell hover before repopulating the grid
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Clear header hover before refreshing captions
        If VBA.Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Clear footer hover before hiding overlays
        If VBA.Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset
    'Clear settings-panel hover before hiding overlays
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then UF_SettingsPanel_HoverReset

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Hide picker panel"

    'Hide the reusable picker panel when available
        If Not Fra_PickerPanel Is Nothing Then
            Fra_PickerPanel.Visible = False
        End If
    'Clear picker-panel mode after normal year navigation
        mPickerPanelMode = PICKER_MODE_NONE
    'Clear picker-panel hover state after normal year navigation
        mHoveredPickerItemIndex = 0

'------------------------------------------------------------------------------
' HIDE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Hide settings panel"

    'Hide the settings panel after year navigation
        UF_SettingsPanel_Hide

'------------------------------------------------------------------------------
' REFRESH CALENDAR
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh day grid"

    'Refresh the day grid using the new displayed year
        UF_DayGrid_Populate mDisplayYear, mDisplayMonth

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
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Header year navigation failed: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' HEADER HOVER
'------------------------------------------------------------------------------

Public Sub UF_Header_HoverApply(ByVal LabelName As String)

'
'------------------------------------------------------------------------------
'                           APPLY HEADER CLICKABLE HOVER
'------------------------------------------------------------------------------
' PURPOSE
'   Applies hover formatting to one clickable DatePicker header label
'
' WHY THIS EXISTS
'   Runtime-created header labels are clickable. The UI should give the user a
'   visible cue when the mouse moves over them
'
' INPUTS
'   LabelName
'     Name of the header label currently under the mouse pointer
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves supported clickable header labels, resets active hover state from
'   other DatePicker areas, exits when the same header label is already hovered,
'   applies a subtle hover state to the requested header label, and stores the
'   current header hover state
'
' ERROR POLICY
'   Raises a descriptive runtime error if LabelName is empty or if a supported
'   header label cannot be resolved or formatted
'
'   Unsupported header labels are treated as neutral hover-reset areas
'
' DEPENDENCIES
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_PickerPanel_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'
' NOTES
'   This routine avoids changing Font properties because header labels have
'   different normal bold states
'
'   Label matching is normalized so behavior is independent from Option Compare
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_Header_HoverApply"

    Dim RawLabelName            As String               'Trimmed label name supplied by caller
    Dim NormalizedLabelName     As String               'Uppercase label name used for routing
    Dim TargetLabelName         As String               'Canonical target label name
    
    Dim ExistingControl         As MSForms.control      'Control resolved by name
    Dim Lbl_Header              As MSForms.Label        'Hovered header label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the supplied label name
        RawLabelName = VBA.Trim$(LabelName)
    'Reject an empty label name
        If VBA.Len(RawLabelName) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "LabelName cannot be empty"
        End If
    'Normalize the label name for Option Compare independent routing
        NormalizedLabelName = VBA.UCase$(RawLabelName)

'------------------------------------------------------------------------------
' RESOLVE TARGET HEADER LABEL
'------------------------------------------------------------------------------
    'Resolve supported clickable header labels
        Select Case NormalizedLabelName
            Case "LBL_HEADERMONTH"
                TargetLabelName = "Lbl_HeaderMonth"
            Case "LBL_HEADERYEAR"
                TargetLabelName = "Lbl_HeaderYear"
            Case "LBL_PREVMONTH"
                TargetLabelName = "Lbl_PrevMonth"
            Case "LBL_NEXTMONTH"
                TargetLabelName = "Lbl_NextMonth"
            Case "LBL_PREVYEAR"
                TargetLabelName = "Lbl_PrevYear"
            Case "LBL_NEXTYEAR"
                TargetLabelName = "Lbl_NextYear"
            Case "LBL_HEADERSETTINGS"
                TargetLabelName = "Lbl_HeaderSettings"
            Case Else
                'Remove previous header hover state when moving over neutral header surfaces
                    If VBA.Len(mHoveredHeaderLabelName) <> 0 Then
                        UF_Header_HoverReset
                    End If
                'Exit because neutral header labels do not receive hover styling
                    Exit Sub
        End Select

'------------------------------------------------------------------------------
' RESET OTHER HOVER STATES
'------------------------------------------------------------------------------
    'Clear active day-cell hover when entering the header
        If mHoveredDayCellIndex <> 0 Then
            UF_DayCell_HoverReset
        End If
    'Clear active picker-panel hover when entering the header
        If mHoveredPickerItemIndex <> 0 Then
            UF_PickerPanel_HoverReset
        End If
    'Clear active footer hover when entering the header
        If VBA.Len(mHoveredFooterActionName) <> 0 Then
            UF_Footer_HoverReset
        End If
    'Clear active settings-panel hover when entering the header
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then
            UF_SettingsPanel_HoverReset
        End If

'------------------------------------------------------------------------------
' EXIT IF ALREADY HOVERED
'------------------------------------------------------------------------------
    'Exit when the same header label is already highlighted
        If VBA.StrComp(mHoveredHeaderLabelName, TargetLabelName, vbBinaryCompare) = 0 Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RESET CURRENT HEADER HOVER STATE
'------------------------------------------------------------------------------
    'Remove hover formatting from the previously highlighted header label
        If VBA.Len(mHoveredHeaderLabelName) <> 0 Then
            UF_Header_HoverReset
        End If

'------------------------------------------------------------------------------
' RETRIEVE TARGET LABEL
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the target header label
        On Error Resume Next
    'Retrieve the target header label
        Set ExistingControl = Me.Controls(TargetLabelName)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing target header label
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Unable to resolve expected header label " & TargetLabelName
        End If
    'Reject a name collision with a non-label control
        If VBA.TypeName(ExistingControl) <> "Label" Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Control '" & TargetLabelName & "' exists but is not an MSForms.Label"
        End If
    'Use the resolved header label
        Set Lbl_Header = ExistingControl

'------------------------------------------------------------------------------
' APPLY HOVER STATE
'------------------------------------------------------------------------------
    'Apply clickable hover visual formatting
        With Lbl_Header
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_HEADER_HOVER_BACK_COLOR
            .ForeColor = vbWhite
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = vbWhite
            .SpecialEffect = fmSpecialEffectFlat
        End With
    'Move the hovered header label to the front
        Lbl_Header.ZOrder 0

'------------------------------------------------------------------------------
' STORE HOVER STATE
'------------------------------------------------------------------------------
    'Store the current hovered header label name
        mHoveredHeaderLabelName = TargetLabelName

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
        Err.Raise Err.Number, PROC_NAME, "Header hover application failed: " & Err.Description

End Sub

Private Sub UF_Header_HoverReset()

'
'------------------------------------------------------------------------------
'                           RESET HEADER CLICKABLE HOVER
'------------------------------------------------------------------------------
' PURPOSE
'   Removes hover formatting from the currently highlighted clickable header
'   label
'
' WHY THIS EXISTS
'   Header hover formatting must be removed when the mouse leaves clickable
'   header labels and moves back over another DatePicker surface
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the currently highlighted header label, restores its normal visual
'   state when available, and clears the stored header hover state
'
' ERROR POLICY
'   Best-effort UI cleanup. Suppresses reset errors because hover reset should
'   never interrupt UserForm interaction
'
' DEPENDENCIES
'   UserForm.Controls collection
'
' NOTES
'   This routine restores only the visual properties changed by
'   UF_Header_HoverApply
'
'   Stale or missing runtime labels are ignored because controls may have been
'   rebuilt while a hover state was active
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentLabelName    As String               'Current hovered header label name
    Dim Lbl_Header          As MSForms.Label        'Previously highlighted header label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress hover-reset errors
        On Error Resume Next
    'Capture the current hovered header label
        CurrentLabelName = VBA.Trim$(mHoveredHeaderLabelName)

'------------------------------------------------------------------------------
' EXIT IF NOTHING IS HOVERED
'------------------------------------------------------------------------------
    'Exit if no header label is currently highlighted
        If VBA.Len(CurrentLabelName) = 0 Then GoTo Clean_Exit

'------------------------------------------------------------------------------
' RETRIEVE CURRENTLY HOVERED LABEL
'------------------------------------------------------------------------------
    'Retrieve the previously highlighted header label
        Set Lbl_Header = Me.Controls(CurrentLabelName)

'------------------------------------------------------------------------------
' RESET VISUAL STATE
'------------------------------------------------------------------------------
    'Restore normal clickable-header formatting when available
        If Not Lbl_Header Is Nothing Then
            With Lbl_Header
                .BackStyle = fmBackStyleTransparent
                .BackColor = DP_HEADER_BACK_COLOR
                .ForeColor = DP_HEADER_FORE_COLOR
                .BorderStyle = fmBorderStyleNone
                .BorderColor = DP_HEADER_BACK_COLOR
                .SpecialEffect = fmSpecialEffectFlat
            End With
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:
    'Clear the current hovered header label name
        mHoveredHeaderLabelName = vbNullString
    'Clear suppressed reset errors
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                               DAY-OF-WEEK ROW
'
'------------------------------------------------------------------------------
'

Public Sub UF_WeekdayRow_Build()

'
'------------------------------------------------------------------------------
'                       CREATE DAY OF WEEK LABELS ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Creates and formats the day-of-week header labels displayed above the
'   DatePicker calendar grid
'
' WHY THIS EXISTS
'   The calendar grid needs a weekday header row that reflects the saved
'   first-day-of-week setting and the saved local-dependent / local-independent
'   caption setting
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses Lbl_DayOfWeek1 to Lbl_DayOfWeek7, resolves each weekday
'   caption, applies the standard weekday-row layout, assigns a clean font, and
'   brings the labels to the front
'
' ERROR POLICY
'   Raises a descriptive runtime error if the weekday setting is invalid or if a
'   weekday label cannot be created or formatted
'
' DEPENDENCIES
'   UF_WeekdayCaption_Get
'   M_Settings_IsValidFirstDayOfWeek
'   UF_Ensure_Label
'   gDP_FirstDayOfWeek
'   gDP_UseLocalNames
'
' NOTES
'   This routine must stay in the UserForm module because it uses Me.Controls
'
'   It is Public so external refresh routines can update the weekday header when
'   settings are changed while the form is already loaded
'
'   Developer-owned layout constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_WeekdayRow_Build"

    Dim Index               As Long                 'Sequential weekday label index
    Dim ControlName         As String               'Runtime control name
    Dim HeaderLeft          As Single               'Calculated weekday row left position
    Dim HeaderTop           As Single               'Calculated weekday row top position
    Dim DayCaption          As String               'Resolved weekday caption
    Dim Lbl_DayOfWeek       As MSForms.Label        'Weekday label control
    Dim WeekdayFont         As Object               'Clean weekday-label font object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE SETTINGS
'------------------------------------------------------------------------------
    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(gDP_FirstDayOfWeek) Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "gDP_FirstDayOfWeek must be vbSunday or vbMonday"
        End If

'------------------------------------------------------------------------------
' CALCULATE LAYOUT
'------------------------------------------------------------------------------
    'Calculate the left position that centers wider weekday labels above day labels
        HeaderLeft = DP_DAY_GRID_START_LEFT + _
            ((DP_DAY_LABEL_WIDTH - DP_DOW_LABEL_WIDTH) / 2)
    'Calculate the top position of the weekday header row
        HeaderTop = DP_DAY_GRID_START_TOP - DP_DAY_LABEL_VERTICAL_STEP

'------------------------------------------------------------------------------
' CREATE CLEAN FONT
'------------------------------------------------------------------------------
    'Create a clean weekday-label font object
        Set WeekdayFont = CreateObject("StdFont")

    'Configure the weekday-label font explicitly
        With WeekdayFont
            .Name = DP_FORM_FONT_NAME
            .Size = DP_FORM_FONT_SIZE
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

'------------------------------------------------------------------------------
' CREATE / FORMAT DAY-OF-WEEK LABELS
'------------------------------------------------------------------------------
    'Loop through the weekday header labels
        For Index = 1 To DP_DAY_LABELS_PER_ROW
            'Build the runtime control name
                ControlName = "Lbl_DayOfWeek" & CStr(Index)
            'Resolve the weekday caption
                DayCaption = UF_WeekdayCaption_Get(Index, gDP_FirstDayOfWeek, gDP_UseLocalNames)
            'Create or retrieve the weekday label
                Set Lbl_DayOfWeek = UF_Ensure_Label(ControlName)
            'Apply layout and visual properties
                With Lbl_DayOfWeek
                    .Caption = vbNullString
                    .Left = HeaderLeft + DP_DOW_LABEL_HORIZONTAL_STEP * (Index - 1)
                    .Top = HeaderTop
                    .Width = DP_DOW_LABEL_WIDTH
                    .Height = DP_DOW_LABEL_HEIGHT
                    .BackStyle = fmBackStyleTransparent
                    .ForeColor = DP_HEADER_FORE_COLOR
                    .BorderStyle = fmBorderStyleNone
                    .SpecialEffect = fmSpecialEffectFlat
                    .TextAlign = fmTextAlignCenter
                    .WordWrap = False
                    .Enabled = True
                    .Visible = True
                End With
            'Assign the clean font to the weekday label
                Set Lbl_DayOfWeek.Font = WeekdayFont
            'Apply the weekday caption after the clean font has been assigned
                Lbl_DayOfWeek.Caption = DayCaption
            'Move the weekday label to the front
                Lbl_DayOfWeek.ZOrder 0
        Next Index

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
        Err.Raise Err.Number, PROC_NAME, "Weekday row creation failed: " & Err.Description

End Sub

Private Function UF_WeekdayCaption_Get( _
    ByVal Index As Long, _
    ByVal FirstDayOfWeek As Long, _
    ByVal UseLocalDayNames As Boolean) As String

'
'------------------------------------------------------------------------------
'                           GET DAY OF WEEK CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the weekday caption for a given weekday label index
'
' WHY THIS EXISTS
'   The DatePicker supports both locale-dependent weekday captions and fixed
'   English weekday captions while allowing Sunday or Monday starts
'
' INPUTS
'   Index
'     One-based display index from 1 to 7
'
'   FirstDayOfWeek
'     First day displayed in the header row
'
'   UseLocalDayNames
'     True to use local weekday names through WeekdayName
'
' RETURNS
'   Uppercase weekday caption
'
' BEHAVIOR
'   Converts the display index into the corresponding absolute VBA weekday
'   number, resolves either the local or fixed-English caption, trims it, and
'   returns it in uppercase
'
' ERROR POLICY
'   Raises a descriptive runtime error if Index or FirstDayOfWeek is invalid
'
' DEPENDENCIES
'   M_Settings_IsValidFirstDayOfWeek
'   UF_WeekdayCaption_GetFixedEnglish
'   VBA.WeekdayName
'
' NOTES
'   Local captions depend on the user's Windows or Office language environment
'
'   The function intentionally supports only the DatePicker first-day policy:
'   vbSunday and vbMonday
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_WeekdayCaption_Get"

    Dim DayNumber           As Long     'Resolved absolute VBA weekday number
    Dim DayCaption          As String   'Resolved weekday caption

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject invalid display index
        If Index < 1 Or Index > DP_DAY_LABELS_PER_ROW Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Index must be between 1 and " & CStr(DP_DAY_LABELS_PER_ROW)
        End If
    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(FirstDayOfWeek) Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "FirstDayOfWeek must be vbSunday or vbMonday"
        End If

'------------------------------------------------------------------------------
' RESOLVE DISPLAY DAY
'------------------------------------------------------------------------------
    'Resolve the absolute VBA weekday number for the requested display index
        DayNumber = ((FirstDayOfWeek - 1 + Index - 1) Mod DP_DAY_LABELS_PER_ROW) + 1

'------------------------------------------------------------------------------
' GET CAPTION
'------------------------------------------------------------------------------
    'Get the requested weekday caption
        If UseLocalDayNames Then
            DayCaption = VBA.WeekdayName(DayNumber, True, vbSunday)
        Else
            DayCaption = UF_WeekdayCaption_GetFixedEnglish(DayNumber)
        End If

'------------------------------------------------------------------------------
' RETURN CAPTION
'------------------------------------------------------------------------------
    'Return the normalized caption
        UF_WeekdayCaption_Get = VBA.UCase$(VBA.Trim$(DayCaption))

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
        Err.Raise Err.Number, PROC_NAME, "Weekday caption resolution failed: " & Err.Description

End Function

Private Function UF_WeekdayCaption_GetFixedEnglish(ByVal DayNumber As Long) As String

'
'------------------------------------------------------------------------------
'                   GET FIXED ENGLISH DAY OF WEEK CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the fixed English weekday caption for a VBA weekday number
'
' WHY THIS EXISTS
'   The DatePicker can operate in local-independent mode, where weekday captions
'   should not depend on Windows, Office, or regional language settings
'
' INPUTS
'   DayNumber
'     VBA weekday number from vbSunday to vbSaturday
'
' RETURNS
'   Fixed three-letter English weekday caption
'
' BEHAVIOR
'   Maps the supplied VBA weekday number to a fixed English caption:
'     - SUN
'     - MON
'     - TUE
'     - WED
'     - THU
'     - FRI
'     - SAT
'
' ERROR POLICY
'   Raises a descriptive runtime error if DayNumber is outside vbSunday to
'   vbSaturday
'
' DEPENDENCIES
'   VBA weekday constants
'
' NOTES
'   Returned captions are intentionally fixed in English
'
'   This function is used when local-dependent weekday names are disabled
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_DatePicker.UF_WeekdayCaption_GetFixedEnglish"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject invalid day numbers
        If DayNumber < vbSunday Or DayNumber > vbSaturday Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DayNumber must be between vbSunday and vbSaturday"
        End If

'------------------------------------------------------------------------------
' RETURN CAPTION
'------------------------------------------------------------------------------
    'Return the fixed English caption
        Select Case DayNumber

            Case vbSunday
                UF_WeekdayCaption_GetFixedEnglish = "SUN"
            Case vbMonday
                UF_WeekdayCaption_GetFixedEnglish = "MON"
            Case vbTuesday
                UF_WeekdayCaption_GetFixedEnglish = "TUE"
            Case vbWednesday
                UF_WeekdayCaption_GetFixedEnglish = "WED"
            Case vbThursday
                UF_WeekdayCaption_GetFixedEnglish = "THU"
            Case vbFriday
                UF_WeekdayCaption_GetFixedEnglish = "FRI"
            Case vbSaturday
                UF_WeekdayCaption_GetFixedEnglish = "SAT"

        End Select

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
        Err.Raise Err.Number, PROC_NAME, "Fixed English weekday caption resolution failed: " & Err.Description

End Function

'
'------------------------------------------------------------------------------
'
'                                   DAY GRID
'
'------------------------------------------------------------------------------
'

Private Sub UF_DayGrid_Build()

'
'------------------------------------------------------------------------------
'                           CREATE DAY LABELS GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Creates and formats the 42 paired day cells used by the DatePicker calendar
'   grid
'
' WHY THIS EXISTS
'   Each visible calendar day is represented by two labels:
'     - a larger background label used as the hover / click target
'     - a smaller foreground text label used to display the day number
'
'   This gives the DatePicker a larger interaction target while preserving a
'   compact and visually centered day-number caption
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses Lbl_DayBg1 to Lbl_DayBg42 and Lbl_Day1 to Lbl_Day42,
'   formats each paired day cell, caches the label references, and connects both
'   labels in each pair to the same routed day action
'
' ERROR POLICY
'   Raises a descriptive runtime error if a day-grid label cannot be created,
'   formatted, cached, layered, or hooked
'
' DEPENDENCIES
'   UF_Ensure_Label
'   cDatePickerLabelHook
'   mDayLabelHooks
'   mDayBackLabels
'   mDayTextLabels
'
' NOTES
'   The background label is placed behind the foreground text label
'
'   The text label remains small and transparent to simulate vertical alignment
'   inside the larger day-cell background
'
'   Developer-owned layout constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_DayGrid_Build"

    Dim Index               As Long                     'Sequential day-cell index
    Dim RowIndex            As Long                     'Zero-based row index
    Dim ColIndex            As Long                     'Zero-based column index
    Dim TextLeft            As Single                   'Day text label left position
    Dim TextTop             As Single                   'Day text label top position
    Dim CellLeft            As Single                   'Day background label left position
    Dim CellTop             As Single                   'Day background label top position
    Dim TextControlName     As String                   'Runtime day text label name
    Dim BgControlName       As String                   'Runtime day background label name
    Dim ActionName          As String                   'Runtime action routed by the label hook
    Dim Lbl_Day             As MSForms.Label            'Day text label control
    Dim Lbl_DayBg           As MSForms.Label            'Day background label control
    Dim LabelHook           As cDatePickerLabelHook     'Runtime label event hook

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' INITIALIZE EVENT HOOK STORAGE
'------------------------------------------------------------------------------
    'Create or reset the day-label hook collection
        Set mDayLabelHooks = New Collection
    'Clear the current day-cell hover index
        mHoveredDayCellIndex = 0
    'Build the day-cell label-name to index map once
        UF_DayGrid_BuildIndexMap
        
'------------------------------------------------------------------------------
' CREATE / FORMAT DAY CELLS
'------------------------------------------------------------------------------
    'Loop through the full calendar grid
        For Index = 1 To DP_DAY_LABEL_COUNT
            'Calculate the zero-based row index
                RowIndex = (Index - 1) \ DP_DAY_LABELS_PER_ROW
            'Calculate the zero-based column index
                ColIndex = (Index - 1) Mod DP_DAY_LABELS_PER_ROW
            
            'Calculate the day text label left position
                TextLeft = DP_DAY_GRID_START_LEFT + DP_DAY_LABEL_HORIZONTAL_STEP * ColIndex
            'Calculate the day text label top position
                TextTop = DP_DAY_GRID_START_TOP + DP_DAY_LABEL_VERTICAL_STEP * RowIndex
            
            'Center the larger background cell horizontally around the text label
                CellLeft = TextLeft - ((DP_DAY_CELL_WIDTH - DP_DAY_LABEL_WIDTH) / 2)
            'Center the larger background cell vertically around the text label
                CellTop = TextTop - ((DP_DAY_CELL_HEIGHT - DP_DAY_LABEL_HEIGHT) / 2)

            'Build the runtime day text label name
                TextControlName = "Lbl_Day" & VBA.CStr(Index)
            'Build the runtime day background label name
                BgControlName = "Lbl_DayBg" & VBA.CStr(Index)

            'Build the action name routed by both day labels
                ActionName = "DAY_PICKED_" & VBA.CStr(Index)

'------------------------------------------------------------------------------
' CREATE / FORMAT BACKGROUND LABEL
'------------------------------------------------------------------------------
            'Create or retrieve the day background label
                Set Lbl_DayBg = UF_Ensure_Label(BgControlName)
            'Cache the day background label for fast reuse
                Set mDayBackLabels(Index) = Lbl_DayBg
            'Apply layout and visual properties to the day background label
                With Lbl_DayBg
                    .Caption = vbNullString
                    .Left = CellLeft
                    .Top = CellTop
                    .Width = DP_DAY_CELL_WIDTH
                    .Height = DP_DAY_CELL_HEIGHT
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = DP_DAY_NORMAL_BACK_COLOR
                    .ForeColor = DP_DAY_NORMAL_FORE_COLOR
                    .BorderStyle = fmBorderStyleNone
                    .BorderColor = DP_DAY_NORMAL_BACK_COLOR
                    .SpecialEffect = fmSpecialEffectFlat
                    .TextAlign = fmTextAlignCenter
                    .WordWrap = False
                    .Enabled = True
                    .Visible = True
                End With
            'Create the background-label event hook
                Set LabelHook = New cDatePickerLabelHook
            'Connect the day background label to its routed action
                LabelHook.Initialize Me, Lbl_DayBg, ActionName, "DAY"
            'Store the hook so that its events remain alive
                mDayLabelHooks.Add LabelHook, BgControlName

'------------------------------------------------------------------------------
' CREATE / FORMAT TEXT LABEL
'------------------------------------------------------------------------------
            'Create or retrieve the day text label
                Set Lbl_Day = UF_Ensure_Label(TextControlName)
            'Cache the day text label for fast reuse
                Set mDayTextLabels(Index) = Lbl_Day
            'Apply layout and visual properties to the day text label
                With Lbl_Day
                    .Caption = VBA.CStr(Index)
                    .Left = TextLeft
                    .Top = TextTop
                    .Width = DP_DAY_LABEL_WIDTH
                    .Height = DP_DAY_LABEL_HEIGHT
                    .BackStyle = fmBackStyleTransparent
                    .BackColor = DP_DAY_NORMAL_BACK_COLOR
                    .ForeColor = DP_DAY_NORMAL_FORE_COLOR
                    .BorderStyle = fmBorderStyleNone
                    .BorderColor = DP_DAY_NORMAL_BACK_COLOR
                    .SpecialEffect = fmSpecialEffectFlat
                    .TextAlign = fmTextAlignCenter
                    .WordWrap = False
                    .Enabled = True
                    .Visible = True
                End With
            'Create the text-label event hook
                Set LabelHook = New cDatePickerLabelHook
            'Connect the day text label to its routed action
                LabelHook.Initialize Me, Lbl_Day, ActionName, "DAY"
            'Store the hook so that its events remain alive
                mDayLabelHooks.Add LabelHook, TextControlName

'------------------------------------------------------------------------------
' RESTORE LAYERING
'------------------------------------------------------------------------------
            'Move the background label behind the text label
                Lbl_DayBg.ZOrder 1
            'Move the text label to the front
                Lbl_Day.ZOrder 0

        Next Index

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
        Err.Raise Err.Number, PROC_NAME, "Day grid creation failed: " & Err.Description

End Sub

Private Sub UF_DayGrid_BuildIndexMap()

'
'------------------------------------------------------------------------------
'                           BUILD DAY GRID INDEX MAP
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the lookup map used to resolve day-label names to day-cell indexes
'
' WHY THIS EXISTS
'   Day cells are represented by two runtime labels:
'     - Lbl_Day1 to Lbl_Day42
'     - Lbl_DayBg1 to Lbl_DayBg42
'
'   Hover and click routing need to resolve either label name to the same
'   numeric day-cell index
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates a case-insensitive dictionary and maps both text-label and
'   background-label names to their day-cell index
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   Scripting.Dictionary through late binding
'   DP_DAY_LABEL_COUNT
'
' NOTES
'   This routine should be called once when the day grid is built, not inside
'   the 42-cell creation loop
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index       As Long     'Sequential day-cell index

'------------------------------------------------------------------------------
' INITIALIZE MAP
'------------------------------------------------------------------------------
    'Create the day-cell index map
        Set mDayCellIndexMap = CreateObject("Scripting.Dictionary")
    'Make label-name lookup case-insensitive
        mDayCellIndexMap.CompareMode = vbTextCompare

'------------------------------------------------------------------------------
' POPULATE MAP
'------------------------------------------------------------------------------
    'Loop through all day-cell indexes
        For Index = 1 To DP_DAY_LABEL_COUNT
            'Map the day text label to the day-cell index
                mDayCellIndexMap.Add "Lbl_Day" & VBA.CStr(Index), Index
            'Map the day background label to the day-cell index
                mDayCellIndexMap.Add "Lbl_DayBg" & VBA.CStr(Index), Index
        Next Index

End Sub

Private Sub UF_DayGrid_EnsureCache()

'
'------------------------------------------------------------------------------
'                           ENSURE DAY GRID CACHE
'------------------------------------------------------------------------------
' PURPOSE
'   Ensures the day-grid label cache is populated
'
' WHY THIS EXISTS
'   The DatePicker uses cached references to Lbl_Day1 to Lbl_Day42 and
'   Lbl_DayBg1 to Lbl_DayBg42 to avoid repeated Me.Controls string lookups
'   during grid refresh, hover, and selection handling
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Rebuilds missing cached references from the UserForm Controls collection
'
' ERROR POLICY
'   Raises a descriptive runtime error if one or more expected labels cannot be
'   resolved from the UserForm
'
' DEPENDENCIES
'   UserForm.Controls collection
'   Lbl_Day1 to Lbl_Day42
'   Lbl_DayBg1 to Lbl_DayBg42
'
' NOTES
'   This routine does not create controls. Controls are created by
'   UF_DayGrid_Build
'
'   Developer-owned grid constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_DayGrid_EnsureCache"

    Dim Index               As Long             'Sequential day-cell index
    Dim TextControlName     As String           'Runtime day text label name
    Dim BgControlName       As String           'Runtime day background label name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REBUILD MISSING CACHE REFERENCES
'------------------------------------------------------------------------------
    'Loop through all day cells
        For Index = 1 To DP_DAY_LABEL_COUNT
            'Build the runtime day text label name
                TextControlName = "Lbl_Day" & VBA.CStr(Index)
            'Build the runtime day background label name
                BgControlName = "Lbl_DayBg" & VBA.CStr(Index)

'------------------------------------------------------------------------------
' RESOLVE DAY TEXT LABEL
'------------------------------------------------------------------------------
            'Resolve the cached day text label when missing
                If mDayTextLabels(Index) Is Nothing Then
                    'Suppress missing-control lookup errors
                        On Error Resume Next
                    'Resolve the day text label from the UserForm controls collection
                        Set mDayTextLabels(Index) = Me.Controls(TextControlName)
                    'Restore controlled error handling
                        On Error GoTo ErrorHandler
                    'Reject unresolved day text label
                        If mDayTextLabels(Index) Is Nothing Then
                            Err.Raise vbObjectError + 513, PROC_NAME, _
                                "Unable to resolve expected day text label " & TextControlName
                        End If
                End If

'------------------------------------------------------------------------------
' RESOLVE DAY BACKGROUND LABEL
'------------------------------------------------------------------------------
            'Resolve the cached day background label when missing
                If mDayBackLabels(Index) Is Nothing Then
                    'Suppress missing-control lookup errors
                        On Error Resume Next
                    'Resolve the day background label from the UserForm controls collection
                        Set mDayBackLabels(Index) = Me.Controls(BgControlName)
                    'Restore controlled error handling
                        On Error GoTo ErrorHandler
                    'Reject unresolved day background label
                        If mDayBackLabels(Index) Is Nothing Then
                            Err.Raise vbObjectError + 514, PROC_NAME, _
                                "Unable to resolve expected day background label " & BgControlName
                        End If
                End If
        Next Index

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
        Err.Raise Err.Number, PROC_NAME, "Day grid cache initialization failed: " & Err.Description

End Sub

Private Sub UF_DayGrid_EnsureFonts()

'
'------------------------------------------------------------------------------
'                           ENSURE DAY GRID FONTS
'------------------------------------------------------------------------------
' PURPOSE
'   Creates cached font objects used by the DatePicker day grid
'
' WHY THIS EXISTS
'   Recreating a StdFont object for every day label during each grid population
'   is unnecessary. The day grid only needs two deterministic font states:
'     - normal day font
'     - weekend day font
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits immediately when both cached fonts already exist. Otherwise, creates
'   the missing cached normal and weekend day fonts
'
' ERROR POLICY
'   Clears the font cache and raises a descriptive runtime error if the cache
'   cannot be initialized
'
' DEPENDENCIES
'   StdFont
'   DP_FORM_FONT_NAME
'   DP_FORM_FONT_SIZE
'
' NOTES
'   The cached fonts intentionally reset Italic, Underline, and Strikethrough so
'   reused runtime labels do not inherit stale font states
'
'   Developer-owned font constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_DayGrid_EnsureFonts"
    
    Dim ErrorNumber         As Long         'Original runtime error number
    Dim ErrorDescription    As String       'Original runtime error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' EXIT IF CACHE IS READY
'------------------------------------------------------------------------------
    'Exit immediately when both cached fonts already exist
        If Not mDayFontNormal Is Nothing Then
            If Not mDayFontWeekend Is Nothing Then Exit Sub
        End If

'------------------------------------------------------------------------------
' CREATE NORMAL DAY FONT
'------------------------------------------------------------------------------
    'Create the normal day font when missing
        If mDayFontNormal Is Nothing Then
            'Create a clean normal day font object
                Set mDayFontNormal = CreateObject("StdFont")
            'Configure the normal day font
                With mDayFontNormal
                    .Name = DP_FORM_FONT_NAME
                    .Size = DP_FORM_FONT_SIZE
                    .Bold = False
                    .Italic = False
                    .Underline = False
                    .Strikethrough = False
                End With
        End If

'------------------------------------------------------------------------------
' CREATE WEEKEND DAY FONT
'------------------------------------------------------------------------------
    'Create the weekend day font when missing
        If mDayFontWeekend Is Nothing Then
            'Create a clean weekend day font object
                Set mDayFontWeekend = CreateObject("StdFont")
            'Configure the weekend day font
                With mDayFontWeekend
                    .Name = DP_FORM_FONT_NAME
                    .Size = DP_FORM_FONT_SIZE
                    .Bold = True
                    .Italic = False
                    .Underline = False
                    .Strikethrough = False
                End With
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
    'Store the original error number
        ErrorNumber = Err.Number
    'Store the original error description
        ErrorDescription = Err.Description
    'Suppress cleanup errors
        On Error Resume Next
    'Clear the normal day font cache
        Set mDayFontNormal = Nothing
    'Clear the weekend day font cache
        Set mDayFontWeekend = Nothing
    'Restore normal error handling
        On Error GoTo 0
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, "Day grid font cache initialization failed: " & ErrorDescription

End Sub

Public Sub UF_DayGrid_Populate( _
    Optional ByVal DisplayYear As Long = 0, _
    Optional ByVal DisplayMonth As Long = 0)

'
'------------------------------------------------------------------------------
'                           POPULATE DAY LABELS GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Populates the 42 DatePicker day cells for the requested month and year
'
' WHY THIS EXISTS
'   The DatePicker uses a fixed 6-row by 7-column calendar grid. Each visible
'   cell must be populated with the correct date, including leading and trailing
'   days from adjacent months
'
' INPUTS
'   DisplayYear
'     Year to display. If zero, the current system year is used
'
'   DisplayMonth
'     Month to display. If zero, the current system month is used
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves and validates the requested display period, updates the stored
'   display state, refreshes the header captions, populates cached day-cell
'   dates, formats the paired day labels, applies date-driven visual state, and
'   assigns cached normal or weekend fonts
'
' ERROR POLICY
'   Raises a descriptive runtime error if the requested month/year is invalid,
'   if the first-day setting is invalid, or if required day-grid controls are not
'   available
'
' DEPENDENCIES
'   UF_DayGrid_EnsureCache
'   UF_DayGrid_EnsureFonts
'   M_Settings_IsValidFirstDayOfWeek
'   M_Caption_GetMonth
'   UF_DayCell_ApplyDateStateByIndex
'
' NOTES
'   Cached StdFont objects are used to avoid repeated font-object creation while
'   still enforcing deterministic normal and weekend font states
'
'   Developer-owned grid layout constants are intentionally not validated here.
'   They should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_DayGrid_Populate"
    Const VBA_DATE_MIN_YEAR     As Long = 100           'Minimum year supported by VBA Date
    Const VBA_DATE_MAX_YEAR     As Long = 9999          'Maximum year supported by VBA Date

    Dim EffectiveYear           As Long                 'Validated display year
    Dim EffectiveMonth          As Long                 'Validated display month
    Dim EffectiveFirstDay       As VbDayOfWeek          'Validated first day of week
    Dim StartOfMonth            As Date                 'First day of display month
    Dim TrackingDate            As Date                 'Current date written to the grid
    Dim StartOfMonthDay         As Long                 'Weekday index of first month day
    Dim LabelIndex              As Long                 'Sequential day-cell index
    
    Dim CellDateOnly            As Date                 'Date-only value for the current grid cell
    Dim TodayDateOnly           As Date                 'Current system date without time
    Dim SelectedVisualDateOnly  As Date                 'Date-only selected visual date
    Dim HasSelectedVisualDate   As Boolean              'True when selected visual date exists
    Dim HighlightWeekends       As Boolean              'Cached weekend-highlight setting
    
    Dim Lbl_Day                 As MSForms.Label        'Day text label control
    Dim Lbl_DayBg               As MSForms.Label        'Day background label control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Ensure cached day-label references are available
        UF_DayGrid_EnsureCache
    'Ensure cached day-grid fonts are available
        UF_DayGrid_EnsureFonts
       
'------------------------------------------------------------------------------
' CACHE RENDER SETTINGS
'------------------------------------------------------------------------------
    'Resolve today's date once for the full grid refresh
        TodayDateOnly = VBA.Date
    'Cache the weekend-highlight setting for the full grid refresh
        HighlightWeekends = gDP_HighlightWeekends
    'Use the keyboard date as the selected visual date when available
        If mHasKeyboardDate Then
            SelectedVisualDateOnly = VBA.DateSerial( _
                VBA.Year(mKeyboardDate), _
                VBA.Month(mKeyboardDate), _
                VBA.Day(mKeyboardDate))
            HasSelectedVisualDate = True
        ElseIf gDP_HasSelectedDate Then
            SelectedVisualDateOnly = VBA.DateSerial( _
                VBA.Year(gDP_SelectedDate), _
                VBA.Month(gDP_SelectedDate), _
                VBA.Day(gDP_SelectedDate))
            HasSelectedVisualDate = True
        End If

'------------------------------------------------------------------------------
' RESOLVE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Use the current year when no display year is supplied
        If DisplayYear = 0 Then
            EffectiveYear = VBA.Year(VBA.Date)
        Else
            EffectiveYear = DisplayYear
        End If
    'Use the current month when no display month is supplied
        If DisplayMonth = 0 Then
            EffectiveMonth = VBA.Month(VBA.Date)
        Else
            EffectiveMonth = DisplayMonth
        End If

'------------------------------------------------------------------------------
' VALIDATE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Reject invalid display years
        If EffectiveYear < DP_MIN_YEAR Or EffectiveYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DisplayYear must be between " & VBA.CStr(DP_MIN_YEAR) & " and " & VBA.CStr(DP_MAX_YEAR)
        End If
    'Reject invalid display months
        If EffectiveMonth < 1 Or EffectiveMonth > 12 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "DisplayMonth must be between 1 and 12"
        End If
    'Reject display periods that cannot render a complete leading grid
        If EffectiveYear = VBA_DATE_MIN_YEAR And EffectiveMonth = 1 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "January " & VBA.CStr(VBA_DATE_MIN_YEAR) & " cannot render a complete DatePicker grid"
        End If
    'Reject display periods that cannot render a complete trailing grid
        If EffectiveYear = VBA_DATE_MAX_YEAR And EffectiveMonth = 12 Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "December " & VBA.CStr(VBA_DATE_MAX_YEAR) & " cannot render a complete DatePicker grid"
        End If

'------------------------------------------------------------------------------
' VALIDATE SETTINGS
'------------------------------------------------------------------------------
    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(gDP_FirstDayOfWeek) Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "gDP_FirstDayOfWeek must be vbSunday or vbMonday"
        End If
    'Store the validated first-day setting
        EffectiveFirstDay = gDP_FirstDayOfWeek

'------------------------------------------------------------------------------
' STORE DISPLAY STATE
'------------------------------------------------------------------------------
    'Store the currently displayed year
        mDisplayYear = EffectiveYear
    'Store the currently displayed month
        mDisplayMonth = EffectiveMonth

'------------------------------------------------------------------------------
' INITIALIZE CALENDAR DATES
'------------------------------------------------------------------------------
    'Compute the first day of the displayed month
        StartOfMonth = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
    'Resolve the weekday position of the first day of the displayed month
        StartOfMonthDay = VBA.Weekday(StartOfMonth, EffectiveFirstDay)
    'Back up to the first visible date in the calendar grid
        TrackingDate = VBA.DateAdd("d", -StartOfMonthDay + 1, StartOfMonth)

'------------------------------------------------------------------------------
' REFRESH HEADER CAPTIONS
'------------------------------------------------------------------------------
    'Update the header month caption
        Me.Controls("Lbl_HeaderMonth").Caption = _
            M_Caption_GetMonth(mDisplayMonth, gDP_UseLocalNames)
    'Update the header year caption
        Me.Controls("Lbl_HeaderYear").Caption = VBA.CStr(mDisplayYear)

'------------------------------------------------------------------------------
' POPULATE DAY LABELS
'------------------------------------------------------------------------------
    'Loop through the visible calendar cells
        For LabelIndex = 1 To DP_DAY_LABEL_COUNT
            'Resolve the date-only value represented by this cell
                CellDateOnly = TrackingDate
            'Cache the date represented by this day cell
                mDayCellDates(LabelIndex) = CellDateOnly
            'Mark the cached day-cell date as available
                mDayCellHasDate(LabelIndex) = True
            'Retrieve the cached day text label
                Set Lbl_Day = mDayTextLabels(LabelIndex)
            'Retrieve the cached day background label
                Set Lbl_DayBg = mDayBackLabels(LabelIndex)
            'Reject missing cached day text label
                If Lbl_Day Is Nothing Then
                    Err.Raise vbObjectError + 518, PROC_NAME, _
                        "Cached day text label is missing for index " & VBA.CStr(LabelIndex)
                End If
            'Reject missing cached day background label
                If Lbl_DayBg Is Nothing Then
                    Err.Raise vbObjectError + 519, PROC_NAME, _
                        "Cached day background label is missing for index " & VBA.CStr(LabelIndex)
                End If
            'Apply date-driven visual state using already-resolved labels
                UF_DayCell_ApplyDateStateFast _
                    CellDateOnly, _
                    Lbl_Day, _
                    Lbl_DayBg, _
                    mDisplayYear, _
                    mDisplayMonth, _
                    TodayDateOnly, _
                    HasSelectedVisualDate, _
                    SelectedVisualDateOnly, _
                    HighlightWeekends, _
                    mDayFontNormal, _
                    mDayFontWeekend
            'Advance to the next visible date
                TrackingDate = TrackingDate + 1
        Next LabelIndex

'------------------------------------------------------------------------------
' RESET HOVER STATE
'------------------------------------------------------------------------------
    'Clear any previous day-cell hover index after repopulating the grid
        mHoveredDayCellIndex = 0

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
        Err.Raise Err.Number, PROC_NAME, "Day grid population failed: " & Err.Description

End Sub



Private Sub UF_DayGrid_ClearCache()

'
'------------------------------------------------------------------------------
'                           CLEAR DAY GRID CACHE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears cached day-grid label references
'
' WHY THIS EXISTS
'   Cached MSForms control references should be released when the UserForm is
'   terminating or when the day-grid cache must be rebuilt from the controls
'   collection
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears cached references to day text labels and day background labels
'
' ERROR POLICY
'   Best-effort cleanup. Suppresses cleanup errors because this routine is
'   normally called during form teardown
'
' DEPENDENCIES
'   mDayTextLabels
'   mDayBackLabels
'
' NOTES
'   This routine does not delete controls
'
'   This routine does not clear cached day-cell date state.
'
'   The explicit loop is intentional. It preserves array dimensions and is safe
'   for both fixed-size arrays and already-dimensioned dynamic arrays
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index       As Long     'Sequential day-cell index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' CLEAR CACHED LABEL REFERENCES
'------------------------------------------------------------------------------
    'Loop through all cached day cells
        For Index = 1 To DP_DAY_LABEL_COUNT
            'Clear the cached day text label
                Set mDayTextLabels(Index) = Nothing
            'Clear the cached day background label
                Set mDayBackLabels(Index) = Nothing
        Next Index
    'Release the day-cell index map
        Set mDayCellIndexMap = Nothing
        
'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub


'
'------------------------------------------------------------------------------
'
'                               DAY CELL STATE AND HOVER
'
'------------------------------------------------------------------------------
'

Private Sub UF_DayCell_ApplyDateStateByIndex( _
    ByVal DayIndex As Long, _
    ByVal CellDate As Date)

'
'------------------------------------------------------------------------------
'                       APPLY DAY CELL DATE STATE BY INDEX
'------------------------------------------------------------------------------
' PURPOSE
'   Applies date-driven visual state to one paired DatePicker day cell by index
'
' WHY THIS EXISTS
'   Some routines know only the day-cell index, not the resolved label
'   references. This wrapper resolves the paired labels from cache and delegates
'   rendering to the fast internal day-cell state routine
'
' INPUTS
'   DayIndex
'     Day-cell index from 1 to DP_DAY_LABEL_COUNT
'
'   CellDate
'     Date represented by the day cell
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates DayIndex, restores cached labels when needed, normalizes date
'   inputs, resolves selected visual state, and delegates to
'   UF_DayCell_ApplyDateStateFast
'
' ERROR POLICY
'   Raises a descriptive runtime error if DayIndex is invalid, cached labels
'   cannot be resolved, or the day-cell state cannot be applied
'
' DEPENDENCIES
'   UF_DayGrid_EnsureCache
'   UF_DayGrid_EnsureFonts
'   UF_DayCell_ApplyDateStateFast
'   mDayTextLabels
'   mDayBackLabels
'
' NOTES
'   This routine preserves the safer index-based API for hover reset and
'   targeted refresh paths
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_DayCell_ApplyDateStateByIndex"

    Dim CellDateOnly            As Date             'Date-only value represented by the cell
    Dim TodayDateOnly           As Date             'Current system date without time
    Dim SelectedVisualDateOnly  As Date             'Date-only selected visual date
    Dim HasSelectedVisualDate   As Boolean          'True when selected visual date exists
    Dim CacheRefreshRequired    As Boolean          'True when cached labels must be rebuilt
    Dim Lbl_Text                As MSForms.Label    'Day text label
    Dim Lbl_Bg                  As MSForms.Label    'Day background label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject invalid day-cell indexes
        If DayIndex < 1 Or DayIndex > DP_DAY_LABEL_COUNT Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DayIndex must be between 1 and " & VBA.CStr(DP_DAY_LABEL_COUNT)
        End If

'------------------------------------------------------------------------------
' ENSURE FONT CACHE
'------------------------------------------------------------------------------
    'Ensure cached day-grid fonts are available
        UF_DayGrid_EnsureFonts

'------------------------------------------------------------------------------
' RETRIEVE CACHED LABELS
'------------------------------------------------------------------------------
    'Suppress cache probing errors
        On Error Resume Next
    'Retrieve the cached day text label
        Set Lbl_Text = mDayTextLabels(DayIndex)
    'Retrieve the cached day background label
        Set Lbl_Bg = mDayBackLabels(DayIndex)
    'Clear any suppressed cache probing error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESTORE CACHE WHEN NEEDED
'------------------------------------------------------------------------------
    'Request cache refresh when the text label is missing
        If Lbl_Text Is Nothing Then CacheRefreshRequired = True
    'Request cache refresh when the background label is missing
        If Lbl_Bg Is Nothing Then CacheRefreshRequired = True
    'Rebuild missing cached references only when needed
        If CacheRefreshRequired Then
            UF_DayGrid_EnsureCache
            Set Lbl_Text = mDayTextLabels(DayIndex)
            Set Lbl_Bg = mDayBackLabels(DayIndex)
        End If

'------------------------------------------------------------------------------
' VALIDATE CACHED LABELS
'------------------------------------------------------------------------------
    'Reject a missing cached text label
        If Lbl_Text Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Cached text label is missing for DayIndex " & VBA.CStr(DayIndex)
        End If
    'Reject a missing cached background label
        If Lbl_Bg Is Nothing Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Cached background label is missing for DayIndex " & VBA.CStr(DayIndex)
        End If

'------------------------------------------------------------------------------
' NORMALIZE DATES
'------------------------------------------------------------------------------
    'Normalize the cell date to a date-only value
        CellDateOnly = VBA.DateSerial(VBA.Year(CellDate), VBA.Month(CellDate), VBA.Day(CellDate))
    'Resolve today's date once
        TodayDateOnly = VBA.Date

'------------------------------------------------------------------------------
' RESOLVE SELECTED VISUAL DATE
'------------------------------------------------------------------------------
    'Use the keyboard date as the selected visual date when available
        If mHasKeyboardDate Then
            SelectedVisualDateOnly = VBA.DateSerial( _
                VBA.Year(mKeyboardDate), _
                VBA.Month(mKeyboardDate), _
                VBA.Day(mKeyboardDate))
            HasSelectedVisualDate = True
        ElseIf gDP_HasSelectedDate Then
            SelectedVisualDateOnly = VBA.DateSerial( _
                VBA.Year(gDP_SelectedDate), _
                VBA.Month(gDP_SelectedDate), _
                VBA.Day(gDP_SelectedDate))
            HasSelectedVisualDate = True
        End If

'------------------------------------------------------------------------------
' APPLY FAST STATE
'------------------------------------------------------------------------------
    'Apply the day-cell state using already-resolved labels
        UF_DayCell_ApplyDateStateFast _
            CellDateOnly, _
            Lbl_Text, _
            Lbl_Bg, _
            mDisplayYear, _
            mDisplayMonth, _
            TodayDateOnly, _
            HasSelectedVisualDate, _
            SelectedVisualDateOnly, _
            gDP_HighlightWeekends, _
            mDayFontNormal, _
            mDayFontWeekend

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
        Err.Raise Err.Number, PROC_NAME, "Day-cell state application failed: " & Err.Description

End Sub


Private Sub UF_DayCell_RefreshVisibleDate(ByVal TargetDate As Date)

'
'------------------------------------------------------------------------------
'                    REFRESH VISIBLE DAY CELL BY DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Re-applies date-driven visual state to one visible day cell matching a date
'
' WHY THIS EXISTS
'   When only one date needs to be refreshed, scanning cached day-cell dates is
'   faster and safer than reading MSForms Label.Tag values
'
' INPUTS
'   TargetDate
'     Date to locate in the currently visible calendar grid
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Scans cached visible day-cell dates and refreshes the first cell matching
'   TargetDate. If TargetDate is not visible in the current calendar grid, the
'   routine exits without changing anything
'
' ERROR POLICY
'   Raises a descriptive runtime error if the visible-date scan or matching cell
'   refresh fails
'
' DEPENDENCIES
'   mDayCellDates
'   mDayCellHasDate
'   UF_DayCell_ApplyDateStateByIndex
'
' NOTES
'   This routine intentionally does not rebuild the day-grid cache. It relies on
'   the current populated-grid date cache
'
'   Not finding TargetDate is an expected condition and is not treated as an
'   error
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_DayCell_RefreshVisibleDate"
    
    Dim Index               As Long         'Day-cell index
    Dim TargetDateOnly      As Date         'Target date without time
    Dim CellDateOnly        As Date         'Cached cell date without time

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Resolve the target date without time
        TargetDateOnly = VBA.DateValue(TargetDate)

'------------------------------------------------------------------------------
' SCAN CACHED DAY CELLS
'------------------------------------------------------------------------------
    'Loop through the visible calendar grid
        For Index = 1 To DP_DAY_LABEL_COUNT
            'Evaluate only populated cached day cells
                If mDayCellHasDate(Index) Then
                    'Resolve the cached cell date without time
                        CellDateOnly = VBA.DateValue(mDayCellDates(Index))
                    'Refresh the matching day cell
                        If CellDateOnly = TargetDateOnly Then
                            UF_DayCell_ApplyDateStateByIndex Index, mDayCellDates(Index)
                            Exit Sub
                        End If
                End If
        Next Index

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit when the target date is not visible
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, "Visible day-cell refresh failed: " & Err.Description

End Sub

Public Sub UF_DayCell_HoverApply(ByVal LabelName As String)

'
'------------------------------------------------------------------------------
'                           APPLY DAY LABEL HOVER
'------------------------------------------------------------------------------
' PURPOSE
'   Applies hover formatting to one paired DatePicker day cell
'
' WHY THIS EXISTS
'   The DatePicker grid should provide a larger hover target than the visible
'   day-number label. Each day cell therefore uses a background label and a text
'   label that resolve to the same day-cell index
'
' INPUTS
'   LabelName
'     Name of the day text or background label currently under the mouse pointer
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the day-cell index, exits when the same day cell is already hovered,
'   resets the previous hover state, retrieves the paired labels from cache,
'   resolves cached date state, and applies hover formatting when the represented
'   date is selectable
'
' ERROR POLICY
'   Raises a descriptive runtime error if LabelName is empty, unsupported, or if
'   the target day cell cannot be resolved or formatted
'
' DEPENDENCIES
'   UF_DayCell_GetIndexFromLabelName
'   UF_DayCell_HoverReset
'   UF_DayGrid_EnsureCache
'   M_DatePolicy_CanSelectDate
'   mDayTextLabels
'   mDayBackLabels
'   mDayCellDates
'   mDayCellHasDate
'
' NOTES
'   Hover formatting is applied primarily to the background label
'
'   This routine does not call ZOrder. Layering is established when the grid is
'   built and should not be repeated during high-frequency hover handling
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_DayCell_HoverApply"

    Dim DayIndex                As Long                 'Resolved day-cell index
    Dim LabelDate               As Date                 'Date represented by the label
    Dim LabelDateOnly           As Date                 'Date-only value represented by the label
    Dim HasLabelDate            As Boolean              'True when cached date state is available
    Dim IsSelected              As Boolean              'True when the day cell is selected
    Dim IsOutsideMonth          As Boolean              'True when the day cell belongs to an adjacent month
    Dim IsSelectable            As Boolean              'True when the represented date can be selected
    Dim CacheRefreshRequired    As Boolean              'True when cached labels must be rebuilt
    Dim Lbl_Text                As MSForms.Label        'Day text label
    Dim Lbl_Bg                  As MSForms.Label        'Day background label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty label name
        If Len(VBA.Trim$(LabelName)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "LabelName cannot be empty"
        End If

'------------------------------------------------------------------------------
' RESOLVE DAY CELL
'------------------------------------------------------------------------------
    'Resolve the paired day-cell index
        DayIndex = UF_DayCell_GetIndexFromLabelName(LabelName)
    'Reject invalid day-cell indexes
        If DayIndex < 1 Or DayIndex > DP_DAY_LABEL_COUNT Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "LabelName does not resolve to a valid day-cell index"
        End If

'------------------------------------------------------------------------------
' EXIT IF ALREADY HOVERED
'------------------------------------------------------------------------------
    'Exit if the requested day cell is already highlighted
        If mHoveredDayCellIndex = DayIndex Then Exit Sub

'------------------------------------------------------------------------------
' RESET CURRENT HOVER STATE
'------------------------------------------------------------------------------
    'Remove hover formatting from the previously highlighted day cell
        If mHoveredDayCellIndex <> 0 Then
            UF_DayCell_HoverReset
        End If

'------------------------------------------------------------------------------
' RETRIEVE CACHED LABELS
'------------------------------------------------------------------------------
    'Suppress cache probing errors
        On Error Resume Next
    'Retrieve the cached day text label
        Set Lbl_Text = mDayTextLabels(DayIndex)
    'Retrieve the cached day background label
        Set Lbl_Bg = mDayBackLabels(DayIndex)
    'Clear any suppressed cache probing error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESTORE CACHE WHEN NEEDED
'------------------------------------------------------------------------------
    'Request cache refresh when the text label is missing
        If Lbl_Text Is Nothing Then CacheRefreshRequired = True
    'Request cache refresh when the background label is missing
        If Lbl_Bg Is Nothing Then CacheRefreshRequired = True
    'Rebuild missing cached references only when needed
        If CacheRefreshRequired Then
            'Ensure cached day-label references are available
                UF_DayGrid_EnsureCache
            'Retrieve the cached day text label after cache refresh
                Set Lbl_Text = mDayTextLabels(DayIndex)
            'Retrieve the cached day background label after cache refresh
                Set Lbl_Bg = mDayBackLabels(DayIndex)
        End If

'------------------------------------------------------------------------------
' VALIDATE CACHED LABELS
'------------------------------------------------------------------------------
    'Reject a missing cached text label
        If Lbl_Text Is Nothing Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Cached text label is missing for DayIndex " & VBA.CStr(DayIndex)
        End If
    'Reject a missing cached background label
        If Lbl_Bg Is Nothing Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Cached background label is missing for DayIndex " & VBA.CStr(DayIndex)
        End If

'------------------------------------------------------------------------------
' RESOLVE LABEL DATE STATE
'------------------------------------------------------------------------------
    'Resolve the label date from the cached day-cell date
        If mDayCellHasDate(DayIndex) Then
            LabelDate = mDayCellDates(DayIndex)
            LabelDateOnly = VBA.DateValue(LabelDate)
            HasLabelDate = True
        End If
    'Exit when the day cell has not yet been populated with a date
        If Not HasLabelDate Then Exit Sub
    'Resolve whether the represented date can be selected
        IsSelectable = M_DatePolicy_CanSelectDate(LabelDateOnly, mDisplayYear, mDisplayMonth)
    'Exit when the represented date is not selectable
        If Not IsSelectable Then Exit Sub
    'Resolve whether the hovered day cell is the selected date
        If mHasKeyboardDate Then
            IsSelected = (LabelDateOnly = VBA.DateValue(mKeyboardDate))
        ElseIf gDP_HasSelectedDate Then
            IsSelected = (LabelDateOnly = VBA.DateValue(gDP_SelectedDate))
        Else
            IsSelected = False
        End If
    'Resolve whether the hovered day cell belongs to an adjacent month
        IsOutsideMonth = _
            (VBA.Year(LabelDateOnly) <> mDisplayYear Or VBA.Month(LabelDateOnly) <> mDisplayMonth)

'------------------------------------------------------------------------------
' APPLY SELECTED HOVER STATE
'------------------------------------------------------------------------------
    'Preserve selected-date colors when hovering the selected day cell
        If IsSelected Then
            With Lbl_Bg
                .BackStyle = fmBackStyleOpaque
                .BackColor = DP_DAY_SELECTED_BACK_COLOR
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = DP_DAY_SELECTED_BACK_COLOR
                .SpecialEffect = fmSpecialEffectFlat
            End With
            'Apply selected text formatting
                Lbl_Text.ForeColor = DP_DAY_SELECTED_FORE_COLOR
            'Store the current hovered day-cell index
                mHoveredDayCellIndex = DayIndex
            'Exit after applying selected hover state
                Exit Sub
        End If

'------------------------------------------------------------------------------
' APPLY NORMAL HOVER STATE
'------------------------------------------------------------------------------
    'Apply standard hover background formatting
        With Lbl_Bg
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_DAY_HOVER_BACK_COLOR
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = DP_DAY_HOVER_BORDER_COLOR
            .SpecialEffect = fmSpecialEffectFlat
        End With
    'Restore readable text color on light hover background
        If IsOutsideMonth Then
            Lbl_Text.ForeColor = DP_DAY_OUTSIDE_MONTH_FORE_COLOR
        Else
            Lbl_Text.ForeColor = DP_DAY_CURRENT_MONTH_FORE_COLOR
        End If

'------------------------------------------------------------------------------
' STORE HOVER STATE
'------------------------------------------------------------------------------
    'Store the current hovered day-cell index
        mHoveredDayCellIndex = DayIndex

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
        Err.Raise Err.Number, PROC_NAME, "Day-cell hover application failed: " & Err.Description

End Sub

Private Sub UF_DayCell_ApplyDateStateFast( _
    ByVal CellDateOnly As Date, _
    ByVal Lbl_Text As MSForms.Label, _
    ByVal Lbl_Bg As MSForms.Label, _
    ByVal EffectiveDisplayYear As Long, _
    ByVal EffectiveDisplayMonth As Long, _
    ByVal TodayDateOnly As Date, _
    ByVal HasSelectedVisualDate As Boolean, _
    ByVal SelectedVisualDateOnly As Date, _
    ByVal HighlightWeekends As Boolean, _
    ByVal FontNormal As Object, _
    ByVal FontWeekend As Object)

'
'------------------------------------------------------------------------------
'                       APPLY DAY CELL DATE STATE FAST
'------------------------------------------------------------------------------
' PURPOSE
'   Applies date-driven visual state to one already-resolved paired day cell
'
' WHY THIS EXISTS
'   UF_DayGrid_Populate already resolves the paired text and background labels
'   while looping through the 42 visible day cells
'
'   Passing those references directly avoids repeated cache probing, repeated
'   Controls lookups, and repeated label-pair resolution during full grid refresh
'
' INPUTS
'   CellDateOnly
'     Date-only value represented by the day cell
'
'   Lbl_Text
'     Already-resolved day text label
'
'   Lbl_Bg
'     Already-resolved day background label
'
'   EffectiveDisplayYear
'     Displayed calendar year
'
'   EffectiveDisplayMonth
'     Displayed calendar month
'
'   TodayDateOnly
'     Current system date without time
'
'   HasSelectedVisualDate
'     True when a selected / keyboard date should be rendered
'
'   SelectedVisualDateOnly
'     Date-only selected / keyboard date when HasSelectedVisualDate is True
'
'   HighlightWeekends
'     True when weekend labels should use the weekend font
'
'   FontNormal
'     Cached normal day font
'
'   FontWeekend
'     Cached weekend day font
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Applies base state, outside-month state, today state, selected state,
'   selectable state, and optional weekend font state without resolving controls
'
' ERROR POLICY
'   Raises descriptive runtime errors when required label references are missing
'
' DEPENDENCIES
'   M_DatePolicy_CanSelectDate
'   DatePicker day-grid visual constants
'
' NOTES
'   This routine is intentionally private and assumes the caller has already
'   resolved the correct paired labels
'
'   It does not call UF_DayGrid_EnsureCache and does not perform any ZOrder work
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_DayCell_ApplyDateStateFast"

    Dim IsOutsideMonth      As Boolean      'True when the cell belongs to an adjacent month
    Dim IsToday             As Boolean      'True when the cell represents today
    Dim IsSelected          As Boolean      'True when the cell is the selected visual date
    Dim IsSelectable        As Boolean      'True when the date can be selected
    Dim WeekdaySunBasis     As Long         'Sunday-based weekday number
    Dim IsWeekend           As Boolean      'True when the cell date is Saturday or Sunday

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE REFERENCES
'------------------------------------------------------------------------------
    'Reject a missing text label
        If Lbl_Text Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "Lbl_Text cannot be Nothing"
        End If
    'Reject a missing background label
        If Lbl_Bg Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, "Lbl_Bg cannot be Nothing"
        End If

'------------------------------------------------------------------------------
' RESOLVE STATE
'------------------------------------------------------------------------------
    'Resolve whether the cell belongs to an adjacent month
        IsOutsideMonth = _
            (VBA.Year(CellDateOnly) <> EffectiveDisplayYear Or _
             VBA.Month(CellDateOnly) <> EffectiveDisplayMonth)
    'Resolve whether the cell represents today
        IsToday = (CellDateOnly = TodayDateOnly)
    'Resolve whether the cell represents the selected visual date
        If HasSelectedVisualDate Then
            IsSelected = (CellDateOnly = SelectedVisualDateOnly)
        End If
    'Resolve whether the date can be selected
        IsSelectable = M_DatePolicy_CanSelectDate( _
            CellDateOnly, _
            EffectiveDisplayYear, _
            EffectiveDisplayMonth)

'------------------------------------------------------------------------------
' APPLY BACKGROUND BASE STATE
'------------------------------------------------------------------------------
    'Apply normal background state
        With Lbl_Bg
            .Caption = vbNullString
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_DAY_NORMAL_BACK_COLOR
            .ForeColor = DP_DAY_CURRENT_MONTH_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .BorderColor = DP_DAY_NORMAL_BACK_COLOR
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = IsSelectable
            .Visible = True
        End With

'------------------------------------------------------------------------------
' APPLY TEXT BASE STATE
'------------------------------------------------------------------------------
    'Apply normal text state
        With Lbl_Text
            .Caption = VBA.CStr(VBA.Day(CellDateOnly))
            .BackStyle = fmBackStyleTransparent
            .BackColor = DP_DAY_NORMAL_BACK_COLOR
            .ForeColor = DP_DAY_CURRENT_MONTH_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .BorderColor = DP_DAY_NORMAL_BACK_COLOR
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = IsSelectable
            .Visible = True
        End With

'------------------------------------------------------------------------------
' APPLY WEEKEND FONT STATE
'------------------------------------------------------------------------------
    'Assign normal font unless weekend highlighting requires the weekend font
        If HighlightWeekends Then
            WeekdaySunBasis = VBA.Weekday(CellDateOnly, vbSunday)
            IsWeekend = (WeekdaySunBasis = vbSaturday Or WeekdaySunBasis = vbSunday)

            If IsWeekend Then
                If Not FontWeekend Is Nothing Then Set Lbl_Text.Font = FontWeekend
            Else
                If Not FontNormal Is Nothing Then Set Lbl_Text.Font = FontNormal
            End If
        Else
            If Not FontNormal Is Nothing Then Set Lbl_Text.Font = FontNormal
        End If

'------------------------------------------------------------------------------
' APPLY OUTSIDE-MONTH STATE
'------------------------------------------------------------------------------
    'Apply muted text for outside-month dates
        If IsOutsideMonth Then
            Lbl_Text.ForeColor = DP_DAY_OUTSIDE_MONTH_FORE_COLOR
        End If

'------------------------------------------------------------------------------
' APPLY TODAY STATE
'------------------------------------------------------------------------------
    'Apply today highlight when applicable
        If IsToday Then
            With Lbl_Bg
                .BackColor = DP_DAY_TODAY_BACK_COLOR
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = DP_DAY_TODAY_BORDER_COLOR
            End With
        End If

'------------------------------------------------------------------------------
' APPLY SELECTED STATE
'------------------------------------------------------------------------------
    'Apply selected-date highlight when applicable
        If IsSelected Then
            With Lbl_Bg
                .BackColor = DP_DAY_SELECTED_BACK_COLOR
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = DP_DAY_SELECTED_BACK_COLOR
            End With
            'Apply selected text formatting
                Lbl_Text.ForeColor = DP_DAY_SELECTED_FORE_COLOR
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
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, "Fast day-cell state application failed: " & Err.Description

End Sub
Private Sub UF_DayCell_HoverReset()

'
'------------------------------------------------------------------------------
'                           RESET DAY LABELS HOVER
'------------------------------------------------------------------------------
' PURPOSE
'   Removes hover formatting from the currently highlighted paired day cell
'
' WHY THIS EXISTS
'   Hover formatting must be removed when the mouse moves away from the day
'   cells and back onto another DatePicker surface
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Restores the currently highlighted day cell to its date-driven visual state
'   when the cached date is available. Applies a safe normal fallback when the
'   cell has no cached date state
'
' ERROR POLICY
'   Best-effort UI cleanup. Suppresses reset failures because hover reset is a
'   high-frequency visual operation and should never interrupt the user workflow
'
' DEPENDENCIES
'   UF_DayGrid_EnsureCache
'   UF_DayCell_ApplyDateStateByIndex
'   mDayTextLabels
'   mDayBackLabels
'   mDayCellDates
'   mDayCellHasDate
'
' NOTES
'   This routine resets only the currently highlighted day cell, not all 42 cells
'
'   The day-grid cache is rebuilt only when the requested cached label pair is
'   missing
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_DayCell_HoverReset"
    
    Dim HoverIndex              As Long                 'Previously hovered day-cell index
    Dim LabelDate               As Date                 'Date represented by the day cell
    Dim HasLabelDate            As Boolean              'True when cached date state is available
    Dim CacheRefreshRequired    As Boolean              'True when cached labels must be rebuilt
    Dim Lbl_Text                As MSForms.Label        'Previously highlighted day text label
    Dim Lbl_Bg                  As MSForms.Label        'Previously highlighted day background label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' EXIT IF NOTHING IS HOVERED
'------------------------------------------------------------------------------
    'Exit if no day cell is currently highlighted
        If mHoveredDayCellIndex = 0 Then Exit Sub
    'Store the hovered day-cell index before clearing state
        HoverIndex = mHoveredDayCellIndex

'------------------------------------------------------------------------------
' VALIDATE HOVER INDEX
'------------------------------------------------------------------------------
    'Clear hover state and exit when the stored hover index is no longer valid
        If HoverIndex < 1 Or HoverIndex > DP_DAY_LABEL_COUNT Then
            GoTo Clean_Exit
        End If

'------------------------------------------------------------------------------
' RETRIEVE CACHED LABELS
'------------------------------------------------------------------------------
    'Suppress cache probing errors
        On Error Resume Next
    'Retrieve the cached day text label
        Set Lbl_Text = mDayTextLabels(HoverIndex)
    'Retrieve the cached day background label
        Set Lbl_Bg = mDayBackLabels(HoverIndex)
    'Clear any suppressed cache probing error
        Err.Clear
    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' RESTORE CACHE WHEN NEEDED
'------------------------------------------------------------------------------
    'Request cache refresh when the text label is missing
        If Lbl_Text Is Nothing Then CacheRefreshRequired = True
    'Request cache refresh when the background label is missing
        If Lbl_Bg Is Nothing Then CacheRefreshRequired = True
    'Rebuild missing cached references only when needed
        If CacheRefreshRequired Then
            'Ensure cached day-label references are available
                UF_DayGrid_EnsureCache
            'Retrieve the cached day text label after cache refresh
                Set Lbl_Text = mDayTextLabels(HoverIndex)
            'Retrieve the cached day background label after cache refresh
                Set Lbl_Bg = mDayBackLabels(HoverIndex)
        End If

'------------------------------------------------------------------------------
' RESOLVE CACHED DATE STATE
'------------------------------------------------------------------------------
    'Resolve the cached date state when available
        If mDayCellHasDate(HoverIndex) Then
            'Read the cached day-cell date
                LabelDate = mDayCellDates(HoverIndex)
            'Mark the cached date as available
                HasLabelDate = True
        End If

'------------------------------------------------------------------------------
' RESTORE DATE-DRIVEN STATE
'------------------------------------------------------------------------------
    'Restore the full date-driven visual state when the cached date and labels are available
        If HasLabelDate Then
            If Not Lbl_Text Is Nothing Then
                If Not Lbl_Bg Is Nothing Then
                    'Restore the complete date-driven state
                        UF_DayCell_ApplyDateStateByIndex HoverIndex, LabelDate
                    'Continue to state cleanup
                        GoTo Clean_Exit
                End If
            End If
        End If

'------------------------------------------------------------------------------
' APPLY SAFE BACKGROUND FALLBACK
'------------------------------------------------------------------------------
    'Apply a safe normal visual fallback to the background label when available
        If Not Lbl_Bg Is Nothing Then
            With Lbl_Bg
                .BackStyle = fmBackStyleOpaque
                .BackColor = DP_DAY_NORMAL_BACK_COLOR
                .BorderStyle = fmBorderStyleNone
                .BorderColor = DP_DAY_NORMAL_BACK_COLOR
                .SpecialEffect = fmSpecialEffectFlat
                .Enabled = True
                .Visible = True
            End With
        End If

'------------------------------------------------------------------------------
' APPLY SAFE TEXT FALLBACK
'------------------------------------------------------------------------------
    'Apply a safe normal visual fallback to the text label when available
        If Not Lbl_Text Is Nothing Then
            With Lbl_Text
                .BackStyle = fmBackStyleTransparent
                .ForeColor = DP_DAY_NORMAL_FORE_COLOR
                .BorderStyle = fmBorderStyleNone
                .BorderColor = DP_DAY_NORMAL_BACK_COLOR
                .SpecialEffect = fmSpecialEffectFlat
                .Enabled = True
                .Visible = True
            End With
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:
    'Clear the current hovered day-cell index
        mHoveredDayCellIndex = 0
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Suppress secondary cleanup errors
        On Error Resume Next
    'Clear the current hovered day-cell index
        mHoveredDayCellIndex = 0
    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                               SHARED HELPERS
'
'------------------------------------------------------------------------------
'

Private Function UF_DayCell_GetIndexFromLabelName(ByVal LabelName As String) As Long

'
'------------------------------------------------------------------------------
'                   GET DAY CELL INDEX FROM LABEL NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the day-cell index from a DatePicker day label name
'
' WHY THIS EXISTS
'   Hovering or clicking either the day text label or the day background label
'   must resolve to the same day-cell index
'
' INPUTS
'   LabelName
'     Runtime day label name
'
' RETURNS
'   Day-cell index from 1 to DP_DAY_LABEL_COUNT
'   Zero when the label name is blank, unsupported, or not mapped
'
' BEHAVIOR
'   Uses the prebuilt case-insensitive mDayCellIndexMap for constant-time lookup
'
' ERROR POLICY
'   Does not raise errors
'
' DEPENDENCIES
'   mDayCellIndexMap
'   DP_DAY_LABEL_COUNT
'
' NOTES
'   The map is built once by UF_DayGrid_BuildIndexMap
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EffectiveName       As String       'Trimmed label name
    Dim ParsedIndex         As Long         'Mapped day-cell index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe lookup
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the supplied label name
        EffectiveName = VBA.Trim$(LabelName)
    'Return zero when the supplied label name is blank
        If VBA.Len(EffectiveName) = 0 Then Exit Function

'------------------------------------------------------------------------------
' VALIDATE MAP STATE
'------------------------------------------------------------------------------
    'Return zero when the day-cell index map is not available
        If mDayCellIndexMap Is Nothing Then Exit Function
    'Return zero when the label name is not mapped
        If Not mDayCellIndexMap.Exists(EffectiveName) Then Exit Function

'------------------------------------------------------------------------------
' RESOLVE INDEX
'------------------------------------------------------------------------------
    'Read the mapped day-cell index
        ParsedIndex = VBA.CLng(mDayCellIndexMap(EffectiveName))
    'Return zero when the mapped index is outside the supported range
        If ParsedIndex < 1 Or ParsedIndex > DP_DAY_LABEL_COUNT Then Exit Function

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return the validated day-cell index
        UF_DayCell_GetIndexFromLabelName = ParsedIndex

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit the function
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Return zero for missing, unsupported, or invalid label mappings
        UF_DayCell_GetIndexFromLabelName = 0

End Function

'
'------------------------------------------------------------------------------
'
'                                   FOOTER
'
'------------------------------------------------------------------------------
'

Private Sub UF_Footer_Build()

'
'------------------------------------------------------------------------------
'                           CREATE FOOTER SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Creates the DatePicker footer divider, footer banner, footer labels, and
'   footer settings entry point
'
' WHY THIS EXISTS
'   The footer area visually separates quick Today / Time shortcuts from the
'   calendar grid and exposes the settings entry point in the normal layout
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds or refreshes the full footer section in the required visual order:
'   divider, banner, footer action labels, and settings area
'
' ERROR POLICY
'   Raises a descriptive runtime error if footer creation fails
'
' DEPENDENCIES
'   UF_Footer_BuildDivider
'   UF_Footer_BuildBanner
'   UF_Footer_BuildLabels
'   UF_Footer_BuildSettingsArea
'
' NOTES
'   This routine coordinates footer creation only
'
'   In compact layout, the footer may not be visible inside the shortened form
'   canvas. Compact settings access is handled separately by the header settings
'   icon
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_DatePicker.UF_Footer_Build"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CREATE DIVIDER
'------------------------------------------------------------------------------
    'Create or format the divider between the day grid and the footer
        UF_Footer_BuildDivider

'------------------------------------------------------------------------------
' CREATE FOOTER BANNER
'------------------------------------------------------------------------------
    'Create or format the footer background banner
        UF_Footer_BuildBanner

'------------------------------------------------------------------------------
' CREATE FOOTER LABELS
'------------------------------------------------------------------------------
    'Create or format the footer action labels
        UF_Footer_BuildLabels

'------------------------------------------------------------------------------
' CREATE SETTINGS AREA
'------------------------------------------------------------------------------
    'Create or format the settings separator and icon
        UF_Footer_BuildSettingsArea

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
        Err.Raise Err.Number, PROC_NAME, "Footer creation failed: " & Err.Description

End Sub

Private Sub UF_Footer_BuildDivider()

'
'------------------------------------------------------------------------------
'                         BUILD FOOTER DIVIDER
'------------------------------------------------------------------------------
' PURPOSE
'   Creates and formats the divider label displayed between the calendar day
'   grid and the footer banner
'
' WHY THIS EXISTS
'   The divider visually separates the calendar grid from the footer area
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses Lbl_Divider and formats it as a passive decorative line
'   between the day grid and the footer section
'
' ERROR POLICY
'   Raises a descriptive runtime error if the divider cannot be created or
'   formatted
'
' DEPENDENCIES
'   UF_Ensure_Label
'   UF_CalendarGrid_GetBottom
'
' NOTES
'   The divider is implemented as a thin runtime-created MSForms.Label
'
'   The divider is sized to the DatePicker designed canvas width. Native-title-
'   bar compensation is handled by UF_Form_ApplyEffectiveSize and should not be
'   duplicated here
'
'   Developer-owned layout constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Footer_BuildDivider"

    Dim DividerTop          As Single               'Divider top position
    Dim DividerWidth        As Single               'Divider width
    Dim Lbl_Divider         As MSForms.Label        'Divider label control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CALCULATE LAYOUT
'------------------------------------------------------------------------------
    'Position the divider approximately in the middle of the footer gap
        DividerTop = UF_CalendarGrid_GetBottom() + _
            ((DP_FOOTER_TOP_GAP - DP_DIVIDER_HEIGHT) / 2)
    'Calculate the divider width
        DividerWidth = DP_FORM_WIDTH - (DP_DIVIDER_SIDE_MARGIN * 2)

'------------------------------------------------------------------------------
' CREATE / FORMAT DIVIDER
'------------------------------------------------------------------------------
    'Create or retrieve the divider label
        Set Lbl_Divider = UF_Ensure_Label("Lbl_Divider")
    'Apply layout and visual properties
        With Lbl_Divider
            .Caption = vbNullString
            .Left = DP_DIVIDER_SIDE_MARGIN
            .Top = DividerTop
            .Width = DividerWidth
            .Height = DP_DIVIDER_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_DIVIDER_COLOR
            .ForeColor = DP_DIVIDER_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = False
            .Visible = True
        End With
    'Bring the divider above passive background surfaces
        Lbl_Divider.ZOrder 0

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
        Err.Raise Err.Number, PROC_NAME, "Footer divider creation failed: " & Err.Description

End Sub

Private Sub UF_Footer_BuildBanner()

'
'------------------------------------------------------------------------------
'                         BUILD FOOTER BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Creates and formats the DatePicker footer background banner
'
' WHY THIS EXISTS
'   The footer banner visually groups the date and time shortcut area into a
'   distinct bottom section
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses Lbl_FooterBanner and formats it as the passive background
'   surface for the footer section
'
' ERROR POLICY
'   Raises a descriptive runtime error if the footer banner cannot be created or
'   formatted
'
' DEPENDENCIES
'   UF_Ensure_Label
'   UF_CalendarGrid_GetWidth
'   UF_CalendarGrid_GetBottom
'
' NOTES
'   The footer banner is implemented as a runtime-created MSForms.Label
'
'   The banner is sized to the DatePicker designed calendar-grid area. Native-
'   title-bar compensation is handled by UF_Form_ApplyEffectiveSize and should
'   not be duplicated here
'
'   Developer-owned layout constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Footer_BuildBanner"

    Dim BannerLeft          As Single               'Footer banner left position
    Dim BannerTop           As Single               'Footer banner top position
    Dim BannerWidth         As Single               'Footer banner width
    Dim Lbl_FooterBanner    As MSForms.Label        'Footer banner control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CALCULATE LAYOUT
'------------------------------------------------------------------------------
    'Calculate the banner size and position
        BannerLeft = DP_DAY_GRID_START_LEFT - DP_FOOTER_SIDE_PADDING
        BannerTop = UF_CalendarGrid_GetBottom() + DP_FOOTER_TOP_GAP
        BannerWidth = UF_CalendarGrid_GetWidth() + (DP_FOOTER_SIDE_PADDING * 2)

'------------------------------------------------------------------------------
' CREATE / FORMAT BANNER
'------------------------------------------------------------------------------
    'Create or retrieve the footer banner label
        Set Lbl_FooterBanner = UF_Ensure_Label("Lbl_FooterBanner")
    'Apply layout and visual properties
        With Lbl_FooterBanner
            .Caption = vbNullString
            .Left = BannerLeft
            .Top = BannerTop
            .Width = BannerWidth
            .Height = DP_FOOTER_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_FOOTER_BACK_COLOR
            .ForeColor = DP_FOOTER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = False
            .Visible = True
        End With
    'Send the banner behind footer action labels and settings controls
        Lbl_FooterBanner.ZOrder 1

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
        Err.Raise Err.Number, PROC_NAME, "Footer banner creation failed: " & Err.Description

End Sub

Private Sub UF_Footer_BuildLabels()

'
'------------------------------------------------------------------------------
'                           CREATE FOOTER LABELS
'------------------------------------------------------------------------------
' PURPOSE
'   Creates and formats the labels displayed on the DatePicker footer banner
'
' WHY THIS EXISTS
'   The footer banner requires visible and clickable labels for Today and Time
'   shortcut actions
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses the Today and Time halo labels, icon labels, caption
'   labels, and value labels. Registers footer click / hover hooks so all
'   visible Today surfaces write today's date and all visible Time surfaces write
'   the current date and time
'
' ERROR POLICY
'   Raises a descriptive runtime error if one or more footer labels cannot be
'   created, formatted, or hooked
'
' DEPENDENCIES
'   UF_Ensure_Label
'   UF_CalendarGrid_GetBottom
'   M_Caption_GetDate
'   cDatePickerLabelHook
'   StdFont
'
' NOTES
'   Icon captions are assigned only after a clean StdFont object has been applied
'   to prevent inherited Italic / Strikethrough states from corrupting glyphs
'
'   Developer-owned footer constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_Footer_BuildLabels"

    Dim BannerLeft              As Single                       'Footer banner left position
    Dim BannerTop               As Single                       'Footer banner top position

    Dim Lbl_TodayHalo           As MSForms.Label                'Footer Today halo label
    Dim Lbl_TimeHalo            As MSForms.Label                'Footer Time halo label

    Dim Lbl_TodayIcon           As MSForms.Label                'Footer Today icon label
    Dim Lbl_TodayCaption        As MSForms.Label                'Footer Today caption label
    Dim Lbl_Today               As MSForms.Label                'Footer Today value label

    Dim Lbl_TimeIcon            As MSForms.Label                'Footer Time icon label
    Dim Lbl_TimeCaption         As MSForms.Label                'Footer Time caption label
    Dim Lbl_Time                As MSForms.Label                'Footer Time value label

    Dim LabelHook               As cDatePickerLabelHook         'Runtime footer click / hover hook
    Dim IconFont                As Object                       'Clean icon font object
    Dim CaptionFont             As Object                       'Clean caption-label font object
    Dim ValueFont               As Object                       'Clean value-label font object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Create or reset footer-label click / hover hooks
        Set mFooterLabelHooks = New Collection
    'Clear any stale footer hover tracker before rebuilding labels
        mHoveredFooterActionName = vbNullString

'------------------------------------------------------------------------------
' CALCULATE LAYOUT
'------------------------------------------------------------------------------
    'Calculate the footer banner left position
        BannerLeft = DP_DAY_GRID_START_LEFT - DP_FOOTER_SIDE_PADDING
    'Calculate the footer banner top position
        BannerTop = UF_CalendarGrid_GetBottom() + DP_FOOTER_TOP_GAP

'------------------------------------------------------------------------------
' CREATE CLEAN FONTS
'------------------------------------------------------------------------------
    'Create a clean icon font object
        Set IconFont = CreateObject("StdFont")
    'Configure the icon font explicitly
        With IconFont
            .Name = DP_FOOTER_ICON_FONT_NAME
            .Size = DP_FOOTER_ICON_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With
    'Create a clean caption-label font object
        Set CaptionFont = CreateObject("StdFont")
    'Configure the caption-label font explicitly
        With CaptionFont
            .Name = DP_FORM_FONT_NAME
            .Size = 10
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With
    'Create a clean value-label font object
        Set ValueFont = CreateObject("StdFont")
    'Configure the value-label font explicitly
        With ValueFont
            .Name = DP_FORM_FONT_NAME
            .Size = 8
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

'------------------------------------------------------------------------------
' CREATE / FORMAT TODAY HALO LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the Today halo label
        Set Lbl_TodayHalo = UF_Ensure_Label("Lbl_TodayHalo")
    'Apply layout and visual properties
        With Lbl_TodayHalo
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_TODAY_HALO_REL_LEFT
            .Top = BannerTop + DP_FOOTER_HALO_TOP
            .Width = DP_FOOTER_TODAY_HALO_WIDTH
            .Height = DP_FOOTER_HALO_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_FOOTER_HALO_BACK_COLOR
            .ForeColor = DP_FOOTER_HOVER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = False
            .Visible = False
        End With
    'Move the Today halo above the footer banner
        Lbl_TodayHalo.ZOrder 0

'------------------------------------------------------------------------------
' CREATE / FORMAT TIME HALO LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the Time halo label
        Set Lbl_TimeHalo = UF_Ensure_Label("Lbl_TimeHalo")
    'Apply layout and visual properties
        With Lbl_TimeHalo
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_TIME_HALO_REL_LEFT
            .Top = BannerTop + DP_FOOTER_HALO_TOP
            .Width = DP_FOOTER_TIME_HALO_WIDTH
            .Height = DP_FOOTER_HALO_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_FOOTER_HALO_BACK_COLOR
            .ForeColor = DP_FOOTER_HOVER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = False
            .Visible = False
        End With
    'Move the Time halo above the footer banner
        Lbl_TimeHalo.ZOrder 0

'------------------------------------------------------------------------------
' CREATE / FORMAT TODAY ICON LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the Today icon label
        Set Lbl_TodayIcon = UF_Ensure_Label("Lbl_TodayIcon")
    'Apply layout and visual properties
        With Lbl_TodayIcon
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_TODAY_ICON_REL_LEFT
            .Top = BannerTop + DP_FOOTER_ICON_TOP
            .Width = DP_FOOTER_ICON_WIDTH
            .Height = DP_FOOTER_ICON_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_FOOTER_ICON_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Select Today"
        End With
    'Assign the clean icon font to the Today icon label
        Set Lbl_TodayIcon.Font = IconFont
    'Apply the Today icon glyph after the clean font has been assigned
        Lbl_TodayIcon.Caption = ChrW$(DP_FOOTER_ICON_CALENDAR_CODEPOINT)
    'Move the Today icon label to the front
        Lbl_TodayIcon.ZOrder 0
    'Create the Today icon click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the Today icon to the write-today action
        LabelHook.Initialize Me, Lbl_TodayIcon, "WRITE_TODAY", "FOOTER"
    'Store the hook so that the click / hover events remain alive
        mFooterLabelHooks.Add LabelHook, Lbl_TodayIcon.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT TODAY CAPTION LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the Today caption label
        Set Lbl_TodayCaption = UF_Ensure_Label("Lbl_TodayCaption")
    'Apply layout and visual properties
        With Lbl_TodayCaption
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_TODAY_CAPTION_REL_LEFT
            .Top = BannerTop + DP_FOOTER_TODAY_CAPTION_TOP
            .Width = DP_FOOTER_TODAY_CAPTION_WIDTH
            .Height = DP_FOOTER_TODAY_CAPTION_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_FOOTER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Select Today"
        End With
    'Assign the clean caption font to the Today caption label
        Set Lbl_TodayCaption.Font = CaptionFont
    'Apply the caption text after the clean font has been assigned
        Lbl_TodayCaption.Caption = "Today"
    'Move the Today caption label to the front
        Lbl_TodayCaption.ZOrder 0
    'Create the Today caption click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the Today caption to the write-today action
        LabelHook.Initialize Me, Lbl_TodayCaption, "WRITE_TODAY", "FOOTER"
    'Store the hook so that the click / hover events remain alive
        mFooterLabelHooks.Add LabelHook, Lbl_TodayCaption.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT TODAY VALUE LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the Today value label
        Set Lbl_Today = UF_Ensure_Label("Lbl_Today")
    'Apply layout and visual properties
        With Lbl_Today
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_TODAY_VALUE_REL_LEFT
            .Top = BannerTop + DP_FOOTER_TODAY_VALUE_TOP
            .Width = DP_FOOTER_TODAY_VALUE_WIDTH
            .Height = DP_FOOTER_TODAY_VALUE_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_FOOTER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Select Today"
        End With
    'Assign the clean value font to the Today value label
        Set Lbl_Today.Font = ValueFont
    'Apply the Today value caption after the clean font has been assigned
        Lbl_Today.Caption = M_Caption_GetDate(VBA.Date, gDP_UseLocalNames)
    'Move the Today value label to the front
        Lbl_Today.ZOrder 0
    'Create the Today value click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the Today value to the write-today action
        LabelHook.Initialize Me, Lbl_Today, "WRITE_TODAY", "FOOTER"
    'Store the hook so that the click / hover events remain alive
        mFooterLabelHooks.Add LabelHook, Lbl_Today.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT TIME ICON LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the Time icon label
        Set Lbl_TimeIcon = UF_Ensure_Label("Lbl_TimeIcon")
    'Apply layout and visual properties
        With Lbl_TimeIcon
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_TIME_ICON_REL_LEFT
            .Top = BannerTop + DP_FOOTER_ICON_TOP
            .Width = DP_FOOTER_ICON_WIDTH
            .Height = DP_FOOTER_ICON_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_FOOTER_ICON_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Select Now"
        End With
    'Assign the clean icon font to the Time icon label
        Set Lbl_TimeIcon.Font = IconFont
    'Apply the Time icon glyph after the clean font has been assigned
        Lbl_TimeIcon.Caption = ChrW$(DP_FOOTER_ICON_CLOCK_CODEPOINT)
    'Move the Time icon label to the front
        Lbl_TimeIcon.ZOrder 0
    'Create the Time icon click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the Time icon to the write-now action
        LabelHook.Initialize Me, Lbl_TimeIcon, "WRITE_NOW", "FOOTER"
    'Store the hook so that the click / hover events remain alive
        mFooterLabelHooks.Add LabelHook, Lbl_TimeIcon.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT TIME CAPTION LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the Time caption label
        Set Lbl_TimeCaption = UF_Ensure_Label("Lbl_TimeCaption")
    'Apply layout and visual properties
        With Lbl_TimeCaption
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_TIME_CAPTION_REL_LEFT
            .Top = BannerTop + DP_FOOTER_TIME_CAPTION_TOP
            .Width = DP_FOOTER_TIME_CAPTION_WIDTH
            .Height = DP_FOOTER_TIME_CAPTION_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_FOOTER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Select Now"
        End With
    'Assign the clean caption font to the Time caption label
        Set Lbl_TimeCaption.Font = CaptionFont
    'Apply the caption text after the clean font has been assigned
        Lbl_TimeCaption.Caption = "Time"
    'Move the Time caption label to the front
        Lbl_TimeCaption.ZOrder 0
    'Create the Time caption click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the Time caption to the write-now action
        LabelHook.Initialize Me, Lbl_TimeCaption, "WRITE_NOW", "FOOTER"
    'Store the hook so that the click / hover events remain alive
        mFooterLabelHooks.Add LabelHook, Lbl_TimeCaption.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT TIME VALUE LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the Time value label
        Set Lbl_Time = UF_Ensure_Label("Lbl_Time")
    'Cache the Time value label for live-clock updates
        Set mLbl_Time = Lbl_Time
    'Apply layout and visual properties
        With Lbl_Time
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_TIME_VALUE_REL_LEFT
            .Top = BannerTop + DP_FOOTER_TIME_VALUE_TOP
            .Width = DP_FOOTER_TIME_VALUE_WIDTH
            .Height = DP_FOOTER_TIME_VALUE_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_FOOTER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = "Select Now"
        End With
    'Assign the clean value font to the Time value label
        Set Lbl_Time.Font = ValueFont
    'Apply the Time value caption after the clean font has been assigned
        Lbl_Time.Caption = VBA.Format$(VBA.Time, "hh:nn:ss")
    'Move the Time value label to the front
        Lbl_Time.ZOrder 0
    'Create the Time value click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the Time value to the write-now action
        LabelHook.Initialize Me, Lbl_Time, "WRITE_NOW", "FOOTER"
    'Store the hook so that the click / hover events remain alive
        mFooterLabelHooks.Add LabelHook, Lbl_Time.Name

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
        Err.Raise Err.Number, PROC_NAME, "Footer label creation failed: " & Err.Description

End Sub

Private Sub UF_Footer_BuildSettingsArea()

'
'------------------------------------------------------------------------------
'                         BUILD FOOTER SETTINGS AREA
'------------------------------------------------------------------------------
' PURPOSE
'   Creates and formats the footer Settings action group
'
' WHY THIS EXISTS
'   The footer exposes Today, Time, and Settings as action groups
'
'   Settings is visually smaller than Today and Time, but it follows the same
'   hover model:
'     - hidden halo label behind the action
'     - visible icon label above the halo
'     - routed click / hover hook on the visible icon only
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses Lbl_SettingsSeparator, Lbl_SettingsHalo, and
'   Lbl_SettingsIcon. Positions the Settings group directly on the UserForm,
'   aligned to the footer banner geometry, and registers the Settings icon hook
'
' ERROR POLICY
'   Raises a descriptive runtime error if the Settings footer area cannot be
'   created, formatted, or hooked
'
' DEPENDENCIES
'   UF_Ensure_Label
'   UF_CalendarGrid_GetBottom
'   cDatePickerLabelHook
'   mFooterLabelHooks
'
' NOTES
'   There is no footer frame in this form
'
'   Footer controls are positioned directly on the UserForm
'
'   The halo and separator are disabled because they are purely visual. The icon
'   is the event surface for click and hover routing
'
'   In compact layout, footer settings access may be outside the visible form
'   canvas. Compact settings access is handled separately by the header settings
'   icon
'
'   Developer-owned footer constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Footer_BuildSettingsArea"
    
    Dim LabelHook           As cDatePickerLabelHook     'Runtime footer click / hover hook
    Dim BannerLeft          As Single                   'Footer banner left position
    Dim BannerTop           As Single                   'Footer banner top position
    Dim Lbl_Separator       As MSForms.Label            'Footer separator label
    Dim Lbl_SettingsHalo    As MSForms.Label            'Footer Settings halo label
    Dim Lbl_SettingsIcon    As MSForms.Label            'Footer Settings icon label
    Dim IconFont            As Object                   'Clean settings icon font object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Create footer hook storage only when this routine is called independently
        If mFooterLabelHooks Is Nothing Then
            Set mFooterLabelHooks = New Collection
        End If

'------------------------------------------------------------------------------
' CALCULATE FOOTER POSITION
'------------------------------------------------------------------------------
    'Calculate the footer banner position
        BannerLeft = DP_DAY_GRID_START_LEFT - DP_FOOTER_SIDE_PADDING
        BannerTop = UF_CalendarGrid_GetBottom() + DP_FOOTER_TOP_GAP

'------------------------------------------------------------------------------
' CREATE / FORMAT SEPARATOR
'------------------------------------------------------------------------------
    'Create or retrieve the settings separator label
        Set Lbl_Separator = UF_Ensure_Label("Lbl_SettingsSeparator")
    'Apply separator layout and visual properties
        With Lbl_Separator
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_SEPARATOR_REL_LEFT
            .Top = BannerTop + DP_FOOTER_SEPARATOR_TOP
            .Width = DP_FOOTER_SEPARATOR_WIDTH
            .Height = DP_FOOTER_SEPARATOR_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_FOOTER_SEPARATOR_COLOR
            .ForeColor = DP_FOOTER_SEPARATOR_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = False
            .Visible = True
        End With

'------------------------------------------------------------------------------
' CREATE / FORMAT SETTINGS HALO
'------------------------------------------------------------------------------
    'Create or retrieve the Settings halo label
        Set Lbl_SettingsHalo = UF_Ensure_Label("Lbl_SettingsHalo")
    'Apply settings halo layout and visual properties
        With Lbl_SettingsHalo
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_SETTINGS_HALO_REL_LEFT
            .Top = BannerTop + DP_FOOTER_SETTINGS_HALO_TOP
            .Width = DP_FOOTER_SETTINGS_HALO_WIDTH
            .Height = DP_FOOTER_SETTINGS_HALO_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_FOOTER_HALO_BACK_COLOR
            .ForeColor = DP_FOOTER_HOVER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = False
            .Visible = False
        End With

'------------------------------------------------------------------------------
' CREATE / FORMAT SETTINGS ICON
'------------------------------------------------------------------------------
    'Create or retrieve the Settings icon label
        Set Lbl_SettingsIcon = UF_Ensure_Label("Lbl_SettingsIcon")
    'Apply settings icon layout and visual properties
        With Lbl_SettingsIcon
            .Caption = vbNullString
            .Left = BannerLeft + DP_FOOTER_SETTINGS_ICON_REL_LEFT
            .Top = BannerTop + DP_FOOTER_SETTINGS_ICON_TOP
            .Width = DP_FOOTER_SETTINGS_ICON_WIDTH
            .Height = DP_FOOTER_SETTINGS_ICON_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_FOOTER_ICON_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .ControlTipText = DP_FOOTER_SETTINGS_ICON_CAPTION
        End With

'------------------------------------------------------------------------------
' APPLY SETTINGS ICON FONT
'------------------------------------------------------------------------------
    'Create a clean settings icon font object
        Set IconFont = CreateObject("StdFont")
    'Configure the settings icon font explicitly
        With IconFont
            .Name = DP_FOOTER_SETTINGS_ICON_FONT_NAME
            .Size = DP_FOOTER_SETTINGS_ICON_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With
    'Assign the clean font to the Settings icon label
        Set Lbl_SettingsIcon.Font = IconFont
    'Apply the Settings glyph after assigning the clean font
        Lbl_SettingsIcon.Caption = ChrW$(DP_FOOTER_SETTINGS_ICON_CODEPOINT)

'------------------------------------------------------------------------------
' REGISTER SETTINGS ICON HOOK
'------------------------------------------------------------------------------
    'Create the Settings icon click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the Settings icon to the settings action
        LabelHook.Initialize Me, Lbl_SettingsIcon, "SHOW_SETTINGS", "FOOTER"
    'Remove any previous Settings icon hook key when rebuilding this area independently
        On Error Resume Next
        mFooterLabelHooks.Remove Lbl_SettingsIcon.Name
        Err.Clear
        On Error GoTo ErrorHandler
    'Store the hook so that the click / hover events remain alive
        mFooterLabelHooks.Add LabelHook, Lbl_SettingsIcon.Name

'------------------------------------------------------------------------------
' RESTORE LAYERING
'------------------------------------------------------------------------------
    'Move the Settings halo above the footer banner
        Lbl_SettingsHalo.ZOrder 0
    'Move the separator above the footer banner
        Lbl_Separator.ZOrder 0
    'Move the Settings icon above the halo
        Lbl_SettingsIcon.ZOrder 0

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
        Err.Raise Err.Number, PROC_NAME, "Footer settings area creation failed: " & Err.Description

End Sub

Public Sub UF_Footer_HoverApply(ByVal LabelName As String)

'
'------------------------------------------------------------------------------
'                           APPLY FOOTER HOVER
'------------------------------------------------------------------------------
' PURPOSE
'   Applies hover formatting to one footer action group
'
' WHY THIS EXISTS
'   Today, Time, and Settings are footer action groups
'
'   Each group needs a consistent hover treatment:
'     - reveal the group halo
'     - recolor the group foreground labels
'     - keep foreground labels above the halo
'
'   Settings is currently an icon-only group, so the routine must not assume
'   that every footer action has caption and value labels
'
' INPUTS
'   LabelName
'     Name of the footer label currently under the mouse pointer
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the footer action group, exits when the same group is already
'   hovered, resets the previous footer hover state, retrieves the current group
'   labels, validates required labels, reveals the halo, applies hover foreground
'   formatting, restores layering, and stores the new hover state
'
' ERROR POLICY
'   Raises a descriptive runtime error if the hovered label cannot be resolved or
'   if the required footer controls are unavailable
'
' DEPENDENCIES
'   UF_Footer_ActionFromLabelName
'   UF_Footer_HoverReset
'   Lbl_TodayHalo
'   Lbl_TodayIcon
'   Lbl_TodayCaption
'   Lbl_Today
'   Lbl_TimeHalo
'   Lbl_TimeIcon
'   Lbl_TimeCaption
'   Lbl_Time
'   Lbl_SettingsHalo
'   Lbl_SettingsIcon
'
' NOTES
'   Settings is treated as a real footer group, not as a special exception
'
'   Caption and value labels are required for Today and Time only
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_Footer_HoverApply"

    Dim ActionName              As String               'Resolved footer action name
    Dim RequiresTextLabels      As Boolean              'True when caption and value are required
    Dim Lbl_Halo                As MSForms.Label        'Footer halo label
    Dim Lbl_Icon                As MSForms.Label        'Footer icon label
    Dim Lbl_Caption             As MSForms.Label        'Footer caption label
    Dim Lbl_Value               As MSForms.Label        'Footer value label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty label name
        If VBA.Len(VBA.Trim$(LabelName)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "LabelName cannot be empty"
        End If

'------------------------------------------------------------------------------
' RESOLVE FOOTER ACTION
'------------------------------------------------------------------------------
    'Resolve the footer action from the hovered label name
        ActionName = UF_Footer_ActionFromLabelName(LabelName)
    'Reject unsupported footer labels
        If VBA.Len(ActionName) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "LabelName does not resolve to a supported footer action"
        End If

'------------------------------------------------------------------------------
' EXIT IF ALREADY HOVERED
'------------------------------------------------------------------------------
    'Exit if the requested footer action is already highlighted
        If VBA.StrComp(mHoveredFooterActionName, ActionName, vbBinaryCompare) = 0 Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RESET CURRENT HOVER STATE
'------------------------------------------------------------------------------
    'Remove hover formatting from the previously highlighted footer action
        If VBA.Len(mHoveredFooterActionName) <> 0 Then
            UF_Footer_HoverReset
        End If

'------------------------------------------------------------------------------
' RETRIEVE FOOTER GROUP LABELS
'------------------------------------------------------------------------------
    'Retrieve labels for the requested footer action
        Select Case ActionName
        
            Case "WRITE_TODAY"
                'Today requires icon, caption, value, and halo
                    RequiresTextLabels = True
                'Suppress lookup errors so missing labels can be reported consistently
                    On Error Resume Next
                'Retrieve the Today halo
                    Set Lbl_Halo = Me.Controls("Lbl_TodayHalo")
                'Retrieve the Today icon
                    Set Lbl_Icon = Me.Controls("Lbl_TodayIcon")
                'Retrieve the Today caption
                    Set Lbl_Caption = Me.Controls("Lbl_TodayCaption")
                'Retrieve the Today value
                    Set Lbl_Value = Me.Controls("Lbl_Today")
                'Clear suppressed lookup errors
                    Err.Clear
                'Restore controlled error handling
                    On Error GoTo ErrorHandler

            Case "WRITE_NOW"
                'Time requires icon, caption, value, and halo
                    RequiresTextLabels = True
                'Suppress lookup errors so missing labels can be reported consistently
                    On Error Resume Next
                'Retrieve the Time halo
                    Set Lbl_Halo = Me.Controls("Lbl_TimeHalo")
                'Retrieve the Time icon
                    Set Lbl_Icon = Me.Controls("Lbl_TimeIcon")
                'Retrieve the Time caption
                    Set Lbl_Caption = Me.Controls("Lbl_TimeCaption")
                'Retrieve the Time value
                    Set Lbl_Value = Me.Controls("Lbl_Time")
                'Clear suppressed lookup errors
                    Err.Clear
                'Restore controlled error handling
                    On Error GoTo ErrorHandler

            Case "SHOW_SETTINGS"
                'Settings requires only icon and halo
                    RequiresTextLabels = False
                'Suppress lookup errors so missing labels can be reported consistently
                    On Error Resume Next
                'Retrieve the Settings halo
                    Set Lbl_Halo = Me.Controls("Lbl_SettingsHalo")
                'Retrieve the Settings icon
                    Set Lbl_Icon = Me.Controls("Lbl_SettingsIcon")
                'Clear suppressed lookup errors
                    Err.Clear
                'Restore controlled error handling
                    On Error GoTo ErrorHandler

            Case Else
                'Reject unsupported footer actions
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Unsupported footer action: " & ActionName

        End Select

'------------------------------------------------------------------------------
' VALIDATE REQUIRED LABELS
'------------------------------------------------------------------------------
    'Reject a missing footer halo
        If Lbl_Halo Is Nothing Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Footer halo label is missing for action " & ActionName
        End If
    'Reject a missing footer icon
        If Lbl_Icon Is Nothing Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Footer icon label is missing for action " & ActionName
        End If
    'Reject missing text labels when required
        If RequiresTextLabels Then
            'Reject a missing footer caption
                If Lbl_Caption Is Nothing Then
                    Err.Raise vbObjectError + 518, PROC_NAME, _
                        "Footer caption label is missing for action " & ActionName
                End If
            'Reject a missing footer value
                If Lbl_Value Is Nothing Then
                    Err.Raise vbObjectError + 519, PROC_NAME, _
                        "Footer value label is missing for action " & ActionName
                End If
        End If

'------------------------------------------------------------------------------
' APPLY HALO STATE
'------------------------------------------------------------------------------
    'Reveal the footer halo
        With Lbl_Halo
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_FOOTER_HALO_BACK_COLOR
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = vbWhite
            .SpecialEffect = fmSpecialEffectRaised
            .Visible = True
        End With
    'Move the halo above the footer banner
        Lbl_Halo.ZOrder 0

'------------------------------------------------------------------------------
' APPLY ICON HOVER STATE
'------------------------------------------------------------------------------
    'Apply hover formatting to the footer icon
        With Lbl_Icon
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_FOOTER_HOVER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
        End With

'------------------------------------------------------------------------------
' APPLY OPTIONAL CAPTION HOVER STATE
'------------------------------------------------------------------------------
    'Apply hover formatting to the footer caption when present
        If Not Lbl_Caption Is Nothing Then
            With Lbl_Caption
                .BackStyle = fmBackStyleTransparent
                .ForeColor = DP_FOOTER_HOVER_FORE_COLOR
                .BorderStyle = fmBorderStyleNone
                .SpecialEffect = fmSpecialEffectFlat
            End With
        End If

'------------------------------------------------------------------------------
' APPLY OPTIONAL VALUE HOVER STATE
'------------------------------------------------------------------------------
    'Apply hover formatting to the footer value when present
        If Not Lbl_Value Is Nothing Then
            With Lbl_Value
                .BackStyle = fmBackStyleTransparent
                .ForeColor = DP_FOOTER_HOVER_FORE_COLOR
                .BorderStyle = fmBorderStyleNone
                .SpecialEffect = fmSpecialEffectFlat
            End With
        End If

'------------------------------------------------------------------------------
' RESTORE FRONT LAYER
'------------------------------------------------------------------------------
    'Move the footer icon above the halo
        Lbl_Icon.ZOrder 0
    'Move the footer caption above the halo when present
        If Not Lbl_Caption Is Nothing Then
            Lbl_Caption.ZOrder 0
        End If
    'Move the footer value above the halo when present
        If Not Lbl_Value Is Nothing Then
            Lbl_Value.ZOrder 0
        End If

'------------------------------------------------------------------------------
' STORE HOVER STATE
'------------------------------------------------------------------------------
    'Store the currently hovered footer action
        mHoveredFooterActionName = ActionName

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
        Err.Raise Err.Number, PROC_NAME, "Footer hover application failed: " & Err.Description

End Sub

Private Sub UF_Footer_HoverReset()

'
'------------------------------------------------------------------------------
'                           RESET FOOTER HOVER
'------------------------------------------------------------------------------
' PURPOSE
'   Removes hover formatting from the currently highlighted footer action group
'
' WHY THIS EXISTS
'   Today, Time, and Settings are all footer action groups
'
'   Each group must be reset consistently:
'     - hide the group halo
'     - restore the normal icon color
'     - restore optional caption and value colors
'     - clear the stored footer hover state
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the currently hovered footer action, retrieves the related group
'   labels, hides the halo, restores normal visual formatting, and clears the
'   footer hover state
'
' ERROR POLICY
'   Best-effort UI cleanup. Suppresses reset errors because hover reset should
'   never interrupt UserForm interaction
'
' DEPENDENCIES
'   UserForm.Controls collection
'   Lbl_TodayHalo
'   Lbl_TodayIcon
'   Lbl_TodayCaption
'   Lbl_Today
'   Lbl_TimeHalo
'   Lbl_TimeIcon
'   Lbl_TimeCaption
'   Lbl_Time
'   Lbl_SettingsHalo
'   Lbl_SettingsIcon
'
' NOTES
'   Settings is treated as a footer group with a halo and an icon
'
'   Settings has no caption or value label, so those references remain optional
'
'   This routine intentionally does not raise outward because it is normally
'   called from high-frequency mouse movement paths
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentAction       As String               'Current hovered footer action
    
    Dim Lbl_Halo            As MSForms.Label        'Footer halo label
    Dim Lbl_Icon            As MSForms.Label        'Footer icon label
    Dim Lbl_Caption         As MSForms.Label        'Footer caption label
    Dim Lbl_Value           As MSForms.Label        'Footer value label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress reset errors
        On Error Resume Next
    'Capture the current hovered action
        CurrentAction = VBA.Trim$(mHoveredFooterActionName)

'------------------------------------------------------------------------------
' EXIT IF NOTHING IS HOVERED
'------------------------------------------------------------------------------
    'Exit if no footer action is currently highlighted
        If VBA.Len(CurrentAction) = 0 Then GoTo Clean_Exit

'------------------------------------------------------------------------------
' RETRIEVE FOOTER GROUP LABELS
'------------------------------------------------------------------------------
    'Retrieve labels for the currently hovered footer action
        Select Case CurrentAction

            Case "WRITE_TODAY"
                'Retrieve the Today halo
                    Set Lbl_Halo = Me.Controls("Lbl_TodayHalo")
                'Retrieve the Today icon
                    Set Lbl_Icon = Me.Controls("Lbl_TodayIcon")
                'Retrieve the Today caption
                    Set Lbl_Caption = Me.Controls("Lbl_TodayCaption")
                'Retrieve the Today value
                    Set Lbl_Value = Me.Controls("Lbl_Today")

            Case "WRITE_NOW"
                'Retrieve the Time halo
                    Set Lbl_Halo = Me.Controls("Lbl_TimeHalo")
                'Retrieve the Time icon
                    Set Lbl_Icon = Me.Controls("Lbl_TimeIcon")
                'Retrieve the Time caption
                    Set Lbl_Caption = Me.Controls("Lbl_TimeCaption")
                'Retrieve the Time value
                    Set Lbl_Value = Me.Controls("Lbl_Time")

            Case "SHOW_SETTINGS"
                'Retrieve the Settings halo
                    Set Lbl_Halo = Me.Controls("Lbl_SettingsHalo")
                'Retrieve the Settings icon
                    Set Lbl_Icon = Me.Controls("Lbl_SettingsIcon")

            Case Else
                'Ignore unsupported stale footer actions
                    GoTo Clean_Exit

        End Select

'------------------------------------------------------------------------------
' HIDE GROUP HALO
'------------------------------------------------------------------------------
    'Reset and hide the footer halo when available
        If Not Lbl_Halo Is Nothing Then
            With Lbl_Halo
                .BackStyle = fmBackStyleOpaque
                .BackColor = DP_FOOTER_HALO_BACK_COLOR
                .BorderStyle = fmBorderStyleNone
                .SpecialEffect = fmSpecialEffectFlat
                .Visible = False
            End With
        End If

'------------------------------------------------------------------------------
' RESET ICON VISUAL STATE
'------------------------------------------------------------------------------
    'Reset the footer icon label when available
        If Not Lbl_Icon Is Nothing Then
            With Lbl_Icon
                .BackStyle = fmBackStyleTransparent
                .ForeColor = DP_FOOTER_ICON_FORE_COLOR
                .BorderStyle = fmBorderStyleNone
                .SpecialEffect = fmSpecialEffectFlat
            End With
        End If

'------------------------------------------------------------------------------
' RESET OPTIONAL CAPTION VISUAL STATE
'------------------------------------------------------------------------------
    'Reset the footer caption label when available
        If Not Lbl_Caption Is Nothing Then
            With Lbl_Caption
                .BackStyle = fmBackStyleTransparent
                .ForeColor = DP_FOOTER_FORE_COLOR
                .BorderStyle = fmBorderStyleNone
                .SpecialEffect = fmSpecialEffectFlat
            End With
        End If

'------------------------------------------------------------------------------
' RESET OPTIONAL VALUE VISUAL STATE
'------------------------------------------------------------------------------
    'Reset the footer value label when available
        If Not Lbl_Value Is Nothing Then
            With Lbl_Value
                .BackStyle = fmBackStyleTransparent
                .ForeColor = DP_FOOTER_FORE_COLOR
                .BorderStyle = fmBorderStyleNone
                .SpecialEffect = fmSpecialEffectFlat
            End With
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:
    'Clear the current footer hover action
        mHoveredFooterActionName = vbNullString
    'Clear suppressed reset errors
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Function UF_Footer_ActionFromLabelName(ByVal LabelName As String) As String

'
'------------------------------------------------------------------------------
'                       GET FOOTER ACTION FROM LABEL NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the footer action associated with a footer label name
'
' WHY THIS EXISTS
'   Today, Time, and Settings are footer action groups composed of one or more
'   labels. Hover and click routing need a single action name regardless of
'   which label within the group raised the event
'
' INPUTS
'   LabelName
'     Name of the footer label to resolve
'
' RETURNS
'   Footer action name when the label belongs to a supported footer group:
'     - WRITE_TODAY
'     - WRITE_NOW
'     - SHOW_SETTINGS
'
'   vbNullString when the label is not part of a supported footer action group
'
' BEHAVIOR
'   Normalizes the supplied label name and maps known Today, Time, and Settings
'   footer labels to their routed action names
'
' ERROR POLICY
'   Does not raise errors. Unsupported labels return vbNullString so callers can
'   decide whether the condition is expected or exceptional
'
' DEPENDENCIES
'   None
'
' NOTES
'   The comparison is normalized with UCase$ so the result is independent from
'   the module-level Option Compare setting
'
'   Lbl_SettingsSeparator is intentionally excluded because it is decorative and
'   not an action surface
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the supplied label name
        LabelName = VBA.UCase$(VBA.Trim$(LabelName))

'------------------------------------------------------------------------------
' RETURN ACTION
'------------------------------------------------------------------------------
    'Resolve footer labels
        Select Case LabelName

            Case "LBL_TODAYICON", _
                 "LBL_TODAYCAPTION", _
                 "LBL_TODAY", _
                 "LBL_TODAYHALO"
                
                UF_Footer_ActionFromLabelName = "WRITE_TODAY"
                Exit Function

            Case "LBL_TIMEICON", _
                 "LBL_TIMECAPTION", _
                 "LBL_TIME", _
                 "LBL_TIMEHALO"

                UF_Footer_ActionFromLabelName = "WRITE_NOW"
                Exit Function

            Case "LBL_SETTINGSICON", _
                 "LBL_SETTINGSHALO"

                UF_Footer_ActionFromLabelName = "SHOW_SETTINGS"
                Exit Function

        End Select

'------------------------------------------------------------------------------
' RETURN FALLBACK
'------------------------------------------------------------------------------
    'Return an empty action for unsupported labels
        UF_Footer_ActionFromLabelName = vbNullString

End Function


'
'------------------------------------------------------------------------------
'
'                                 PICKER PANEL
'
'------------------------------------------------------------------------------
'

Private Sub UF_PickerPanel_Build()

'
'------------------------------------------------------------------------------
'                           CREATE PICKER PANEL
'------------------------------------------------------------------------------
' PURPOSE
'   Creates the reusable flat picker panel used for month and year selection
'
' WHY THIS EXISTS
'   The DatePicker needs a flat selector that is visually consistent with the
'   label-based UI and reusable for both month and year selection
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses Fra_PickerPanel and 12 background/text label pairs,
'   formats the panel and picker items, caches label references, and registers
'   click / hover hooks for each item surface
'
' ERROR POLICY
'   Raises a descriptive runtime error if the picker panel cannot be created,
'   formatted, cached, layered, or hooked
'
' DEPENDENCIES
'   UF_Ensure_FrameLabel
'   UF_CalendarGrid_GetWidth
'   cDatePickerLabelHook
'   mPickerPanelHooks
'   mPickerBackLabels
'   mPickerTextLabels
'
' NOTES
'   Year navigation is handled by the header year-arrow labels
'
'   Developer-owned picker-panel constants are intentionally not validated here.
'   They should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_PickerPanel_Build"
    
    Dim ExistingControl     As MSForms.control              'Existing control using the picker-panel name
    Dim Fra_PickerPanel     As MSForms.Frame                'Reusable picker panel
    Dim LabelHook           As cDatePickerLabelHook         'Runtime click / hover hook
    Dim Index               As Long                         'Picker item index
    Dim RowIndex            As Long                         'Zero-based row index
    Dim ColIndex            As Long                         'Zero-based column index
    Dim ControlName         As String                       'Runtime control name
    Dim ActionName          As String                       'Runtime routed picker action
    Dim ItemGridWidth       As Single                       'Total width of the picker item grid
    Dim ItemGridLeft        As Single                       'Centered item-grid left position
    Dim Lbl_ItemBg          As MSForms.Label                'Picker item background label
    Dim Lbl_Item            As MSForms.Label                'Picker item text label
    Dim PickerItemFont      As Object                       'Clean picker-item text font object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Reset picker-panel click / hover hooks
        Set mPickerPanelHooks = New Collection
    'Clear any stale picker-panel hover state
        mHoveredPickerItemIndex = 0

'------------------------------------------------------------------------------
' CREATE CLEAN ITEM FONT
'------------------------------------------------------------------------------
    'Create a clean picker-item font object
        Set PickerItemFont = CreateObject("StdFont")
    'Configure the picker-item font explicitly
        With PickerItemFont
            .Name = DP_FORM_FONT_NAME
            .Size = DP_FORM_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

'------------------------------------------------------------------------------
' CREATE / RETRIEVE FRAME
'------------------------------------------------------------------------------
    'Suppress lookup errors while checking for an existing picker panel
        On Error Resume Next
    'Try to retrieve an existing control with the picker-panel name
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler
    
    'Create the picker panel if it does not already exist
        If ExistingControl Is Nothing Then
            'Create the reusable picker panel frame
                Set Fra_PickerPanel = Me.Controls.Add("Forms.Frame.1", DP_PICKER_PANEL_NAME, True)
        Else
            'Reject a name collision with a non-frame control
                If VBA.TypeName(ExistingControl) <> "Frame" Then
                    Err.Raise vbObjectError + 513, PROC_NAME, _
                        "Control '" & DP_PICKER_PANEL_NAME & "' exists but is not an MSForms.Frame"
                End If
            'Use the existing picker panel frame
                Set Fra_PickerPanel = ExistingControl
        End If

'------------------------------------------------------------------------------
' FORMAT FRAME
'------------------------------------------------------------------------------
    'Apply layout and visual properties to the picker panel
        With Fra_PickerPanel
            .Caption = vbNullString
            .Width = DP_PICKER_PANEL_WIDTH
            .Height = DP_PICKER_PANEL_HEIGHT
            .Left = DP_DAY_GRID_START_LEFT + _
                ((UF_CalendarGrid_GetWidth() - DP_PICKER_PANEL_WIDTH) / 2)
            .Top = DP_DAY_GRID_START_TOP - 7
            .BackColor = DP_PICKER_PANEL_BACK_COLOR
            .BorderStyle = fmBorderStyleSingle
            .SpecialEffect = fmSpecialEffectFlat
            .Visible = False
        End With

'------------------------------------------------------------------------------
' CALCULATE ITEM GRID POSITION
'------------------------------------------------------------------------------
    'Calculate the total width of the picker item grid
        ItemGridWidth = _
            (DP_PICKER_ITEM_WIDTH * DP_PICKER_ITEMS_PER_ROW) + _
            ((DP_PICKER_ITEM_HORIZONTAL_STEP - DP_PICKER_ITEM_WIDTH) * _
            (DP_PICKER_ITEMS_PER_ROW - 1))
    'Calculate the centered left position of the item grid
        ItemGridLeft = (Fra_PickerPanel.Width - ItemGridWidth) / 2

'------------------------------------------------------------------------------
' CREATE BACKGROUND AND TEXT LABEL PAIRS
'------------------------------------------------------------------------------
    'Loop through the picker items
        For Index = 1 To DP_PICKER_ITEM_COUNT
            'Calculate the zero-based row index
                RowIndex = (Index - 1) \ DP_PICKER_ITEMS_PER_ROW
            'Calculate the zero-based column index
                ColIndex = (Index - 1) Mod DP_PICKER_ITEMS_PER_ROW
            'Build the routed action name
                ActionName = "PICKER_ITEM_" & VBA.CStr(Index)

'------------------------------------------------------------------------------
' CREATE / FORMAT BACKGROUND LABEL
'------------------------------------------------------------------------------
            'Build the background control name
                ControlName = "Lbl_MonthYearBg" & VBA.CStr(Index)
            'Create or retrieve the background label
                Set Lbl_ItemBg = UF_Ensure_FrameLabel(Fra_PickerPanel, ControlName)
            'Cache the picker-panel background label for fast reuse
                Set mPickerBackLabels(Index) = Lbl_ItemBg
            'Apply background label layout and visual properties
                With Lbl_ItemBg
                    .Caption = vbNullString
                    .Left = ItemGridLeft + DP_PICKER_ITEM_HORIZONTAL_STEP * ColIndex
                    .Top = DP_PICKER_ITEM_TOP + DP_PICKER_ITEM_VERTICAL_STEP * RowIndex
                    .Width = DP_PICKER_ITEM_WIDTH
                    .Height = DP_PICKER_ITEM_HEIGHT
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = DP_PICKER_ITEM_BACK_COLOR
                    .ForeColor = DP_PICKER_ITEM_FORE_COLOR
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = DP_PICKER_ITEM_BORDER_COLOR
                    .SpecialEffect = fmSpecialEffectFlat
                    .Enabled = True
                    .Visible = True
                End With

'------------------------------------------------------------------------------
' CREATE / FORMAT TEXT LABEL
'------------------------------------------------------------------------------
            'Build the text control name
                ControlName = "Lbl_MonthYear" & VBA.CStr(Index)
            'Create or retrieve the text label
                Set Lbl_Item = UF_Ensure_FrameLabel(Fra_PickerPanel, ControlName)
            'Cache the picker-panel text label for fast reuse
                Set mPickerTextLabels(Index) = Lbl_Item
            'Apply text label layout and visual properties
                With Lbl_Item
                    .Caption = vbNullString
                    .Left = Lbl_ItemBg.Left
                    .Top = Lbl_ItemBg.Top + DP_PICKER_ITEM_TEXT_TOP_OFFSET
                    .Width = Lbl_ItemBg.Width
                    .Height = DP_PICKER_ITEM_TEXT_HEIGHT
                    .BackStyle = fmBackStyleTransparent
                    .ForeColor = DP_PICKER_ITEM_FORE_COLOR
                    .BorderStyle = fmBorderStyleNone
                    .SpecialEffect = fmSpecialEffectFlat
                    .TextAlign = fmTextAlignCenter
                    .WordWrap = False
                    .Enabled = True
                    .Visible = True
                End With
            'Assign the clean picker-item font to the text label
                Set Lbl_Item.Font = PickerItemFont

'------------------------------------------------------------------------------
' RESTORE ITEM LAYERING
'------------------------------------------------------------------------------
            'Move the background behind the text label
                Lbl_ItemBg.ZOrder 1
            'Move the text label to the front
                Lbl_Item.ZOrder 0

'------------------------------------------------------------------------------
' REGISTER BACKGROUND LABEL HOOK
'------------------------------------------------------------------------------
            'Create the background-label click / hover hook
                Set LabelHook = New cDatePickerLabelHook
            'Connect the background label to its routed picker action
                LabelHook.Initialize Me, Lbl_ItemBg, ActionName, "PICKER"
            'Store the hook so that the click / hover events remain alive
                mPickerPanelHooks.Add LabelHook, Lbl_ItemBg.Name

'------------------------------------------------------------------------------
' REGISTER TEXT LABEL HOOK
'------------------------------------------------------------------------------
            'Create the text-label click / hover hook
                Set LabelHook = New cDatePickerLabelHook
            'Connect the text label to its routed picker action
                LabelHook.Initialize Me, Lbl_Item, ActionName, "PICKER"
            'Store the hook so that the click / hover events remain alive
                mPickerPanelHooks.Add LabelHook, Lbl_Item.Name
        Next Index

'------------------------------------------------------------------------------
' MOVE PANEL TO FRONT
'------------------------------------------------------------------------------
    'Move the picker panel in front of the day grid
        Fra_PickerPanel.ZOrder 0

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
        Err.Raise Err.Number, PROC_NAME, "Picker panel creation failed: " & Err.Description

End Sub

Private Sub UF_PickerPanel_EnsureCache()

'
'------------------------------------------------------------------------------
'                           ENSURE PICKER PANEL CACHE
'------------------------------------------------------------------------------
' PURPOSE
'   Ensures the picker-panel label cache is populated
'
' WHY THIS EXISTS
'   The DatePicker uses cached references to Lbl_MonthYear1 to Lbl_MonthYear12
'   and Lbl_MonthYearBg1 to Lbl_MonthYearBg12 to avoid repeated frame Controls
'   string lookups during month/year panel population and hover handling
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Rebuilds missing cached references from Fra_PickerPanel.Controls
'
' ERROR POLICY
'   Raises a descriptive runtime error if the picker panel or one or more
'   expected labels cannot be resolved
'
' DEPENDENCIES
'   DP_PICKER_PANEL_NAME
'   mPickerTextLabels
'   mPickerBackLabels
'   Fra_PickerPanel
'   Lbl_MonthYear1 to Lbl_MonthYear12
'   Lbl_MonthYearBg1 to Lbl_MonthYearBg12
'
' NOTES
'   This routine does not create controls. Controls are created by
'   UF_PickerPanel_Build
'
'   Developer-owned picker-panel constants are intentionally not validated here.
'   They should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_PickerPanel_EnsureCache"
    
    Dim Index               As Long                 'Picker-panel item index
    Dim ExistingControl     As MSForms.control      'Existing control using the picker-panel name
    Dim Fra_PickerPanel     As MSForms.Frame        'Reusable picker panel
    Dim TextControlName     As String               'Runtime picker text label name
    Dim BgControlName       As String               'Runtime picker background label name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RETRIEVE PICKER PANEL
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the picker panel
        On Error Resume Next
    'Retrieve the existing picker panel control
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject missing picker panel
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Unable to resolve expected picker panel " & DP_PICKER_PANEL_NAME
        End If
    'Reject a name collision with a non-frame control
        If VBA.TypeName(ExistingControl) <> "Frame" Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Control '" & DP_PICKER_PANEL_NAME & "' exists but is not an MSForms.Frame"
        End If

    'Use the resolved picker panel frame
        Set Fra_PickerPanel = ExistingControl

'------------------------------------------------------------------------------
' REBUILD MISSING CACHE REFERENCES
'------------------------------------------------------------------------------
    'Loop through all picker-panel items
        For Index = 1 To DP_PICKER_ITEM_COUNT
            'Build the runtime picker text label name
                TextControlName = "Lbl_MonthYear" & VBA.CStr(Index)
            'Build the runtime picker background label name
                BgControlName = "Lbl_MonthYearBg" & VBA.CStr(Index)

'------------------------------------------------------------------------------
' RESOLVE PICKER TEXT LABEL
'------------------------------------------------------------------------------
            'Resolve the cached picker-panel text label when missing
                If mPickerTextLabels(Index) Is Nothing Then
                    'Suppress missing-control lookup errors
                        On Error Resume Next
                    'Resolve the picker-panel text label from the frame controls collection
                        Set mPickerTextLabels(Index) = Fra_PickerPanel.Controls(TextControlName)
                    'Clear any suppressed lookup error
                        Err.Clear
                    'Restore controlled error handling
                        On Error GoTo ErrorHandler

                    'Reject unresolved picker-panel text label
                        If mPickerTextLabels(Index) Is Nothing Then
                            Err.Raise vbObjectError + 515, PROC_NAME, _
                                "Unable to resolve expected picker-panel text label " & TextControlName
                        End If
                End If

'------------------------------------------------------------------------------
' RESOLVE PICKER BACKGROUND LABEL
'------------------------------------------------------------------------------
            'Resolve the cached picker-panel background label when missing
                If mPickerBackLabels(Index) Is Nothing Then
                    'Suppress missing-control lookup errors
                        On Error Resume Next
                    'Resolve the picker-panel background label from the frame controls collection
                        Set mPickerBackLabels(Index) = Fra_PickerPanel.Controls(BgControlName)
                    'Clear any suppressed lookup error
                        Err.Clear
                    'Restore controlled error handling
                        On Error GoTo ErrorHandler

                    'Reject unresolved picker-panel background label
                        If mPickerBackLabels(Index) Is Nothing Then
                            Err.Raise vbObjectError + 516, PROC_NAME, _
                                "Unable to resolve expected picker-panel background label " & BgControlName
                        End If
                End If
        Next Index

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
        Err.Raise Err.Number, PROC_NAME, "Picker panel cache initialization failed: " & Err.Description

End Sub

Private Sub UF_PickerPanel_ClearCache()

'
'------------------------------------------------------------------------------
'                           CLEAR PICKER PANEL CACHE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears cached picker-panel label references
'
' WHY THIS EXISTS
'   Cached MSForms control references should be released when the UserForm is
'   terminating or when the picker-panel cache must be rebuilt from the frame
'   Controls collection
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears cached references to picker-panel text labels and background labels
'
' ERROR POLICY
'   Best-effort cleanup. Suppresses cleanup errors because this routine is
'   normally called during form teardown
'
' DEPENDENCIES
'   mPickerTextLabels
'   mPickerBackLabels
'
' NOTES
'   This routine does not delete controls
'
'   The explicit loop is intentional. It preserves array dimensions and is safe
'   for both fixed-size arrays and already-dimensioned dynamic arrays
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index       As Long     'Picker-panel item index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' CLEAR CACHED LABEL REFERENCES
'------------------------------------------------------------------------------
    'Loop through all cached picker-panel items
        For Index = 1 To DP_PICKER_ITEM_COUNT
            'Clear the cached picker-panel text label
                Set mPickerTextLabels(Index) = Nothing
            'Clear the cached picker-panel background label
                Set mPickerBackLabels(Index) = Nothing
        Next Index

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub UF_PickerPanel_ShowMonths()

'
'------------------------------------------------------------------------------
'                           SHOW MONTH PICKER PANEL
'------------------------------------------------------------------------------
' PURPOSE
'   Shows the reusable picker panel in month-selection mode
'
' WHY THIS EXISTS
'   Clicking the header month label or pressing M should display a flat month
'   selector without using a native ComboBox
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures picker-panel cache is available, resets active hover states,
'   hides the settings panel, populates the reusable picker items with month
'   captions, highlights the currently displayed month, and shows the picker
'   panel above the day grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if the picker panel or month labels cannot
'   be resolved, populated, or displayed
'
' DEPENDENCIES
'   UF_PickerPanel_EnsureCache
'   UF_PickerPanel_HoverReset
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'   UF_SettingsPanel_Hide
'   M_Caption_GetMonth
'   mPickerTextLabels
'   mPickerBackLabels
'
' NOTES
'   The current displayed month is highlighted
'
'   Developer-owned picker-panel constants are intentionally not validated here.
'   They should be checked by a dedicated debug or regression routine
'
'   Item layering is established when the picker panel is built. This routine
'   does not repeat item-level ZOrder calls inside the population loop
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_PickerPanel_ShowMonths"

    Dim ExistingControl     As MSForms.control          'Existing control using the picker-panel name
    Dim Fra_PickerPanel     As MSForms.Frame            'Reusable picker panel
    Dim Lbl_Item            As MSForms.Label            'Picker item text label
    Dim Lbl_ItemBg          As MSForms.Label            'Picker item background label
    Dim Index               As Long                     'Month index
    Dim MonthCaption        As String                   'Resolved month caption

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Ensure cached picker-panel references are available
        UF_PickerPanel_EnsureCache

'------------------------------------------------------------------------------
' RESET ACTIVE HOVER STATE
'------------------------------------------------------------------------------
    'Clear active picker hover before changing picker-panel mode
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Clear active day-cell hover before showing the picker panel
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Clear active header-label hover before showing the picker panel
        If VBA.Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Clear active footer hover before showing the picker panel
        If VBA.Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset
    'Clear active settings-panel hover before showing the picker panel
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then UF_SettingsPanel_HoverReset

'------------------------------------------------------------------------------
' HIDE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Hide the settings panel before showing the month picker
        UF_SettingsPanel_Hide

'------------------------------------------------------------------------------
' RETRIEVE PICKER PANEL
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the picker panel
        On Error Resume Next
    'Retrieve the existing picker panel control
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject missing picker panel
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Unable to resolve expected picker panel " & DP_PICKER_PANEL_NAME
        End If
    'Reject a name collision with a non-frame control
        If VBA.TypeName(ExistingControl) <> "Frame" Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Control '" & DP_PICKER_PANEL_NAME & "' exists but is not an MSForms.Frame"
        End If

    'Use the resolved picker panel frame
        Set Fra_PickerPanel = ExistingControl

'------------------------------------------------------------------------------
' SET PANEL MODE
'------------------------------------------------------------------------------
    'Store month-panel mode
        mPickerPanelMode = 1

'------------------------------------------------------------------------------
' POPULATE MONTHS
'------------------------------------------------------------------------------
    'Loop through the month picker items
        For Index = 1 To DP_PICKER_ITEM_COUNT
            'Retrieve the cached picker item text label
                Set Lbl_Item = mPickerTextLabels(Index)
            'Retrieve the cached picker item background label
                Set Lbl_ItemBg = mPickerBackLabels(Index)
            'Reject a missing picker item text label
                If Lbl_Item Is Nothing Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Cached picker text label is missing for index " & VBA.CStr(Index)
                End If
            'Reject a missing picker item background label
                If Lbl_ItemBg Is Nothing Then
                    Err.Raise vbObjectError + 516, PROC_NAME, _
                        "Cached picker background label is missing for index " & VBA.CStr(Index)
                End If
            'Resolve the three-letter month caption
                MonthCaption = VBA.UCase$(VBA.Left$(M_Caption_GetMonth(Index, gDP_UseLocalNames), 3))

'------------------------------------------------------------------------------
' APPLY MONTH TEXT STATE
'------------------------------------------------------------------------------
            'Apply the month caption and value
                With Lbl_Item
                    .Caption = MonthCaption
                    .Tag = VBA.CStr(Index)
                    .Left = Lbl_ItemBg.Left
                    .Top = Lbl_ItemBg.Top + DP_PICKER_ITEM_TEXT_TOP_OFFSET
                    .Width = Lbl_ItemBg.Width
                    .Height = DP_PICKER_ITEM_TEXT_HEIGHT
                    .BackStyle = fmBackStyleTransparent
                    .ForeColor = DP_PICKER_ITEM_FORE_COLOR
                    .BorderStyle = fmBorderStyleNone
                    .SpecialEffect = fmSpecialEffectFlat
                    .TextAlign = fmTextAlignCenter
                    .WordWrap = False
                    .Enabled = True
                    .Visible = True
                End With

'------------------------------------------------------------------------------
' APPLY MONTH BACKGROUND STATE
'------------------------------------------------------------------------------
            'Fully reset the picker item background
                With Lbl_ItemBg
                    .Caption = vbNullString
                    .Tag = VBA.CStr(Index)
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = DP_PICKER_ITEM_BACK_COLOR
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = DP_PICKER_ITEM_BORDER_COLOR
                    .SpecialEffect = fmSpecialEffectFlat
                    .Enabled = True
                    .Visible = True
                End With

'------------------------------------------------------------------------------
' APPLY SELECTED MONTH STATE
'------------------------------------------------------------------------------
            'Highlight the currently displayed month
                If Index = mDisplayMonth Then
                    With Lbl_ItemBg
                        .BackColor = DP_DAY_SELECTED_BACK_COLOR
                        .BorderColor = DP_DAY_SELECTED_BACK_COLOR
                    End With
                    'Apply selected-month text formatting
                        Lbl_Item.ForeColor = DP_DAY_SELECTED_FORE_COLOR
                End If
        Next Index

'------------------------------------------------------------------------------
' SHOW PANEL
'------------------------------------------------------------------------------
    'Show the panel above the day grid
        Fra_PickerPanel.Visible = True
    'Move the panel in front of the day grid
        Fra_PickerPanel.ZOrder 0

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
        Err.Raise Err.Number, PROC_NAME, "Month picker panel display failed: " & Err.Description

End Sub

Private Sub UF_PickerPanel_ShowYears()

'
'------------------------------------------------------------------------------
'                           SHOW YEAR PICKER PANEL
'------------------------------------------------------------------------------
' PURPOSE
'   Shows the reusable picker panel in year-selection mode
'
' WHY THIS EXISTS
'   Clicking the header year label or pressing Y should display a flat year
'   selector without using a native ComboBox
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures picker-panel cache is available, resets active hover states, hides
'   the settings panel, initializes the visible 12-year range around the
'   displayed year, populates the reusable picker items, and shows the picker
'   panel above the day grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if the picker panel cannot be resolved,
'   populated, or displayed
'
' DEPENDENCIES
'   UF_PickerPanel_EnsureCache
'   UF_PickerPanel_HoverReset
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'   UF_SettingsPanel_Hide
'   UF_DisplayPeriod_Initialize
'   UF_PickerPanel_PopulateYears
'
' NOTES
'   Year navigation arrows are handled by header labels, not by controls inside
'   the picker panel
'
'   The current displayed year is highlighted by UF_PickerPanel_PopulateYears
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_PickerPanel_ShowYears"

    Dim ExistingControl     As MSForms.control      'Existing control using the picker-panel name
    Dim Fra_PickerPanel     As MSForms.Frame        'Reusable picker panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Ensure cached picker-panel references are available
        UF_PickerPanel_EnsureCache

'------------------------------------------------------------------------------
' RESET ACTIVE HOVER STATE
'------------------------------------------------------------------------------
    'Clear active picker hover before changing picker-panel mode
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Clear active day-cell hover before showing the picker panel
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Clear active header-label hover before showing the picker panel
        If VBA.Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Clear active footer hover before showing the picker panel
        If VBA.Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset
    'Clear active settings-panel hover before showing the picker panel
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then UF_SettingsPanel_HoverReset

'------------------------------------------------------------------------------
' HIDE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Hide the settings panel before showing the year picker
        UF_SettingsPanel_Hide

'------------------------------------------------------------------------------
' ENSURE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Initialize the display period if the year state is not yet usable
        If mDisplayYear = 0 Then
            UF_DisplayPeriod_Initialize VBA.Date
        End If

'------------------------------------------------------------------------------
' RETRIEVE PICKER PANEL
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the picker panel
        On Error Resume Next
    'Retrieve the existing picker panel control
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject missing picker panel
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Unable to resolve expected picker panel " & DP_PICKER_PANEL_NAME
        End If
    'Reject a name collision with a non-frame control
        If VBA.TypeName(ExistingControl) <> "Frame" Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Control '" & DP_PICKER_PANEL_NAME & "' exists but is not an MSForms.Frame"
        End If

    'Use the resolved picker panel frame
        Set Fra_PickerPanel = ExistingControl

'------------------------------------------------------------------------------
' SET PANEL MODE
'------------------------------------------------------------------------------
    'Store year-panel mode
        mPickerPanelMode = 2
    'Initialize the first visible year around the displayed year
        mYearPanelStart = mDisplayYear - 5
    'Clamp the lower year-panel boundary
        If mYearPanelStart < DP_MIN_YEAR Then
            mYearPanelStart = DP_MIN_YEAR
        End If
    'Clamp the upper year-panel boundary
        If mYearPanelStart > DP_YEAR_PANEL_MAX_START Then
            mYearPanelStart = DP_YEAR_PANEL_MAX_START
        End If

'------------------------------------------------------------------------------
' POPULATE YEARS
'------------------------------------------------------------------------------
    'Populate the year labels
        UF_PickerPanel_PopulateYears

'------------------------------------------------------------------------------
' SHOW PANEL
'------------------------------------------------------------------------------
    'Show the panel above the day grid
        Fra_PickerPanel.Visible = True
    'Move the panel in front of the day grid
        Fra_PickerPanel.ZOrder 0

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
        Err.Raise Err.Number, PROC_NAME, "Year picker panel display failed: " & Err.Description

End Sub

Private Sub UF_PickerPanel_PopulateYears()

'
'------------------------------------------------------------------------------
'                           POPULATE YEAR PICKER PANEL
'------------------------------------------------------------------------------
' PURPOSE
'   Populates the reusable picker panel with 12 year values
'
' WHY THIS EXISTS
'   The picker panel is reused for month and year selection. Year values should
'   be refreshed without repeatedly resolving picker-panel label controls by
'   string name once the label cache is available
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures the picker-panel cache is available, normalizes the visible year
'   range, populates cached text and background labels, and highlights the
'   currently displayed year
'
' ERROR POLICY
'   Raises a descriptive runtime error if the picker-panel cache cannot be
'   resolved or if one or more cached labels are missing
'
' DEPENDENCIES
'   UF_PickerPanel_EnsureCache
'   mPickerTextLabels
'   mPickerBackLabels
'   mYearPanelStart
'   mDisplayYear
'
' NOTES
'   This routine does not show or hide the panel
'
'   Layering is established when the picker panel is built. This routine does
'   not repeat item-level ZOrder calls during population
'
'   Developer-owned picker-panel constants are intentionally not validated here.
'   They should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_PickerPanel_PopulateYears"
    
    Dim Lbl_Item            As MSForms.Label        'Picker item text label
    Dim Lbl_ItemBg          As MSForms.Label        'Picker item background label
    Dim Index               As Long                 'Panel item index
    Dim DisplayYear         As Long                 'Year displayed in item

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Ensure cached picker-panel references are available
        UF_PickerPanel_EnsureCache

'------------------------------------------------------------------------------
' NORMALIZE YEAR RANGE
'------------------------------------------------------------------------------
    'Initialize the year-panel start when missing
        If mYearPanelStart = 0 Then
            mYearPanelStart = mDisplayYear - 5
        End If
    'Clamp the lower year-panel boundary
        If mYearPanelStart < DP_MIN_YEAR Then
            mYearPanelStart = DP_MIN_YEAR
        End If
    'Clamp the upper year-panel boundary
        If mYearPanelStart > DP_YEAR_PANEL_MAX_START Then
            mYearPanelStart = DP_YEAR_PANEL_MAX_START
        End If

'------------------------------------------------------------------------------
' POPULATE YEARS
'------------------------------------------------------------------------------
    'Loop through the year picker items
        For Index = 1 To DP_PICKER_ITEM_COUNT

            'Resolve the year for this picker item
                DisplayYear = mYearPanelStart + Index - 1
            'Retrieve the cached picker item text label
                Set Lbl_Item = mPickerTextLabels(Index)
            'Retrieve the cached picker item background label
                Set Lbl_ItemBg = mPickerBackLabels(Index)
            'Reject a missing cached picker item text label
                If Lbl_Item Is Nothing Then
                    Err.Raise vbObjectError + 513, PROC_NAME, _
                        "Cached picker text label is missing for item " & VBA.CStr(Index)
                End If
            'Reject a missing cached picker item background label
                If Lbl_ItemBg Is Nothing Then
                    Err.Raise vbObjectError + 514, PROC_NAME, _
                        "Cached picker background label is missing for item " & VBA.CStr(Index)
                End If

'------------------------------------------------------------------------------
' APPLY YEAR TEXT STATE
'------------------------------------------------------------------------------
            'Apply the year caption and value
                With Lbl_Item
                    .Caption = VBA.CStr(DisplayYear)
                    .Tag = VBA.CStr(DisplayYear)
                    .Left = Lbl_ItemBg.Left
                    .Top = Lbl_ItemBg.Top + DP_PICKER_ITEM_TEXT_TOP_OFFSET
                    .Width = Lbl_ItemBg.Width
                    .Height = DP_PICKER_ITEM_TEXT_HEIGHT
                    .BackStyle = fmBackStyleTransparent
                    .ForeColor = DP_PICKER_ITEM_FORE_COLOR
                    .BorderStyle = fmBorderStyleNone
                    .SpecialEffect = fmSpecialEffectFlat
                    .TextAlign = fmTextAlignCenter
                    .WordWrap = False
                    .Enabled = True
                    .Visible = True
                End With

'------------------------------------------------------------------------------
' APPLY YEAR BACKGROUND STATE
'------------------------------------------------------------------------------
            'Fully reset the picker item background
                With Lbl_ItemBg
                    .Caption = vbNullString
                    .Tag = VBA.CStr(DisplayYear)
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = DP_PICKER_ITEM_BACK_COLOR
                    .ForeColor = DP_PICKER_ITEM_FORE_COLOR
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = DP_PICKER_ITEM_BORDER_COLOR
                    .SpecialEffect = fmSpecialEffectFlat
                    .Enabled = True
                    .Visible = True
                End With

'------------------------------------------------------------------------------
' APPLY SELECTED YEAR STATE
'------------------------------------------------------------------------------
            'Highlight the currently displayed year
                If DisplayYear = mDisplayYear Then
                    With Lbl_ItemBg
                        .BackColor = DP_DAY_SELECTED_BACK_COLOR
                        .BorderColor = DP_DAY_SELECTED_BACK_COLOR
                    End With
                    'Apply selected-year text formatting
                        Lbl_Item.ForeColor = DP_DAY_SELECTED_FORE_COLOR
                End If

        Next Index

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
        Err.Raise Err.Number, PROC_NAME, "Year picker panel population failed: " & Err.Description

End Sub

Public Sub UF_PickerPanel_HoverApply(ByVal LabelName As String)

'
'------------------------------------------------------------------------------
'                           APPLY PICKER PANEL HOVER
'------------------------------------------------------------------------------
' PURPOSE
'   Applies hover formatting to one month/year picker-panel item
'
' WHY THIS EXISTS
'   Picker-panel items are built from two runtime labels. Hovering either label
'   should highlight the full picker item consistently
'
' INPUTS
'   LabelName
'     Name of the picker-panel label currently under the mouse pointer
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the picker item index, exits when the same item is already hovered,
'   resets previous picker/footer hover state, retrieves the paired labels from
'   cache, resolves selected state, and applies hover formatting to the current
'   picker item
'
' ERROR POLICY
'   Raises a descriptive runtime error if LabelName is blank, if the label name
'   does not map to a picker-panel item, or if the target picker item cannot be
'   resolved or formatted
'
' DEPENDENCIES
'   UF_PickerPanel_ItemIndexFromLabelName
'   UF_PickerPanel_HoverReset
'   UF_PickerPanel_EnsureCache
'   UF_Footer_HoverReset
'   mPickerBackLabels
'   mPickerTextLabels
'
' NOTES
'   This routine is called by cDatePickerLabelHook when the hook category is
'   PICKER
'
'   The picker-panel cache is rebuilt only when the requested cached label pair
'   is missing
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_PickerPanel_HoverApply"
    
    Const PICKER_MODE_MONTHS    As Long = 1             'Month picker-panel mode
    Const PICKER_MODE_YEARS     As Long = 2             'Year picker-panel mode

    Dim EffectiveLabelName      As String               'Normalized label name
    Dim ItemIndex               As Long                 'Resolved picker item index
    Dim IsSelected              As Boolean              'True when hovered item is selected
    Dim CacheRefreshRequired    As Boolean              'True when cached labels must be rebuilt
    Dim RawTagValue             As String               'Picker text label Tag value
    
    Dim Lbl_ItemBg              As MSForms.Label        'Picker item background label
    Dim Lbl_Item                As MSForms.Label        'Picker item text label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Normalize the supplied label name
        EffectiveLabelName = VBA.Trim$(LabelName)
    'Reject an empty label name
        If VBA.Len(EffectiveLabelName) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "LabelName cannot be empty"
        End If

'------------------------------------------------------------------------------
' RESOLVE PICKER ITEM
'------------------------------------------------------------------------------
    'Resolve the picker item index from the hovered label name
        ItemIndex = UF_PickerPanel_ItemIndexFromLabelName(EffectiveLabelName)
    'Reject invalid picker item indexes
        If ItemIndex < 1 Or ItemIndex > DP_PICKER_ITEM_COUNT Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "LabelName does not resolve to a valid picker-panel item"
        End If

'------------------------------------------------------------------------------
' EXIT IF ALREADY HOVERED
'------------------------------------------------------------------------------
    'Exit if the requested item is already highlighted
        If mHoveredPickerItemIndex = ItemIndex Then Exit Sub

'------------------------------------------------------------------------------
' RESET OTHER HOVER STATE
'------------------------------------------------------------------------------
    'Remove footer hover when entering the picker panel
        If VBA.Len(mHoveredFooterActionName) <> 0 Then
            UF_Footer_HoverReset
        End If

'------------------------------------------------------------------------------
' RESET CURRENT PICKER HOVER STATE
'------------------------------------------------------------------------------
    'Remove hover formatting from the previously highlighted picker item
        If mHoveredPickerItemIndex <> 0 Then
            UF_PickerPanel_HoverReset
        End If

'------------------------------------------------------------------------------
' RETRIEVE CACHED PICKER LABELS
'------------------------------------------------------------------------------
    'Suppress cache probing errors
        On Error Resume Next
    'Retrieve the cached picker item background label
        Set Lbl_ItemBg = mPickerBackLabels(ItemIndex)
    'Retrieve the cached picker item text label
        Set Lbl_Item = mPickerTextLabels(ItemIndex)
    'Clear any suppressed cache probing error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESTORE CACHE WHEN NEEDED
'------------------------------------------------------------------------------
    'Request cache refresh when the background label is missing
        If Lbl_ItemBg Is Nothing Then CacheRefreshRequired = True
    'Request cache refresh when the text label is missing
        If Lbl_Item Is Nothing Then CacheRefreshRequired = True
    'Rebuild missing cached references only when needed
        If CacheRefreshRequired Then
            'Ensure cached picker-panel references are available
                UF_PickerPanel_EnsureCache
            'Retrieve the cached picker item background label after cache refresh
                Set Lbl_ItemBg = mPickerBackLabels(ItemIndex)
            'Retrieve the cached picker item text label after cache refresh
                Set Lbl_Item = mPickerTextLabels(ItemIndex)
        End If

'------------------------------------------------------------------------------
' VALIDATE CACHED LABELS
'------------------------------------------------------------------------------
    'Reject a missing cached picker background label
        If Lbl_ItemBg Is Nothing Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Cached picker background label is missing for item " & VBA.CStr(ItemIndex)
        End If
    'Reject a missing cached picker text label
        If Lbl_Item Is Nothing Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Cached picker text label is missing for item " & VBA.CStr(ItemIndex)
        End If

'------------------------------------------------------------------------------
' RESOLVE SELECTED STATE
'------------------------------------------------------------------------------
    'Resolve whether the hovered month item is currently selected
        If mPickerPanelMode = PICKER_MODE_MONTHS Then
            IsSelected = (ItemIndex = mDisplayMonth)
        End If
    'Resolve whether the hovered year item is currently selected
        If mPickerPanelMode = PICKER_MODE_YEARS Then
            'Read the picker item Tag safely
                RawTagValue = VBA.Trim$(VBA.CStr(Lbl_Item.Tag))
            'Evaluate numeric year tags only
                If VBA.Len(RawTagValue) <> 0 Then
                    If Not RawTagValue Like "*[!0-9]*" Then
                        IsSelected = (VBA.CLng(RawTagValue) = mDisplayYear)
                    End If
                End If
        End If

'------------------------------------------------------------------------------
' APPLY SELECTED HOVER STATE
'------------------------------------------------------------------------------
    'Preserve selected-item formatting when hovering the selected item
        If IsSelected Then
            'Apply selected picker-item background formatting
                With Lbl_ItemBg
                    .BackColor = DP_DAY_SELECTED_BACK_COLOR
                    .BorderColor = DP_DAY_SELECTED_BACK_COLOR
                    .SpecialEffect = fmSpecialEffectFlat
                End With
            'Apply selected picker-item text formatting
                Lbl_Item.ForeColor = DP_DAY_SELECTED_FORE_COLOR
            'Store the current hovered picker item index
                mHoveredPickerItemIndex = ItemIndex
            'Exit after applying selected hover state
                Exit Sub
        End If

'------------------------------------------------------------------------------
' APPLY NORMAL HOVER STATE
'------------------------------------------------------------------------------
    'Apply hover background formatting
        With Lbl_ItemBg
            .BackColor = DP_PICKER_ITEM_HOVER_BACK_COLOR
            .BorderColor = DP_PICKER_ITEM_HOVER_BORDER_COLOR
            .SpecialEffect = fmSpecialEffectFlat
        End With
    'Apply normal picker item text color
        Lbl_Item.ForeColor = DP_PICKER_ITEM_FORE_COLOR

'------------------------------------------------------------------------------
' STORE HOVER STATE
'------------------------------------------------------------------------------
    'Store the current hovered picker item index
        mHoveredPickerItemIndex = ItemIndex

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
        Err.Raise Err.Number, PROC_NAME, "Picker-panel hover application failed: " & Err.Description

End Sub

Private Sub UF_PickerPanel_HoverReset()

'
'------------------------------------------------------------------------------
'                           RESET PICKER PANEL HOVER
'------------------------------------------------------------------------------
' PURPOSE
'   Removes hover formatting from the currently highlighted picker-panel item
'
' WHY THIS EXISTS
'   Picker-panel hover formatting must be removed when the mouse moves away from
'   an item, when the picker panel is hidden, or when another item is highlighted
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Restores the currently highlighted picker item to its normal or selected
'   visual state and clears the current picker-panel hover tracker
'
' ERROR POLICY
'   Best-effort UI cleanup. Suppresses reset failures because hover reset should
'   never interrupt UserForm interaction
'
' DEPENDENCIES
'   UF_PickerPanel_EnsureCache
'   mPickerTextLabels
'   mPickerBackLabels
'
' NOTES
'   This routine resets only the currently hovered picker item, not all 12 items
'
'   The picker-panel cache is rebuilt only when the requested cached label pair
'   is missing
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PICKER_MODE_MONTHS    As Long = 1             'Month picker-panel mode
    Const PICKER_MODE_YEARS     As Long = 2             'Year picker-panel mode

    Dim HoveredIndex            As Long                 'Cached hovered picker item index
    Dim IsSelected              As Boolean              'True when hovered item is selected
    Dim CacheRefreshRequired    As Boolean              'True when cached labels must be rebuilt
    Dim RawTagValue             As String               'Picker text label Tag value
    
    Dim Lbl_ItemBg              As MSForms.Label        'Picker item background label
    Dim Lbl_Item                As MSForms.Label        'Picker item text label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' EXIT IF NOTHING IS HOVERED
'------------------------------------------------------------------------------
    'Exit if no picker item is currently highlighted
        If mHoveredPickerItemIndex = 0 Then Exit Sub
    'Capture the hovered index before any best-effort operation
        HoveredIndex = mHoveredPickerItemIndex

'------------------------------------------------------------------------------
' VALIDATE HOVER INDEX
'------------------------------------------------------------------------------
    'Clear hover state and exit when the stored hover index is no longer valid
        If HoveredIndex < 1 Or HoveredIndex > DP_PICKER_ITEM_COUNT Then
            GoTo Clean_Exit
        End If

'------------------------------------------------------------------------------
' RETRIEVE CACHED PICKER LABELS
'------------------------------------------------------------------------------
    'Suppress cache probing errors
        On Error Resume Next
    'Retrieve the cached picker item background label
        Set Lbl_ItemBg = mPickerBackLabels(HoveredIndex)
    'Retrieve the cached picker item text label
        Set Lbl_Item = mPickerTextLabels(HoveredIndex)
    'Clear any suppressed cache probing error
        Err.Clear
    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' RESTORE CACHE WHEN NEEDED
'------------------------------------------------------------------------------
    'Request cache refresh when the background label is missing
        If Lbl_ItemBg Is Nothing Then CacheRefreshRequired = True
    'Request cache refresh when the text label is missing
        If Lbl_Item Is Nothing Then CacheRefreshRequired = True
    'Rebuild missing cached references only when needed
        If CacheRefreshRequired Then
            'Ensure cached picker-panel references are available
                UF_PickerPanel_EnsureCache
            'Retrieve the cached picker item background label after cache refresh
                Set Lbl_ItemBg = mPickerBackLabels(HoveredIndex)
            'Retrieve the cached picker item text label after cache refresh
                Set Lbl_Item = mPickerTextLabels(HoveredIndex)
        End If

'------------------------------------------------------------------------------
' EXIT IF LABELS ARE UNAVAILABLE
'------------------------------------------------------------------------------
    'Clear hover state and exit if the cached labels are unavailable
        If Lbl_ItemBg Is Nothing Then GoTo Clean_Exit
    'Clear hover state and exit if the cached labels are unavailable
        If Lbl_Item Is Nothing Then GoTo Clean_Exit

'------------------------------------------------------------------------------
' RESOLVE SELECTED STATE
'------------------------------------------------------------------------------
    'Resolve whether the hovered month item is currently selected
        If mPickerPanelMode = PICKER_MODE_MONTHS Then
            IsSelected = (HoveredIndex = mDisplayMonth)
        End If
    'Resolve whether the hovered year item is currently selected
        If mPickerPanelMode = PICKER_MODE_YEARS Then
            'Read the picker item Tag safely
                RawTagValue = VBA.Trim$(VBA.CStr(Lbl_Item.Tag))
            'Evaluate numeric year tags only
                If VBA.Len(RawTagValue) <> 0 Then
                    If Not RawTagValue Like "*[!0-9]*" Then
                        IsSelected = (VBA.CLng(RawTagValue) = mDisplayYear)
                    End If
                End If
        End If

'------------------------------------------------------------------------------
' RESTORE NORMAL BACKGROUND STATE
'------------------------------------------------------------------------------
    'Restore the normal picker-item background state
        With Lbl_ItemBg
            .Caption = vbNullString
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_PICKER_ITEM_BACK_COLOR
            .ForeColor = DP_PICKER_ITEM_FORE_COLOR
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = DP_PICKER_ITEM_BORDER_COLOR
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = True
            .Visible = True
        End With

'------------------------------------------------------------------------------
' RESTORE NORMAL TEXT STATE
'------------------------------------------------------------------------------
    'Restore the normal picker-item text state
        With Lbl_Item
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_PICKER_ITEM_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = True
            .Visible = True
        End With

'------------------------------------------------------------------------------
' RESTORE SELECTED STATE
'------------------------------------------------------------------------------
    'Restore selected-item formatting when applicable
        If IsSelected Then
            With Lbl_ItemBg
                .BackColor = DP_DAY_SELECTED_BACK_COLOR
                .BorderColor = DP_DAY_SELECTED_BACK_COLOR
            End With
            'Apply selected picker-item text formatting
                Lbl_Item.ForeColor = DP_DAY_SELECTED_FORE_COLOR
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:
    'Clear the current hovered picker item index
        mHoveredPickerItemIndex = 0
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Suppress secondary cleanup errors
        On Error Resume Next
    'Clear the current hovered picker item index
        mHoveredPickerItemIndex = 0
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Function UF_PickerPanel_ItemIndexFromLabelName(ByVal LabelName As String) As Long

'
'------------------------------------------------------------------------------
'                    GET PICKER ITEM INDEX FROM LABEL NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the picker-panel item index from a picker-panel label name
'
' WHY THIS EXISTS
'   Each picker-panel item is composed of two labels:
'     - a background label
'     - a text label
'
'   Hovering or clicking either label must resolve to the same picker item index
'
' INPUTS
'   LabelName
'     Picker-panel label name
'
' RETURNS
'   Picker-panel item index from 1 to DP_PICKER_ITEM_COUNT
'
'   Zero when the label name is blank, unsupported, malformed, or outside the
'   supported picker-item range
'
' BEHAVIOR
'   Normalizes the supplied label name, checks the background-label prefix first,
'   checks the text-label prefix second, extracts the numeric suffix, applies
'   strict digit-only parsing, and returns the validated picker-item index
'
' ERROR POLICY
'   Does not raise errors. Unsupported or malformed label names return zero
'
' DEPENDENCIES
'   DP_PICKER_ITEM_COUNT
'
' NOTES
'   The background-label prefix must be tested before the text-label prefix
'   because Lbl_MonthYearBg also starts with Lbl_MonthYear
'
'   Matching is normalized with UCase$ so behavior is independent from
'   Option Compare
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const BG_PREFIX         As String = "LBL_MONTHYEARBG"       'Picker background label prefix
    Const TEXT_PREFIX       As String = "LBL_MONTHYEAR"         'Picker text label prefix

    Dim EffectiveName       As String                           'Normalized picker label name
    Dim RawIndex            As String                           'Raw numeric suffix
    Dim ParsedIndex         As Long                             'Parsed picker item index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe parsing
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the supplied label name
        EffectiveName = VBA.UCase$(VBA.Trim$(LabelName))
    'Return zero for empty label names
        If VBA.Len(EffectiveName) = 0 Then
            UF_PickerPanel_ItemIndexFromLabelName = 0
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE INDEX FROM BACKGROUND LABEL
'------------------------------------------------------------------------------
    'Resolve indexes from picker background labels first
        If VBA.Left$(EffectiveName, VBA.Len(BG_PREFIX)) = BG_PREFIX Then
            'Extract the raw numeric suffix
                RawIndex = VBA.Mid$(EffectiveName, VBA.Len(BG_PREFIX) + 1)
            'Return zero when the suffix is empty
                If VBA.Len(RawIndex) = 0 Then Exit Function
            'Return zero when the suffix is not strictly numeric
                If RawIndex Like "*[!0-9]*" Then Exit Function
            'Parse the picker item index
                ParsedIndex = VBA.CLng(RawIndex)
            'Return zero when the index is outside the supported picker range
                If ParsedIndex < 1 Or ParsedIndex > DP_PICKER_ITEM_COUNT Then Exit Function
            'Return the validated picker item index
                UF_PickerPanel_ItemIndexFromLabelName = ParsedIndex
            'Exit after resolving the background label
                Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE INDEX FROM TEXT LABEL
'------------------------------------------------------------------------------
    'Resolve indexes from picker text labels
        If VBA.Left$(EffectiveName, VBA.Len(TEXT_PREFIX)) = TEXT_PREFIX Then
            'Extract the raw numeric suffix
                RawIndex = VBA.Mid$(EffectiveName, VBA.Len(TEXT_PREFIX) + 1)
            'Return zero when the suffix is empty
                If VBA.Len(RawIndex) = 0 Then Exit Function
            'Return zero when the suffix is not strictly numeric
                If RawIndex Like "*[!0-9]*" Then Exit Function
            'Parse the picker item index
                ParsedIndex = VBA.CLng(RawIndex)
            'Return zero when the index is outside the supported picker range
                If ParsedIndex < 1 Or ParsedIndex > DP_PICKER_ITEM_COUNT Then Exit Function
            'Return the validated picker item index
                UF_PickerPanel_ItemIndexFromLabelName = ParsedIndex
            'Exit after resolving the text label
                Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN FALLBACK
'------------------------------------------------------------------------------
    'Return zero for unsupported picker label names
        UF_PickerPanel_ItemIndexFromLabelName = 0

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit the function
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Return zero for malformed or oversized numeric suffixes
        UF_PickerPanel_ItemIndexFromLabelName = 0

End Function


'
'------------------------------------------------------------------------------
'
'                                SETTINGS PANEL
'
'------------------------------------------------------------------------------
'

Private Sub UF_Settings_Show()

'
'------------------------------------------------------------------------------
'                           SHOW SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Shows the in-form DatePicker settings panel
'
' WHY THIS EXISTS
'   The DatePicker exposes a Settings entry point. Settings should be displayed
'   inside the DatePicker itself instead of opening a separate modal form
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resets active hover states, hides the month/year picker panel when visible,
'   clears picker-panel transient state, and shows the in-form settings panel
'
' ERROR POLICY
'   Raises a descriptive runtime error if the settings panel cannot be shown
'
' DEPENDENCIES
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_PickerPanel_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'   UF_SettingsPanel_Show
'
' NOTES
'   This replaces the previous Frm_DpSettings.Show vbModal behavior
'
'   Settings panel geometry and settings-control refresh are delegated to
'   UF_SettingsPanel_Show
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Settings_Show"

    Dim ExistingControl     As MSForms.control      'Existing picker-panel control
    Dim Fra_PickerPanel     As MSForms.Frame        'Reusable picker panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESET ACTIVE HOVER STATE
'------------------------------------------------------------------------------
    'Clear active day-cell hover before showing settings
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Clear active header-label hover before showing settings
        If VBA.Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Clear active picker-panel hover before showing settings
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Clear active footer hover before showing settings
        If VBA.Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset
    'Clear active settings-panel hover before showing settings
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then UF_SettingsPanel_HoverReset

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the picker panel
        On Error Resume Next
    'Retrieve the existing picker panel control
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Hide the picker panel when it exists and is a frame
        If Not ExistingControl Is Nothing Then
            'Use the picker panel only when the resolved control is a frame
                If VBA.TypeName(ExistingControl) = "Frame" Then
                    'Store the picker panel reference
                        Set Fra_PickerPanel = ExistingControl
                    'Hide the picker panel before showing settings
                        Fra_PickerPanel.Visible = False
                End If
        End If
    'Clear picker-panel mode
        mPickerPanelMode = 0
    'Clear picker-panel hover state
        mHoveredPickerItemIndex = 0

'------------------------------------------------------------------------------
' SHOW SETTINGS PANEL
'------------------------------------------------------------------------------
    'Show the in-form settings panel
        UF_SettingsPanel_Show

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
        Err.Raise Err.Number, PROC_NAME, "Settings panel display failed: " & Err.Description

End Sub

Private Sub UF_SettingsPanel_Build()

'
'------------------------------------------------------------------------------
'                           BUILD SETTINGS PANEL
'------------------------------------------------------------------------------
' PURPOSE
'   Creates the reusable in-form settings panel
'
' WHY THIS EXISTS
'   Settings should be displayed as an overlay inside the DatePicker form using
'   the same size and position as Fra_PickerPanel
'
'   The settings content is organized through a runtime MultiPage control so
'   display, behavior, and integration settings can evolve without requiring
'   separate forms or modal dialogs
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses Fra_Settings, the runtime MultiPage, the Display,
'   Behavior, and Integration pages, page backgrounds, fake tab labels, header
'   labels, Save / Close actions, first-day ComboBox, and all settings CheckBox
'   controls
'
' ERROR POLICY
'   Raises a descriptive runtime error if the settings panel, MultiPage, pages,
'   labels, ComboBox, CheckBox controls, or hooks cannot be created or formatted
'
' DEPENDENCIES
'   UF_Ensure_FrameLabel
'   UF_Ensure_FrameMultiPage
'   UF_Ensure_PageLabel
'   UF_Ensure_PageComboBox
'   UF_Ensure_PageCheckBox
'   UF_CalendarGrid_GetWidth
'   UF_SettingsTabs_Build
'   UF_SettingsTabs_RefreshVisualState
'   UF_SettingsPage_BackgroundBuild
'   UF_SettingsPanel_RefreshCaptions
'   UF_SettingsCheckBox_Format
'   cDatePickerLabelHook
'
' NOTES
'   This routine builds the visual settings surface only
'
'   Individual MSForms.Page objects are not assigned BackColor because Excel
'   UserForm pages do not reliably expose a writable BackColor property
'
'   Runtime ComboBox and CheckBox values are persisted by the explicit Save
'   action handled by UF_SettingsPanel_Save
'
'   Developer-owned settings layout constants are intentionally not validated
'   here. They should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_SettingsPanel_Build"   'Current procedure name

    Dim ExistingControl         As MSForms.control                                  'Existing control using the settings-panel name
    Dim Fra_Settings            As MSForms.Frame                                    'Reusable settings panel
    Dim Mp_Settings             As MSForms.MultiPage                                'Settings MultiPage control
    
    Dim Pge_Display             As MSForms.Page                                     'Display settings page
    Dim Pge_Behavior            As MSForms.Page                                     'Behavior settings page
    Dim Pge_Integration         As MSForms.Page                                     'Integration settings page
    
    Dim PanelLeft               As Single                                           'Settings panel left position
    Dim PanelTop                As Single                                           'Settings panel top position
    
    Dim Lbl_Title               As MSForms.Label                                    'Settings title label
    Dim Lbl_Save                As MSForms.Label                                    'Settings save label
    Dim Lbl_Close               As MSForms.Label                                    'Settings close label
    Dim Lbl_FirstDay            As MSForms.Label                                    'First-day label
    Dim Cbo_FirstDay            As MSForms.ComboBox                                 'First-day ComboBox
    Dim Chk_LocalNames          As MSForms.CheckBox                                 'Use-local-names checkbox
    Dim Chk_LiveClock           As MSForms.CheckBox                                 'Live-clock checkbox
    Dim Chk_Compact             As MSForms.CheckBox                                 'Compact-layout checkbox
    Dim Chk_Weekends            As MSForms.CheckBox                                 'Highlight-weekends checkbox
    Dim Chk_AllowOutside        As MSForms.CheckBox                                 'Allow outside-month selection checkbox
    Dim Chk_CloseAfter          As MSForms.CheckBox                                 'Close-after-selection checkbox
    Dim Chk_RightClick          As MSForms.CheckBox                                 'Right-click menu checkbox
    Dim Chk_InGridIcon          As MSForms.CheckBox                                 'In-grid icon checkbox
    Dim Chk_WinAPIStyle         As MSForms.CheckBox                                 'WinAPI styling checkbox
    
    Dim LabelHook               As cDatePickerLabelHook                             'Runtime label hook
    
    Dim TitleFont               As Object                                           'Clean title font object
    Dim BodyFont                As Object                                           'Clean body font object
    Dim SaveIconFont            As Object                                           'Clean save icon font object
    Dim CloseIconFont           As Object                                           'Clean close icon font object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Create or reset settings-panel click / hover hooks
        Set mSettingsPanelHooks = New Collection
    'Clear any stale settings-panel hover tracker before rebuilding controls
        mHoveredSettingsPanelLabelName = vbNullString

'------------------------------------------------------------------------------
' CALCULATE PANEL GEOMETRY
'------------------------------------------------------------------------------
    'Calculate the same left position used by Fra_PickerPanel
        PanelLeft = DP_DAY_GRID_START_LEFT + _
            ((UF_CalendarGrid_GetWidth() - DP_PICKER_PANEL_WIDTH) / 2)
    'Calculate the same top position used by Fra_PickerPanel
        PanelTop = DP_DAY_GRID_START_TOP - 7

'------------------------------------------------------------------------------
' CREATE / RETRIEVE SETTINGS FRAME
'------------------------------------------------------------------------------
    'Suppress lookup errors while checking for an existing settings panel
        On Error Resume Next
    'Try to retrieve an existing control with the settings-panel name
        Set ExistingControl = Me.Controls(DP_SETTINGS_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Create the settings panel if it does not already exist
        If ExistingControl Is Nothing Then
            'Create the reusable settings panel frame
                Set Fra_Settings = Me.Controls.Add("Forms.Frame.1", DP_SETTINGS_PANEL_NAME, True)
        Else
            'Reject a name collision with a non-frame control
                If VBA.TypeName(ExistingControl) <> "Frame" Then
                    Err.Raise vbObjectError + 513, PROC_NAME, _
                        "Control '" & DP_SETTINGS_PANEL_NAME & "' exists but is not an MSForms.Frame"
                End If
            'Use the existing settings panel frame
                Set Fra_Settings = ExistingControl
        End If

'------------------------------------------------------------------------------
' FORMAT SETTINGS FRAME
'------------------------------------------------------------------------------
    'Apply layout and visual properties to the settings panel
        With Fra_Settings
            .Caption = vbNullString
            .Width = DP_PICKER_PANEL_WIDTH
            .Height = DP_PICKER_PANEL_HEIGHT
            .Left = PanelLeft
            .Top = PanelTop
            .BackColor = DP_SETTINGS_PANEL_BACK_COLOR
            .BorderStyle = fmBorderStyleSingle
            .SpecialEffect = fmSpecialEffectFlat
            .Visible = False
        End With

'------------------------------------------------------------------------------
' HIDE LEGACY SETTINGS CONTROLS
'------------------------------------------------------------------------------
    'Temporarily ignore missing legacy labels
        On Error Resume Next

    'Hide the legacy labels if present
        Fra_Settings.Controls("Lbl_SettingsOption1").Visible = False
        Fra_Settings.Controls("Lbl_SettingsOption2").Visible = False
        Fra_Settings.Controls("Lbl_SettingsOption3").Visible = False
        Fra_Settings.Controls("Lbl_SettingsOption4").Visible = False
        Fra_Settings.Controls("Lbl_SettingsHint").Visible = False

    'Clear any suppressed legacy-control error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CREATE FONT OBJECTS
'------------------------------------------------------------------------------
    'Create a clean title font object
        Set TitleFont = CreateObject("StdFont")
    'Configure the title font
        With TitleFont
            .Name = DP_FORM_FONT_NAME
            .Size = 10
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Create a clean body font object
        Set BodyFont = CreateObject("StdFont")
    'Configure the body font
        With BodyFont
            .Name = DP_FORM_FONT_NAME
            .Size = DP_FORM_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

'------------------------------------------------------------------------------
' CREATE / FORMAT TITLE LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the settings title label
        Set Lbl_Title = UF_Ensure_FrameLabel(Fra_Settings, "Lbl_SettingsTitle")

    'Apply layout and visual properties to the settings title label
        With Lbl_Title
            .Caption = DP_SETTINGS_TITLE_CAPTION
            .Left = DP_SETTINGS_TITLE_LEFT
            .Top = DP_SETTINGS_TITLE_TOP
            .Width = DP_SETTINGS_TITLE_WIDTH
            .Height = DP_SETTINGS_TITLE_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_HEADER_BACK_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
        End With

    'Assign the clean title font
        Set Lbl_Title.Font = TitleFont
    'Create a settings-title hover-reset hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the title label to settings-panel hover reset behavior
        LabelHook.Initialize Me, Lbl_Title, vbNullString, "SETTINGS"
    'Store the hook so that the MouseMove event remains alive
        mSettingsPanelHooks.Add LabelHook, Lbl_Title.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT SAVE LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the settings save label
        Set Lbl_Save = UF_Ensure_FrameLabel(Fra_Settings, "Lbl_SettingsSave")

    'Apply layout and visual properties to the settings save label
        With Lbl_Save
            .Caption = vbNullString
            .Left = DP_SETTINGS_SAVE_LEFT
            .Top = DP_SETTINGS_HEADER_ICON_TOP - 1
            .Width = DP_SETTINGS_SAVE_WIDTH
            .Height = DP_SETTINGS_SAVE_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_SETTINGS_HEADER_ICON_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .MousePointer = fmMousePointerCustom
            .ControlTipText = DP_SETTINGS_SAVE_TOOLTIP
        End With

    'Create a clean save icon font object
        Set SaveIconFont = CreateObject("StdFont")
    'Configure the settings save icon font
        With SaveIconFont
            .Name = DP_SETTINGS_HEADER_ICON_FONT_NAME
            .Size = DP_SETTINGS_SAVE_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With
    'Assign the clean icon font to the settings save label
        Set Lbl_Save.Font = SaveIconFont
    'Apply the Segoe MDL2 save glyph
        Lbl_Save.Caption = ChrW$(DP_SETTINGS_SAVE_CODEPOINT)
    'Create the settings save click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the save label to the save-settings action
        LabelHook.Initialize Me, Lbl_Save, "SAVE_SETTINGS", "SETTINGS"
    'Store the hook so that the click / hover events remain alive
        mSettingsPanelHooks.Add LabelHook, Lbl_Save.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT CLOSE LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the settings close label
        Set Lbl_Close = UF_Ensure_FrameLabel(Fra_Settings, "Lbl_SettingsClose")
    'Apply layout and visual properties to the settings close label
        With Lbl_Close
            .Caption = vbNullString
            .Left = DP_SETTINGS_CLOSE_LEFT
            .Top = DP_SETTINGS_HEADER_ICON_TOP
            .Width = DP_SETTINGS_HEADER_ICON_WIDTH
            .Height = DP_SETTINGS_HEADER_ICON_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_SETTINGS_HEADER_ICON_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .MousePointer = fmMousePointerCustom
            .ControlTipText = DP_SETTINGS_CLOSE_TOOLTIP
        End With

    'Create a clean close icon font object
        Set CloseIconFont = CreateObject("StdFont")
    'Configure the settings close icon font
        With CloseIconFont
            .Name = DP_SETTINGS_HEADER_ICON_FONT_NAME
            .Size = DP_SETTINGS_CLOSE_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With
    'Assign the clean icon font to the settings close label
        Set Lbl_Close.Font = CloseIconFont
    'Apply the Segoe MDL2 close glyph
        Lbl_Close.Caption = ChrW$(DP_SETTINGS_CLOSE_CODEPOINT)
    'Create the settings close click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the close label to the hide-settings action
        LabelHook.Initialize Me, Lbl_Close, "HIDE_SETTINGS", "SETTINGS"
    'Store the hook so that the click / hover events remain alive
        mSettingsPanelHooks.Add LabelHook, Lbl_Close.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT MULTIPAGE
'------------------------------------------------------------------------------
    'Create or retrieve the settings MultiPage
        Set Mp_Settings = UF_Ensure_FrameMultiPage(Fra_Settings, DP_SETTINGS_MULTIPAGE_NAME)
    'Apply layout and visual properties to the settings MultiPage
        With Mp_Settings
            .Left = DP_SETTINGS_MULTIPAGE_LEFT
            .Top = DP_SETTINGS_MULTIPAGE_TOP
            .Width = DP_SETTINGS_MULTIPAGE_WIDTH
            .Height = DP_SETTINGS_MULTIPAGE_HEIGHT
            .Style = fmTabStyleNone
            .TabOrientation = fmTabOrientationTop
            .MultiRow = False
            .Value = 0
            .BackColor = DP_SETTINGS_PAGE_BACK_COLOR
            .ForeColor = DP_FOOTER_FORE_COLOR
            .Visible = True
            .Enabled = True
        End With
    'Assign the clean body font to the MultiPage
        Set Mp_Settings.Font = BodyFont

'------------------------------------------------------------------------------
' ENSURE MULTIPAGE PAGE COUNT
'------------------------------------------------------------------------------
    'Add pages until the MultiPage has three pages
        Do While Mp_Settings.Pages.Count < 3
            Mp_Settings.Pages.Add
        Loop
    'Remove extra pages if a previous build created more than three
        Do While Mp_Settings.Pages.Count > 3
            Mp_Settings.Pages.Remove Mp_Settings.Pages.Count - 1
        Loop

'------------------------------------------------------------------------------
' RETRIEVE MULTIPAGE PAGES
'------------------------------------------------------------------------------
    'Retrieve Settings pages
        Set Pge_Display = Mp_Settings.Pages(0)
        Set Pge_Behavior = Mp_Settings.Pages(1)
        Set Pge_Integration = Mp_Settings.Pages(2)

'------------------------------------------------------------------------------
' FORMAT MULTIPAGE PAGE CAPTIONS
'------------------------------------------------------------------------------
    'Apply the Settings page caption
        Pge_Display.Caption = DP_SETTINGS_PAGE_DISPLAY_CAPTION
        Pge_Behavior.Caption = DP_SETTINGS_PAGE_BEHAVIOR_CAPTION
        Pge_Integration.Caption = DP_SETTINGS_PAGE_INTEGRATION_CAPTION

'------------------------------------------------------------------------------
' CREATE / FORMAT FAKE SETTINGS TABS
'------------------------------------------------------------------------------
    'Create the label-based settings tab strip
        UF_SettingsTabs_Build Fra_Settings, BodyFont

'------------------------------------------------------------------------------
' BUILD PAGE BACKGROUNDS
'------------------------------------------------------------------------------
    'Build the Settings page background
        UF_SettingsPage_BackgroundBuild Mp_Settings, 0, "Lbl_DisplaySettingsBack"
        UF_SettingsPage_BackgroundBuild Mp_Settings, 1, "Lbl_BehaviorSettingsBack"
        UF_SettingsPage_BackgroundBuild Mp_Settings, 2, "Lbl_IntegrationSettingsBack"

'------------------------------------------------------------------------------
' HIDE LEGACY PAGE CONTROLS
'------------------------------------------------------------------------------
    'Temporarily ignore missing legacy controls
        On Error Resume Next
    'Hide the old behavior placeholder when present
        Pge_Behavior.Controls("Lbl_SettingsBehaviorPlaceholder").Visible = False
    'Hide the old integration placeholder when present
        Pge_Integration.Controls("Lbl_SettingsIntegrationPlaceholder").Visible = False
    'Clear any suppressed legacy-control error
        Err.Clear

    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CREATE / FORMAT FIRST-DAY LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the first-day label
        Set Lbl_FirstDay = UF_Ensure_PageLabel(Pge_Display, "Lbl_SettingsFirstDay")
    'Apply layout and visual properties to the first-day label
        With Lbl_FirstDay
            .Caption = DP_SETTINGS_FIRSTDAY_LABEL_CAPTION
            .Left = DP_SETTINGS_FIRSTDAY_LABEL_LEFT
            .Top = DP_SETTINGS_FIRSTDAY_LABEL_TOP
            .Width = DP_SETTINGS_FIRSTDAY_LABEL_WIDTH
            .Height = DP_SETTINGS_FIRSTDAY_LABEL_HEIGHT
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_FOOTER_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignLeft
            .WordWrap = False
            .Enabled = True
            .Visible = True
        End With
    'Assign the clean body font to the first-day label
        Set Lbl_FirstDay.Font = BodyFont

'------------------------------------------------------------------------------
' CREATE / FORMAT FIRST-DAY COMBOBOX
'------------------------------------------------------------------------------
    'Create or retrieve the first-day ComboBox
        Set Cbo_FirstDay = UF_Ensure_PageComboBox(Pge_Display, "Cbo_SettingsFirstDay")
    'Apply layout and visual properties to the first-day ComboBox
        With Cbo_FirstDay
            .Left = DP_SETTINGS_FIRSTDAY_COMBO_LEFT
            .Top = DP_SETTINGS_FIRSTDAY_COMBO_TOP
            .Width = DP_SETTINGS_FIRSTDAY_COMBO_WIDTH
            .Height = DP_SETTINGS_FIRSTDAY_COMBO_HEIGHT
            .BackColor = vbWhite
            .ForeColor = DP_FOOTER_FORE_COLOR
            .BorderStyle = fmBorderStyleSingle
            .SpecialEffect = fmSpecialEffectFlat
            .Style = fmStyleDropDownList
            .MatchRequired = True
            .Enabled = True
            .Visible = True
        End With
    'Assign the clean body font to the first-day ComboBox
        Set Cbo_FirstDay.Font = BodyFont

'------------------------------------------------------------------------------
' CREATE / FORMAT DISPLAY CHECKBOXES
'------------------------------------------------------------------------------
    'Create or retrieve the local-names checkbox
        Set Chk_LocalNames = UF_Ensure_PageCheckBox(Pge_Display, "Chk_SettingsUseLocalNames")
    'Apply deterministic CheckBox formatting
        UF_SettingsCheckBox_Format _
            Chk_LocalNames, _
            DP_SETTINGS_CHECK_LOCAL_NAMES_CAPTION, _
            DP_SETTINGS_CHECKBOX_LEFT, _
            DP_SETTINGS_CHECKBOX_TOP_1
    
    'Create or retrieve the live-clock checkbox
        Set Chk_LiveClock = UF_Ensure_PageCheckBox(Pge_Display, "Chk_SettingsLiveClock")
    'Apply deterministic CheckBox formatting
        UF_SettingsCheckBox_Format _
            Chk_LiveClock, _
            DP_SETTINGS_CHECK_LIVE_CLOCK_CAPTION, _
            DP_SETTINGS_CHECKBOX_LEFT, _
            DP_SETTINGS_CHECKBOX_TOP_1 + DP_SETTINGS_CHECKBOX_VERTICAL_STEP

    'Create or retrieve the compact-layout checkbox
        Set Chk_Compact = UF_Ensure_PageCheckBox(Pge_Display, "Chk_SettingsCompactLayout")
    'Apply deterministic CheckBox formatting
        UF_SettingsCheckBox_Format _
            Chk_Compact, _
            DP_SETTINGS_CHECK_COMPACT_CAPTION, _
            DP_SETTINGS_CHECKBOX_LEFT, _
            DP_SETTINGS_CHECKBOX_TOP_1 + (DP_SETTINGS_CHECKBOX_VERTICAL_STEP * 2)

    'Create or retrieve the weekend-highlight checkbox
        Set Chk_Weekends = UF_Ensure_PageCheckBox(Pge_Display, "Chk_SettingsHighlightWeekends")
    'Apply deterministic CheckBox formatting
        UF_SettingsCheckBox_Format _
            Chk_Weekends, _
            DP_SETTINGS_CHECK_WEEKENDS_CAPTION, _
            DP_SETTINGS_CHECKBOX_LEFT, _
            DP_SETTINGS_CHECKBOX_TOP_1 + (DP_SETTINGS_CHECKBOX_VERTICAL_STEP * 3)

'------------------------------------------------------------------------------
' CREATE / FORMAT BEHAVIOR CHECKBOXES
'------------------------------------------------------------------------------
    'Create or retrieve the allow-outside-month checkbox
        Set Chk_AllowOutside = UF_Ensure_PageCheckBox(Pge_Behavior, "Chk_SettingsAllowOutsideMonth")
    'Apply deterministic CheckBox formatting
        UF_SettingsCheckBox_Format _
            Chk_AllowOutside, _
            DP_SETTINGS_CHECK_ALLOW_OUTSIDE_CAPTION, _
            DP_SETTINGS_CHECKBOX_LEFT, _
            DP_SETTINGS_CHECKBOX_TOP_1

    'Create or retrieve the close-after-selection checkbox
        Set Chk_CloseAfter = UF_Ensure_PageCheckBox(Pge_Behavior, "Chk_SettingsCloseAfterSelection")
    'Apply deterministic CheckBox formatting
        UF_SettingsCheckBox_Format _
            Chk_CloseAfter, _
            DP_SETTINGS_CHECK_CLOSE_AFTER_CAPTION, _
            DP_SETTINGS_CHECKBOX_LEFT, _
            DP_SETTINGS_CHECKBOX_TOP_1 + DP_SETTINGS_CHECKBOX_VERTICAL_STEP

'------------------------------------------------------------------------------
' CREATE / FORMAT INTEGRATION CHECKBOXES
'------------------------------------------------------------------------------
    'Create or retrieve the right-click menu checkbox
        Set Chk_RightClick = UF_Ensure_PageCheckBox(Pge_Integration, "Chk_SettingsRightClickMenu")
    'Apply deterministic CheckBox formatting
        UF_SettingsCheckBox_Format _
            Chk_RightClick, _
            DP_SETTINGS_CHECK_RIGHT_CLICK_CAPTION, _
            DP_SETTINGS_CHECKBOX_LEFT, _
            DP_SETTINGS_CHECKBOX_TOP_1

    'Create or retrieve the in-grid icon checkbox
        Set Chk_InGridIcon = UF_Ensure_PageCheckBox(Pge_Integration, "Chk_SettingsInGridIcon")
    'Apply deterministic CheckBox formatting
        UF_SettingsCheckBox_Format _
            Chk_InGridIcon, _
            DP_SETTINGS_CHECK_IN_GRID_ICON_CAPTION, _
            DP_SETTINGS_CHECKBOX_LEFT, _
            DP_SETTINGS_CHECKBOX_TOP_1 + DP_SETTINGS_CHECKBOX_VERTICAL_STEP

    'Create or retrieve the WinAPI styling checkbox
        Set Chk_WinAPIStyle = UF_Ensure_PageCheckBox(Pge_Integration, "Chk_SettingsWinAPIStyling")
    'Apply deterministic CheckBox formatting
        UF_SettingsCheckBox_Format _
            Chk_WinAPIStyle, _
            DP_SETTINGS_CHECK_WINAPI_STYLE_CAPTION, _
            DP_SETTINGS_CHECKBOX_LEFT, _
            DP_SETTINGS_CHECKBOX_TOP_1 + (DP_SETTINGS_CHECKBOX_VERTICAL_STEP * 2)

'------------------------------------------------------------------------------
' POPULATE CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Refresh the settings panel values
        UF_SettingsPanel_RefreshCaptions

'------------------------------------------------------------------------------
' RESTORE LAYERING
'------------------------------------------------------------------------------
    'Move the MultiPage to the front of its layer
        Mp_Settings.ZOrder 0
    'Refresh the fake settings tab visual state after moving the MultiPage
        UF_SettingsTabs_RefreshVisualState Mp_Settings.Value
    'Move the settings title to the front
        Lbl_Title.ZOrder 0
    'Move the settings save label above the title label
        Lbl_Save.ZOrder 0
    'Move the settings close label above the title label
        Lbl_Close.ZOrder 0
    'Move the settings panel in front of the day grid
        Fra_Settings.ZOrder 0
    'Keep the settings panel hidden after build
        Fra_Settings.Visible = False

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
        Err.Raise Err.Number, PROC_NAME, "Settings panel creation failed: " & Err.Description

End Sub

Private Sub UF_SettingsPanel_Show()

'
'------------------------------------------------------------------------------
'                           SHOW SETTINGS PANEL
'------------------------------------------------------------------------------
' PURPOSE
'   Shows the reusable settings panel inside the DatePicker form
'
' WHY THIS EXISTS
'   The settings panel shares the same overlay surface used by the month/year
'   picker panel and should be shown without opening a separate modal form
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures the settings panel exists, resets transient hover state, hides the
'   picker panel, refreshes current settings values, refreshes fake tab visual
'   state, and shows Fra_Settings above the day grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if the settings panel cannot be created,
'   resolved, refreshed, or shown
'
' DEPENDENCIES
'   UF_SettingsPanel_Build
'   UF_SettingsPanel_RefreshCaptions
'   UF_SettingsTabs_RefreshVisualState
'   UF_PickerPanel_HoverReset
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'
' NOTES
'   This routine does not open a separate UserForm
'
'   The routine is intentionally self-contained because it may be called either
'   directly or through UF_Settings_Show
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_SettingsPanel_Show"
    
    Dim ExistingControl     As MSForms.control              'Existing control resolved by name
    Dim Fra_Settings        As MSForms.Frame                'Reusable settings panel
    Dim Fra_PickerPanel     As MSForms.Frame                'Reusable picker panel
    Dim Mp_Settings         As MSForms.MultiPage            'Settings MultiPage control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' ENSURE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Suppress lookup errors when the settings panel does not exist yet
        On Error Resume Next
    'Try to retrieve the existing settings panel
        Set ExistingControl = Me.Controls(DP_SETTINGS_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Build the settings panel if it has not been created yet
        If ExistingControl Is Nothing Then
            'Build the reusable in-form settings panel
                UF_SettingsPanel_Build
            'Retrieve the newly built settings panel
                Set ExistingControl = Me.Controls(DP_SETTINGS_PANEL_NAME)
        End If
    'Reject a missing settings panel after build
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Unable to resolve expected settings panel " & DP_SETTINGS_PANEL_NAME
        End If
    'Reject a name collision with a non-frame control
        If VBA.TypeName(ExistingControl) <> "Frame" Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Control '" & DP_SETTINGS_PANEL_NAME & "' exists but is not an MSForms.Frame"
        End If
    'Use the resolved settings panel frame
        Set Fra_Settings = ExistingControl

'------------------------------------------------------------------------------
' RESOLVE SETTINGS MULTIPAGE
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the settings MultiPage
        On Error Resume Next
    'Retrieve the settings MultiPage from the settings frame
        Set ExistingControl = Fra_Settings.Controls(DP_SETTINGS_MULTIPAGE_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing settings MultiPage
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Unable to resolve expected settings MultiPage " & DP_SETTINGS_MULTIPAGE_NAME
        End If
    'Reject a name collision with a non-MultiPage control
        If VBA.TypeName(ExistingControl) <> "MultiPage" Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Control '" & DP_SETTINGS_MULTIPAGE_NAME & "' exists but is not an MSForms.MultiPage"
        End If
    'Use the resolved settings MultiPage
        Set Mp_Settings = ExistingControl

'------------------------------------------------------------------------------
' RESET TRANSIENT HOVER STATE
'------------------------------------------------------------------------------
    'Clear active picker hover before showing settings
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Clear active day-cell hover before showing settings
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Clear active header-label hover before showing settings
        If VBA.Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Clear active footer hover before showing settings
        If VBA.Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset
    'Clear active settings-panel hover before showing settings
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then UF_SettingsPanel_HoverReset

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Suppress lookup errors when the picker panel is not available
        On Error Resume Next
    'Retrieve the picker panel
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Hide the picker panel when available and valid
        If Not ExistingControl Is Nothing Then
            'Hide only when the resolved control is the expected frame
                If VBA.TypeName(ExistingControl) = "Frame" Then
                    'Use the picker panel reference
                        Set Fra_PickerPanel = ExistingControl
                    'Hide the picker panel before showing settings
                        Fra_PickerPanel.Visible = False
                End If
        End If
    'Clear picker-panel mode
        mPickerPanelMode = 0
    'Clear picker-panel hover state
        mHoveredPickerItemIndex = 0

'------------------------------------------------------------------------------
' REFRESH SETTINGS CONTENT
'------------------------------------------------------------------------------
    'Refresh the displayed settings values
        UF_SettingsPanel_RefreshCaptions
    'Refresh the fake settings tab visual state
        UF_SettingsTabs_RefreshVisualState Mp_Settings.Value

'------------------------------------------------------------------------------
' SHOW SETTINGS PANEL
'------------------------------------------------------------------------------
    'Show the settings panel and move it in front of the day grid
        Fra_Settings.Visible = True: Fra_Settings.ZOrder 0

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
        Err.Raise Err.Number, PROC_NAME, "Settings panel show failed: " & Err.Description

End Sub

Private Sub UF_SettingsPanel_Save()

'
'------------------------------------------------------------------------------
'                           SAVE SETTINGS PANEL
'------------------------------------------------------------------------------
' PURPOSE
'   Persists the current values displayed in the in-form settings panel
'
' WHY THIS EXISTS
'   The settings panel exposes an explicit Save icon. Runtime-created MSForms
'   controls should not mutate project-wide DatePicker settings until the user
'   confirms the change
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads Display, Behavior, and Integration settings from the runtime MultiPage,
'   updates the canonical DatePicker settings state, prevents a dead access
'   configuration, persists the settings once, and applies immediate side
'   effects that are safe to apply without reopening the picker
'
' ERROR POLICY
'   Raises a descriptive runtime error if controls cannot be resolved, if the
'   first-day selection is invalid, or if settings cannot be persisted
'
'   If persistence fails after in-memory settings were changed, the previous
'   in-memory settings are restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_Save
'   M_ContextMenu_Update
'   M_KeyboardShortcut_Update
'   M_GridIcon_Remove
'   M_Timer_ApplyClockMode
'   M_Platform_ShouldUseWinAPI
'   M_Window_RemoveTitleBar
'   UF_Form_Format
'   UF_DP_RefreshSettings
'   UF_SettingsPanel_RefreshCaptions
'
' NOTES
'   Enabling WinAPI styling attempts title-bar removal immediately when supported
'
'   Disabling WinAPI styling is persisted immediately, but an already removed
'   title bar can only be restored by reopening the UserForm because the current
'   window helper intentionally exposes removal only
'
'   If right-click menu and in-grid icon are both disabled, the keyboard shortcut
'   fallback is kept enabled to avoid a configuration with no visible/manual
'   DatePicker entry point
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_SettingsPanel_Save"

    Dim ExistingControl         As MSForms.control              'Existing control resolved by name
    Dim Fra_Settings            As MSForms.Frame                'Reusable settings panel
    Dim Mp_Settings             As MSForms.MultiPage            'Settings MultiPage
    Dim Pge_Display             As MSForms.Page                 'Display settings page
    Dim Pge_Behavior            As MSForms.Page                 'Behavior settings page
    Dim Pge_Integration         As MSForms.Page                 'Integration settings page

    Dim Cbo_FirstDay            As MSForms.ComboBox             'First-day ComboBox

    Dim Chk_LocalNames          As MSForms.CheckBox             'Use-local-names checkbox
    Dim Chk_LiveClock           As MSForms.CheckBox             'Live-clock checkbox
    Dim Chk_Compact             As MSForms.CheckBox             'Compact-layout checkbox
    Dim Chk_Weekends            As MSForms.CheckBox             'Highlight-weekends checkbox

    Dim Chk_AllowOutside        As MSForms.CheckBox             'Allow outside-month selection checkbox
    Dim Chk_CloseAfter          As MSForms.CheckBox             'Close-after-selection checkbox

    Dim Chk_RightClick          As MSForms.CheckBox             'Right-click menu checkbox
    Dim Chk_InGridIcon          As MSForms.CheckBox             'In-grid icon checkbox
    Dim Chk_WinAPIStyle         As MSForms.CheckBox             'WinAPI styling checkbox

    Dim NewFirstDay             As Long                         'New first-day setting
    Dim NewUseLocalNames        As Boolean                      'New local-name setting
    Dim NewClockMode            As DP_ClockMode                 'New clock-mode setting
    Dim NewSizeMode             As DP_SizeMode                  'New size-mode setting
    Dim NewHighlightWeekends    As Boolean                      'New weekend-highlight setting
    Dim NewAllowOutside         As Boolean                      'New outside-month setting
    Dim NewCloseAfter           As Boolean                      'New close-after-selection setting
    Dim NewShowRightClick       As Boolean                      'New right-click setting
    Dim NewShowGridIcon         As Boolean                      'New grid-icon setting
    Dim NewUseWinAPI            As Boolean                      'New WinAPI setting
    Dim NewEnableKeyboard       As Boolean                      'New keyboard shortcut fallback setting

    Dim OldFirstDay             As Long                         'Previous first-day setting
    Dim OldUseLocalNames        As Boolean                      'Previous local-name setting
    Dim OldClockMode            As DP_ClockMode                 'Previous clock-mode setting
    Dim OldSizeMode             As DP_SizeMode                  'Previous size-mode setting
    Dim OldHighlightWeekends    As Boolean                      'Previous weekend-highlight setting
    Dim OldAllowOutside         As Boolean                      'Previous outside-month setting
    Dim OldCloseAfter           As Boolean                      'Previous close-after-selection setting
    Dim OldShowRightClick       As Boolean                      'Previous right-click setting
    Dim OldShowGridIcon         As Boolean                      'Previous grid-icon setting
    Dim OldUseWinAPI            As Boolean                      'Previous WinAPI setting
    Dim OldEnableKeyboard       As Boolean                      'Previous keyboard shortcut fallback setting

    Dim FirstDayChanged         As Boolean                      'True when first-day setting changed
    Dim LocalNamesChanged       As Boolean                      'True when local-name setting changed
    Dim ClockModeChanged        As Boolean                      'True when clock mode changed
    Dim SizeModeChanged         As Boolean                      'True when size mode changed
    Dim WeekendChanged          As Boolean                      'True when weekend highlight setting changed
    Dim AllowOutsideChanged     As Boolean                      'True when outside-month selection changed
    Dim RightClickChanged       As Boolean                      'True when right-click integration changed
    Dim GridIconChanged         As Boolean                      'True when in-grid icon integration changed
    Dim WinAPIChanged           As Boolean                      'True when WinAPI styling setting changed
    Dim KeyboardChanged         As Boolean                      'True when keyboard shortcut setting changed

    Dim SettingsMutated         As Boolean                      'True after global settings are changed
    Dim SettingsPersisted       As Boolean                      'True after settings are saved successfully
    Dim ErrorNumber             As Long                         'Captured error number
    Dim ErrorDescription        As String                       'Captured error description
    Dim HandlerStep             As String                       'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve settings panel"

    'Suppress lookup errors while resolving the settings panel
        On Error Resume Next
    'Retrieve the settings panel control
        Set ExistingControl = Me.Controls(DP_SETTINGS_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing settings panel
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Unable to resolve expected settings panel " & DP_SETTINGS_PANEL_NAME
        End If
    'Reject a name collision with a non-frame control
        If VBA.TypeName(ExistingControl) <> "Frame" Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Control '" & DP_SETTINGS_PANEL_NAME & "' exists but is not an MSForms.Frame"
        End If
    'Use the resolved settings panel frame
        Set Fra_Settings = ExistingControl

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS MULTIPAGE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve settings MultiPage"

    'Suppress lookup errors while resolving the settings MultiPage
        On Error Resume Next
    'Retrieve the settings MultiPage control
        Set ExistingControl = Fra_Settings.Controls(DP_SETTINGS_MULTIPAGE_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing settings MultiPage
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Unable to resolve expected settings MultiPage " & DP_SETTINGS_MULTIPAGE_NAME
        End If
    'Reject a name collision with a non-MultiPage control
        If VBA.TypeName(ExistingControl) <> "MultiPage" Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Control '" & DP_SETTINGS_MULTIPAGE_NAME & "' exists but is not an MSForms.MultiPage"
        End If
    
    'Use the resolved settings MultiPage
        Set Mp_Settings = ExistingControl
    'Reject an incomplete settings MultiPage
        If Mp_Settings.Pages.Count < 3 Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Settings MultiPage must contain at least three pages"
        End If

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS PAGES
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve settings pages"
    'Retrieve settings pages
        Set Pge_Display = Mp_Settings.Pages(0)
        Set Pge_Behavior = Mp_Settings.Pages(1)
        Set Pge_Integration = Mp_Settings.Pages(2)

'------------------------------------------------------------------------------
' RETRIEVE DISPLAY CONTROLS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve Display Settings controls"
    'Retrieve the first-day ComboBox
        Set Cbo_FirstDay = Pge_Display.Controls("Cbo_SettingsFirstDay")
    'Retrieve the local-names checkbox
        Set Chk_LocalNames = Pge_Display.Controls("Chk_SettingsUseLocalNames")
    'Retrieve the live-clock checkbox
        Set Chk_LiveClock = Pge_Display.Controls("Chk_SettingsLiveClock")
    'Retrieve the compact-layout checkbox
        Set Chk_Compact = Pge_Display.Controls("Chk_SettingsCompactLayout")
    'Retrieve the weekend-highlight checkbox
        Set Chk_Weekends = Pge_Display.Controls("Chk_SettingsHighlightWeekends")

'------------------------------------------------------------------------------
' RETRIEVE BEHAVIOR CONTROLS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve Behavior Settings controls"
    'Retrieve the allow-outside-month checkbox
        Set Chk_AllowOutside = Pge_Behavior.Controls("Chk_SettingsAllowOutsideMonth")
    'Retrieve the close-after-selection checkbox
        Set Chk_CloseAfter = Pge_Behavior.Controls("Chk_SettingsCloseAfterSelection")

'------------------------------------------------------------------------------
' RETRIEVE INTEGRATION CONTROLS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve Integration Settings controls"
    'Retrieve the right-click menu checkbox
        Set Chk_RightClick = Pge_Integration.Controls("Chk_SettingsRightClickMenu")
    'Retrieve the in-grid icon checkbox
        Set Chk_InGridIcon = Pge_Integration.Controls("Chk_SettingsInGridIcon")
    'Retrieve the WinAPI styling checkbox
        Set Chk_WinAPIStyle = Pge_Integration.Controls("Chk_SettingsWinAPIStyling")

'------------------------------------------------------------------------------
' RESOLVE FIRST-DAY SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve first-day setting"

    'Resolve the first-day setting from the selected ComboBox item
        Select Case Cbo_FirstDay.ListIndex

            Case 0
                'Resolve Sunday from the first ComboBox item
                    NewFirstDay = vbSunday

            Case 1
                'Resolve Monday from the second ComboBox item
                    NewFirstDay = vbMonday

            Case Else
                'Resolve the first-day setting from the ComboBox value fallback
                    Select Case VBA.UCase$(VBA.Trim$(VBA.CStr(Cbo_FirstDay.Value)))
                        Case "SUNDAY"
                            'Resolve Sunday from text value
                                NewFirstDay = vbSunday

                        Case "MONDAY"
                            'Resolve Monday from text value
                                NewFirstDay = vbMonday

                        Case Else
                            'Reject unsupported first-day values
                                Err.Raise vbObjectError + 518, PROC_NAME, _
                                    "First-day setting must be Sunday or Monday"
                    End Select
        
        End Select

    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(NewFirstDay) Then
            Err.Raise vbObjectError + 519, PROC_NAME, _
                "NewFirstDay must be vbSunday or vbMonday"
        End If

'------------------------------------------------------------------------------
' RESOLVE DISPLAY SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve Display Settings values"

    'Resolve the local-name setting
        NewUseLocalNames = CBool(Chk_LocalNames.Value = True)
    'Resolve the clock-mode setting
        If CBool(Chk_LiveClock.Value = True) Then
            NewClockMode = DP_ClockMode_Live
        Else
            NewClockMode = DP_ClockMode_Static
        End If
    'Resolve the size-mode setting
        If CBool(Chk_Compact.Value = True) Then
            NewSizeMode = DP_SizeMode_Compact
        Else
            NewSizeMode = DP_SizeMode_Normal
        End If
    'Resolve the weekend-highlight setting
        NewHighlightWeekends = CBool(Chk_Weekends.Value = True)

'------------------------------------------------------------------------------
' RESOLVE BEHAVIOR SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve Behavior Settings values"
    'Resolve the outside-month setting
        NewAllowOutside = CBool(Chk_AllowOutside.Value = True)
    'Resolve the close-after-selection setting
        NewCloseAfter = CBool(Chk_CloseAfter.Value = True)

'------------------------------------------------------------------------------
' RESOLVE INTEGRATION SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve Integration Settings values"

    'Resolve the right-click setting
        NewShowRightClick = CBool(Chk_RightClick.Value = True)
    'Resolve the grid-icon setting
        NewShowGridIcon = CBool(Chk_InGridIcon.Value = True)
    'Resolve the WinAPI styling setting
        NewUseWinAPI = CBool(Chk_WinAPIStyle.Value = True)
    'Resolve the keyboard shortcut setting through the shared save-resolution
    'seam. Zero built-in entry paths is a valid configuration, so the panel
    'preserves an explicitly disabled shortcut instead of forcing it back on
        NewEnableKeyboard = M_Settings_ResolveKeyboardShortcutOnSave( _
            gDP_EnableKeyboardShortcut, NewShowRightClick, NewShowGridIcon)

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture current settings"

    'Capture the current first-day setting
        OldFirstDay = gDP_FirstDayOfWeek
    'Capture the current local-name setting
        OldUseLocalNames = gDP_UseLocalNames
    'Capture the current clock-mode setting
        OldClockMode = gDP_ClockMode
    'Capture the current size-mode setting
        OldSizeMode = gDP_SizeMode
    'Capture the current weekend-highlight setting
        OldHighlightWeekends = gDP_HighlightWeekends
    'Capture the current outside-month setting
        OldAllowOutside = gDP_AllowOutsideMonthSelection
    'Capture the current close-after-selection setting
        OldCloseAfter = gDP_CloseAfterSelection
    'Capture the current right-click setting
        OldShowRightClick = gDP_ShowRightClick
    'Capture the current grid-icon setting
        OldShowGridIcon = gDP_ShowGridIcon
    'Capture the current WinAPI styling setting
        OldUseWinAPI = gDP_UseWinAPI
    'Capture the current keyboard shortcut fallback setting
        OldEnableKeyboard = gDP_EnableKeyboardShortcut

'------------------------------------------------------------------------------
' RESOLVE CHANGE FLAGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve change flags"
    'Resolve whether first-day setting changed
        FirstDayChanged = (NewFirstDay <> OldFirstDay)
    'Resolve whether local-name setting changed
        LocalNamesChanged = (NewUseLocalNames <> OldUseLocalNames)
    'Resolve whether clock mode changed
        ClockModeChanged = (NewClockMode <> OldClockMode)
    'Resolve whether size mode changed
        SizeModeChanged = (NewSizeMode <> OldSizeMode)
    'Resolve whether weekend-highlight setting changed
        WeekendChanged = (NewHighlightWeekends <> OldHighlightWeekends)
    'Resolve whether outside-month selection changed
        AllowOutsideChanged = (NewAllowOutside <> OldAllowOutside)
    'Resolve whether right-click integration changed
        RightClickChanged = (NewShowRightClick <> OldShowRightClick)
    'Resolve whether grid-icon integration changed
        GridIconChanged = (NewShowGridIcon <> OldShowGridIcon)
    'Resolve whether WinAPI styling changed
        WinAPIChanged = (NewUseWinAPI <> OldUseWinAPI)
    'Resolve whether keyboard shortcut fallback changed
        KeyboardChanged = (NewEnableKeyboard <> OldEnableKeyboard)

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Update in-memory settings"

    'Store the new first-day setting
        gDP_FirstDayOfWeek = NewFirstDay
    'Store the new local-name setting
        gDP_UseLocalNames = NewUseLocalNames
    'Store the new clock-mode setting
        gDP_ClockMode = NewClockMode
    'Store the new size-mode setting
        gDP_SizeMode = NewSizeMode
    'Store the new weekend-highlight setting
        gDP_HighlightWeekends = NewHighlightWeekends
    'Store the new outside-month setting
        gDP_AllowOutsideMonthSelection = NewAllowOutside
    'Store the new close-after-selection setting
        gDP_CloseAfterSelection = NewCloseAfter
    'Store the new right-click setting
        gDP_ShowRightClick = NewShowRightClick
    'Store the new grid-icon setting
        gDP_ShowGridIcon = NewShowGridIcon
    'Store the new WinAPI styling setting
        gDP_UseWinAPI = NewUseWinAPI
    'Store the new keyboard shortcut fallback setting
        gDP_EnableKeyboardShortcut = NewEnableKeyboard

    'Mark settings as mutated
        SettingsMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Persist settings"

    'Persist the updated settings once
        M_Settings_Save
    'Mark settings as persisted
        SettingsPersisted = True

'------------------------------------------------------------------------------
' APPLY LIGHTWEIGHT INTEGRATION SIDE EFFECTS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply lightweight integration side effects"
        
    'Synchronize right-click menu integration only when it changed
        If RightClickChanged Then M_ContextMenu_Update
    'Synchronize keyboard shortcut fallback only when it changed
        If KeyboardChanged Then M_KeyboardShortcut_Update
    'Remove any stale in-grid icon only when the feature was disabled
        If GridIconChanged Then
            If Not gDP_ShowGridIcon Then M_GridIcon_Hide
        End If

'------------------------------------------------------------------------------
' APPLY LIGHTWEIGHT CLOCK SIDE EFFECTS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply lightweight clock side effects"

    'Apply clock mode only when it changed
        If ClockModeChanged Then M_Timer_ApplyClockMode

'------------------------------------------------------------------------------
' UNLOAD AFTER STRUCTURAL SETTINGS CHANGE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Unload after structural settings change"

    'Unload the current UserForm when shell layout or WinAPI styling changed
        If SizeModeChanged Or WinAPIChanged Then
            'Unload the form so the next open rebuilds it with the new shell settings
                Unload Me
            'Exit because the current UserForm instance is terminating
                Exit Sub
        End If

'------------------------------------------------------------------------------
' APPLY NON-STRUCTURAL VISUAL REFRESH
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply non-structural visual refresh"

    'Refresh DatePicker captions and day grid only when needed
        If FirstDayChanged _
        Or LocalNamesChanged _
        Or WeekendChanged _
        Or AllowOutsideChanged Then
            UF_DP_RefreshSettings
        End If

'------------------------------------------------------------------------------
' REFRESH SETTINGS PANEL VALUES
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh settings panel values after save"

    'Refresh the settings panel values after save
        UF_SettingsPanel_RefreshCaptions

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
    'Restore previous in-memory settings when persistence failed after mutation
        If SettingsMutated And Not SettingsPersisted Then

            'Suppress rollback errors
                On Error Resume Next

                gDP_FirstDayOfWeek = OldFirstDay
                gDP_UseLocalNames = OldUseLocalNames
                gDP_ClockMode = OldClockMode
                gDP_SizeMode = OldSizeMode
                gDP_HighlightWeekends = OldHighlightWeekends
                gDP_AllowOutsideMonthSelection = OldAllowOutside
                gDP_CloseAfterSelection = OldCloseAfter
                gDP_ShowRightClick = OldShowRightClick
                gDP_ShowGridIcon = OldShowGridIcon
                gDP_UseWinAPI = OldUseWinAPI
                gDP_EnableKeyboardShortcut = OldEnableKeyboard

                On Error GoTo 0

        End If

    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Settings panel save failed: " & ErrorDescription

End Sub

Private Sub UF_SettingsPanel_Hide()

'
'------------------------------------------------------------------------------
'                           HIDE SETTINGS PANEL
'------------------------------------------------------------------------------
' PURPOSE
'   Hides the reusable settings panel
'
' WHY THIS EXISTS
'   Settings is an overlay panel and must be removable without rebuilding the
'   DatePicker grid or changing persisted settings
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resets settings-panel hover state, resolves Fra_Settings when available, and
'   hides it safely
'
' ERROR POLICY
'   Best-effort UI cleanup. Missing or stale settings-panel controls are treated
'   as safe no-ops
'
' DEPENDENCIES
'   UF_SettingsPanel_HoverReset
'   DP_SETTINGS_PANEL_NAME
'
' NOTES
'   This routine does not clear saved settings
'
'   This routine does not rebuild the settings panel
'
'   This routine intentionally does not raise outward because hiding an overlay
'   panel should never interrupt DatePicker interaction
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ExistingControl     As MSForms.control      'Control resolved by settings-panel name
    Dim Fra_Settings        As MSForms.Frame        'Reusable settings panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress hide errors
        On Error Resume Next

'------------------------------------------------------------------------------
' RESET HOVER STATE
'------------------------------------------------------------------------------
    'Reset settings-panel hover before hiding the panel
        UF_SettingsPanel_HoverReset
    'Clear the settings-panel hover tracker defensively
        mHoveredSettingsPanelLabelName = vbNullString

'------------------------------------------------------------------------------
' RESOLVE PANEL
'------------------------------------------------------------------------------
    'Retrieve the settings panel control when available
        Set ExistingControl = Me.Controls(DP_SETTINGS_PANEL_NAME)
    'Use the resolved control only when it is the expected frame
        If Not ExistingControl Is Nothing Then
            If VBA.TypeName(ExistingControl) = "Frame" Then
                Set Fra_Settings = ExistingControl
            End If
        End If

'------------------------------------------------------------------------------
' HIDE PANEL
'------------------------------------------------------------------------------
    'Hide the settings panel when available
        If Not Fra_Settings Is Nothing Then
            Fra_Settings.Visible = False
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:
    'Clear suppressed hide errors
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Sub UF_SettingsPanel_HoverApply(ByVal LabelName As String)

'
'------------------------------------------------------------------------------
'                       APPLY SETTINGS-PANEL HOVER
'------------------------------------------------------------------------------
' PURPOSE
'   Applies hover formatting to the settings-panel header action labels
'
' WHY THIS EXISTS
'   Lbl_SettingsSave and Lbl_SettingsClose are runtime-created labels hosted
'   inside Fra_Settings. They therefore do not receive normal UserForm label
'   event procedures and must be routed through cDatePickerLabelHook
'
' INPUTS
'   LabelName
'     Name of the settings-panel label currently under the mouse pointer
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves whether the hovered label is a supported settings header action,
'   resets any previous settings-panel hover state, applies hover formatting to
'   Save or Close when applicable, and stores the currently hovered label name
'
' ERROR POLICY
'   Raises a descriptive runtime error if a supported settings header label
'   cannot be resolved or formatted
'
'   Unsupported settings-panel labels are treated as neutral hover-reset areas
'
' DEPENDENCIES
'   DP_SETTINGS_PANEL_NAME
'   UF_SettingsPanel_HoverReset
'   Lbl_SettingsSave
'   Lbl_SettingsClose
'
' NOTES
'   Tab, title, page, and body controls may route here only to reset Save / Close
'   hover state. They are intentionally not restyled by this routine
'
'   Label matching is normalized so behavior is independent from Option Compare
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_SettingsPanel_HoverApply"

    Dim RawLabelName            As String               'Trimmed label name supplied by caller
    Dim NormalizedLabelName     As String               'Uppercase label name used for routing
    Dim TargetLabelName         As String               'Canonical target label name
    Dim ExistingControl         As MSForms.control      'Existing control resolved by name
    Dim Fra_Settings            As MSForms.Frame        'Reusable settings panel
    Dim Lbl_Target              As MSForms.Label        'Hovered settings label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the supplied label name
        RawLabelName = VBA.Trim$(LabelName)
    'Reject an empty label name
        If VBA.Len(RawLabelName) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "LabelName cannot be empty"
        End If
    'Normalize the label name for Option Compare independent routing
        NormalizedLabelName = VBA.UCase$(RawLabelName)

'------------------------------------------------------------------------------
' RESOLVE TARGET LABEL
'------------------------------------------------------------------------------
    'Resolve supported settings-panel header labels
        Select Case NormalizedLabelName

            Case "LBL_SETTINGSSAVE"
                'Store the canonical Save label name
                    TargetLabelName = "Lbl_SettingsSave"

            Case "LBL_SETTINGSCLOSE"
                'Store the canonical Close label name
                    TargetLabelName = "Lbl_SettingsClose"

            Case Else
                'Remove previous settings hover state when moving over neutral settings surfaces
                    If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then
                        UF_SettingsPanel_HoverReset
                    End If
                'Exit because neutral settings labels do not receive hover styling
                    Exit Sub

        End Select

'------------------------------------------------------------------------------
' EXIT IF SAME LABEL IS ALREADY HOVERED
'------------------------------------------------------------------------------
    'Exit when the same settings label is already highlighted
        If VBA.StrComp(mHoveredSettingsPanelLabelName, TargetLabelName, vbBinaryCompare) = 0 Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RESET PREVIOUS HOVER STATE
'------------------------------------------------------------------------------
    'Remove hover formatting from the previously highlighted settings label
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then
            UF_SettingsPanel_HoverReset
        End If

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the settings panel
        On Error Resume Next
    'Retrieve the settings panel control
        Set ExistingControl = Me.Controls(DP_SETTINGS_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing settings panel
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Unable to resolve expected settings panel " & DP_SETTINGS_PANEL_NAME
        End If
    'Reject a name collision with a non-frame control
        If VBA.TypeName(ExistingControl) <> "Frame" Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Control '" & DP_SETTINGS_PANEL_NAME & "' exists but is not an MSForms.Frame"
        End If
    'Use the resolved settings panel frame
        Set Fra_Settings = ExistingControl

'------------------------------------------------------------------------------
' RETRIEVE TARGET LABEL
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the target label
        On Error Resume Next
    'Retrieve the target settings label
        Set ExistingControl = Fra_Settings.Controls(TargetLabelName)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing target label
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Unable to resolve expected settings label " & TargetLabelName
        End If
    'Reject a name collision with a non-label control
        If VBA.TypeName(ExistingControl) <> "Label" Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Control '" & TargetLabelName & "' exists but is not an MSForms.Label"
        End If
    'Use the resolved settings label
        Set Lbl_Target = ExistingControl

'------------------------------------------------------------------------------
' APPLY HOVER STATE
'------------------------------------------------------------------------------
    'Apply hover formatting to the target settings header label
        With Lbl_Target
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_SETTINGS_HEADER_ICON_HOVER_BACK_COLOR
            .ForeColor = DP_SETTINGS_HEADER_ICON_HOVER_FORE_COLOR
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = DP_SETTINGS_HEADER_ICON_HOVER_BORDER_COLOR
            .SpecialEffect = fmSpecialEffectFlat
        End With
    'Move the hovered label to the front
        Lbl_Target.ZOrder 0

'------------------------------------------------------------------------------
' STORE HOVER STATE
'------------------------------------------------------------------------------
    'Store the currently hovered settings-panel label
        mHoveredSettingsPanelLabelName = TargetLabelName

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
        Err.Raise Err.Number, PROC_NAME, "Settings-panel hover application failed: " & Err.Description

End Sub

Private Sub UF_SettingsPanel_HoverReset()

'
'------------------------------------------------------------------------------
'                       RESET SETTINGS-PANEL HOVER
'------------------------------------------------------------------------------
' PURPOSE
'   Removes hover formatting from the active settings-panel header action label
'
' WHY THIS EXISTS
'   The Save and Close header icons should return to their normal visual state
'   when the mouse leaves them or moves to a neutral area of the settings panel
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the currently hovered settings-panel label, restores its normal
'   transparent icon formatting, and clears the stored hover state
'
' ERROR POLICY
'   Best-effort UI cleanup. Suppresses reset errors because hover reset should
'   never interrupt UserForm interaction
'
' DEPENDENCIES
'   DP_SETTINGS_PANEL_NAME
'   Fra_Settings
'   Lbl_SettingsSave
'   Lbl_SettingsClose
'
' NOTES
'   The routine restores only properties modified by
'   UF_SettingsPanel_HoverApply
'
'   Stale or missing labels are ignored because controls may have been rebuilt
'   while a hover state was active
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentLabelName    As String           'Current hovered settings label name
    Dim Fra_Settings        As MSForms.Frame    'Reusable settings panel
    Dim Lbl_Target          As MSForms.Label    'Settings label to reset

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress hover-reset errors
        On Error Resume Next
    'Capture the current hovered settings label
        CurrentLabelName = VBA.Trim$(mHoveredSettingsPanelLabelName)

'------------------------------------------------------------------------------
' EXIT IF NOTHING IS HOVERED
'------------------------------------------------------------------------------
    'Exit if no settings-panel label is currently highlighted
        If VBA.Len(CurrentLabelName) = 0 Then GoTo Clean_Exit

'------------------------------------------------------------------------------
' RETRIEVE TARGET LABEL
'------------------------------------------------------------------------------
    'Retrieve the settings panel
        Set Fra_Settings = Me.Controls(DP_SETTINGS_PANEL_NAME)
    'Retrieve the settings label to reset
        Set Lbl_Target = Fra_Settings.Controls(CurrentLabelName)

'------------------------------------------------------------------------------
' RESTORE NORMAL STATE
'------------------------------------------------------------------------------
    'Restore normal settings header icon formatting when available
        If Not Lbl_Target Is Nothing Then
            With Lbl_Target
                .BackStyle = fmBackStyleTransparent
                .BackColor = DP_SETTINGS_PANEL_BACK_COLOR
                .ForeColor = DP_SETTINGS_HEADER_ICON_FORE_COLOR
                .BorderStyle = fmBorderStyleNone
                .BorderColor = DP_SETTINGS_PANEL_BACK_COLOR
                .SpecialEffect = fmSpecialEffectFlat
            End With
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:
    'Clear the current settings-panel hover label
        mHoveredSettingsPanelLabelName = vbNullString
    'Clear suppressed reset errors
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub UF_SettingsPanel_RefreshCaptions()

'
'------------------------------------------------------------------------------
'                       REFRESH SETTINGS PANEL VALUES
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes the values shown inside Fra_Settings
'
' WHY THIS EXISTS
'   The settings panel should reflect the current in-memory DatePicker settings
'   whenever it is shown
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Updates the Display, Behavior, and Integration settings pages from current
'   global DatePicker settings
'
' ERROR POLICY
'   Raises a descriptive runtime error if the settings panel, MultiPage, page, or
'   settings controls cannot be resolved or updated
'
' DEPENDENCIES
'   gDP_FirstDayOfWeek
'   gDP_UseLocalNames
'   gDP_ClockMode
'   DP_ClockMode_Live
'   gDP_SizeMode
'   DP_SizeMode_Compact
'   gDP_HighlightWeekends
'   gDP_AllowOutsideMonthSelection
'   gDP_CloseAfterSelection
'   gDP_ShowRightClick
'   gDP_ShowGridIcon
'   gDP_UseWinAPI
'   M_Settings_IsValidFirstDayOfWeek
'
' NOTES
'   This routine displays current values only
'
'   Persistence is intentionally routed through the explicit Save action
'
'   Developer-owned settings layout constants are intentionally not validated
'   here. They should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_SettingsPanel_RefreshCaptions"

    Dim ExistingControl         As MSForms.control          'Control resolved by name
    Dim Fra_Settings            As MSForms.Frame            'Reusable settings panel
    Dim Mp_Settings             As MSForms.MultiPage        'Settings MultiPage
    
    Dim Pge_Display             As MSForms.Page             'Display settings page
    Dim Pge_Behavior            As MSForms.Page             'Behavior settings page
    Dim Pge_Integration         As MSForms.Page             'Integration settings page
    
    Dim Cbo_FirstDay            As MSForms.ComboBox         'First-day ComboBox
    Dim Chk_LocalNames          As MSForms.CheckBox         'Use-local-names checkbox
    Dim Chk_LiveClock           As MSForms.CheckBox         'Live-clock checkbox
    Dim Chk_Compact             As MSForms.CheckBox         'Compact-layout checkbox
    Dim Chk_Weekends            As MSForms.CheckBox         'Highlight-weekends checkbox
    Dim Chk_AllowOutside        As MSForms.CheckBox         'Allow outside-month selection checkbox
    Dim Chk_CloseAfter          As MSForms.CheckBox         'Close-after-selection checkbox
    Dim Chk_RightClick          As MSForms.CheckBox         'Right-click menu checkbox
    Dim Chk_InGridIcon          As MSForms.CheckBox         'In-grid icon checkbox
    Dim Chk_WinAPIStyle         As MSForms.CheckBox         'WinAPI styling checkbox
    
    Dim HandlerStep             As String                   'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' VALIDATE SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate first-day setting"

    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(gDP_FirstDayOfWeek) Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "gDP_FirstDayOfWeek must be vbSunday or vbMonday"
        End If

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve settings panel"

    'Suppress lookup errors while resolving the settings panel
        On Error Resume Next
    'Retrieve the settings panel control
        Set ExistingControl = Me.Controls(DP_SETTINGS_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing settings panel
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Unable to resolve expected settings panel " & DP_SETTINGS_PANEL_NAME
        End If
    'Reject a name collision with a non-frame control
        If VBA.TypeName(ExistingControl) <> "Frame" Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Control '" & DP_SETTINGS_PANEL_NAME & "' exists but is not an MSForms.Frame"
        End If
    'Use the resolved settings panel frame
        Set Fra_Settings = ExistingControl

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS MULTIPAGE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve settings MultiPage"

    'Suppress lookup errors while resolving the settings MultiPage
        On Error Resume Next
    'Retrieve the settings MultiPage control
        Set ExistingControl = Fra_Settings.Controls(DP_SETTINGS_MULTIPAGE_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing settings MultiPage
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Unable to resolve expected settings MultiPage " & DP_SETTINGS_MULTIPAGE_NAME
        End If
    'Reject a name collision with a non-MultiPage control
        If VBA.TypeName(ExistingControl) <> "MultiPage" Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Control '" & DP_SETTINGS_MULTIPAGE_NAME & "' exists but is not an MSForms.MultiPage"
        End If
    'Use the resolved settings MultiPage
        Set Mp_Settings = ExistingControl
    'Reject an incomplete settings MultiPage
        If Mp_Settings.Pages.Count < 3 Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "Settings MultiPage must contain at least three pages"
        End If

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS PAGES
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve settings pages"

    'Retrieve the settings pages
        Set Pge_Display = Mp_Settings.Pages(0)
        Set Pge_Behavior = Mp_Settings.Pages(1)
        Set Pge_Integration = Mp_Settings.Pages(2)

'------------------------------------------------------------------------------
' RETRIEVE DISPLAY CONTROLS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve Display Settings controls"

    'Retrieve display controls
        Set Cbo_FirstDay = Pge_Display.Controls("Cbo_SettingsFirstDay")
        Set Chk_LocalNames = Pge_Display.Controls("Chk_SettingsUseLocalNames")
        Set Chk_LiveClock = Pge_Display.Controls("Chk_SettingsLiveClock")
        Set Chk_Compact = Pge_Display.Controls("Chk_SettingsCompactLayout")
        Set Chk_Weekends = Pge_Display.Controls("Chk_SettingsHighlightWeekends")

'------------------------------------------------------------------------------
' RETRIEVE BEHAVIOR CONTROLS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve Behavior Settings controls"

    'Retrieve behavior controls
        Set Chk_AllowOutside = Pge_Behavior.Controls("Chk_SettingsAllowOutsideMonth")
        Set Chk_CloseAfter = Pge_Behavior.Controls("Chk_SettingsCloseAfterSelection")

'------------------------------------------------------------------------------
' RETRIEVE INTEGRATION CONTROLS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve Integration Settings controls"

    'Retrieve integration controls
        Set Chk_RightClick = Pge_Integration.Controls("Chk_SettingsRightClickMenu")
        Set Chk_InGridIcon = Pge_Integration.Controls("Chk_SettingsInGridIcon")
        Set Chk_WinAPIStyle = Pge_Integration.Controls("Chk_SettingsWinAPIStyling")

'------------------------------------------------------------------------------
' REFRESH FIRST-DAY COMBOBOX
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh first-day ComboBox"

    'Clear the current ComboBox values
        Cbo_FirstDay.Clear
    'Add the Sunday option
        Cbo_FirstDay.AddItem "Sunday"
    'Add the Monday option
        Cbo_FirstDay.AddItem "Monday"
    'Select Monday when it is the active first-day setting
        If gDP_FirstDayOfWeek = vbMonday Then
            Cbo_FirstDay.ListIndex = 1
        Else
            Cbo_FirstDay.ListIndex = 0
        End If

'------------------------------------------------------------------------------
' REFRESH DISPLAY CHECKBOX VALUES
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh Display Settings values"

    'Refresh the local-names checkbox value
        Chk_LocalNames.Value = CBool(gDP_UseLocalNames)
    'Refresh the live-clock checkbox value
        Chk_LiveClock.Value = (gDP_ClockMode = DP_ClockMode_Live)
    'Refresh the compact-layout checkbox value
        Chk_Compact.Value = (gDP_SizeMode = DP_SizeMode_Compact)
    'Refresh the weekend-highlight checkbox value
        Chk_Weekends.Value = CBool(gDP_HighlightWeekends)

'------------------------------------------------------------------------------
' REFRESH BEHAVIOR CHECKBOX VALUES
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh Behavior Settings values"

    'Refresh the allow-outside-month checkbox value
        Chk_AllowOutside.Value = CBool(gDP_AllowOutsideMonthSelection)
    'Refresh the close-after-selection checkbox value
        Chk_CloseAfter.Value = CBool(gDP_CloseAfterSelection)

'------------------------------------------------------------------------------
' REFRESH INTEGRATION CHECKBOX VALUES
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh Integration Settings values"

    'Refresh the right-click menu checkbox value
        Chk_RightClick.Value = CBool(gDP_ShowRightClick)
    'Refresh the in-grid icon checkbox value
        Chk_InGridIcon.Value = CBool(gDP_ShowGridIcon)
    'Refresh the WinAPI styling checkbox value
        Chk_WinAPIStyle.Value = CBool(gDP_UseWinAPI)

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
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Settings panel value refresh failed: " & Err.Description

End Sub


'
'------------------------------------------------------------------------------
'
'                           SETTINGS TABS AND PAGE SUPPORT
'
'------------------------------------------------------------------------------
'

Private Sub UF_SettingsTabs_Build( _
    ByVal ParentFrame As MSForms.Frame, _
    ByVal BodyFont As Object)

'
'------------------------------------------------------------------------------
'                           BUILD SETTINGS TABS
'------------------------------------------------------------------------------
' PURPOSE
'   Creates the label-based tab strip used by the settings panel
'
' WHY THIS EXISTS
'   Native MSForms.MultiPage tabs are rendered by the host environment and cannot
'   be styled reliably. A label-based tab strip gives the DatePicker full visual
'   control while the hidden MultiPage remains responsible for page switching
'
' INPUTS
'   ParentFrame
'     Settings frame that owns the fake tab labels
'
'   BodyFont
'     Clean font object used for the tab captions
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses the Display, Behavior, and Integration tab labels,
'   formats them, registers their click / hover hooks, applies the initial tab
'   visual state, and restores tab layering
'
' ERROR POLICY
'   Raises a descriptive runtime error if the parent frame, font, labels, or
'   hooks cannot be resolved or created
'
' DEPENDENCIES
'   UF_Ensure_FrameLabel
'   UF_SettingsTab_ApplyVisualState
'   cDatePickerLabelHook
'   mSettingsPanelHooks
'
' NOTES
'   The native MultiPage tab strip must remain hidden through fmTabStyleNone
'
'   This routine does not reset mSettingsPanelHooks because it may be called as
'   part of the wider settings-panel build. It initializes the collection only
'   when missing
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_SettingsTabs_Build"

    Dim LabelHook               As cDatePickerLabelHook         'Runtime tab click / hover hook
    Dim Lbl_TabDisplay          As MSForms.Label                'Display tab label
    Dim Lbl_TabBehavior         As MSForms.Label                'Behavior tab label
    Dim Lbl_TabIntegration      As MSForms.Label                'Integration tab label
    Dim TabLeft                 As Single                       'Current tab left position

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Create settings-panel hook storage when this routine is called independently
        If mSettingsPanelHooks Is Nothing Then
            Set mSettingsPanelHooks = New Collection
        End If

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing parent frame
        If ParentFrame Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentFrame cannot be Nothing"
        End If
    'Reject a missing body font
        If BodyFont Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, "BodyFont cannot be Nothing"
        End If

'------------------------------------------------------------------------------
' CREATE / FORMAT DISPLAY TAB
'------------------------------------------------------------------------------
    'Initialize the first tab left position
        TabLeft = DP_SETTINGS_TAB_LEFT

    'Create or retrieve the Display tab label
        Set Lbl_TabDisplay = UF_Ensure_FrameLabel(ParentFrame, DP_SETTINGS_TAB_DISPLAY_NAME)
    'Apply layout and visual properties to the Display tab
        With Lbl_TabDisplay
            .Caption = DP_SETTINGS_TAB_DISPLAY_CAPTION
            .Left = TabLeft
            .Top = DP_SETTINGS_TAB_TOP
            .Width = DP_SETTINGS_TAB_DISPLAY_WIDTH
            .Height = DP_SETTINGS_TAB_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_SETTINGS_TAB_SELECTED_BACK_COLOR
            .ForeColor = DP_SETTINGS_TAB_SELECTED_FORE_COLOR
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = DP_SETTINGS_TAB_BORDER_COLOR
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .MousePointer = fmMousePointerCustom
            .ControlTipText = DP_SETTINGS_PAGE_DISPLAY_CAPTION
        End With
    'Assign the clean body font to the Display tab
        Set Lbl_TabDisplay.Font = BodyFont
    'Remove any previous Display tab hook key when rebuilding independently
        On Error Resume Next
        mSettingsPanelHooks.Remove Lbl_TabDisplay.Name
        Err.Clear
        On Error GoTo ErrorHandler
    'Create the Display tab click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the Display tab to the Display Settings page
        LabelHook.Initialize Me, Lbl_TabDisplay, "SHOW_SETTINGS_DISPLAY", "SETTINGS"
    'Store the hook so that the click / hover events remain alive
        mSettingsPanelHooks.Add LabelHook, Lbl_TabDisplay.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT BEHAVIOR TAB
'------------------------------------------------------------------------------
    'Advance the tab left position
        TabLeft = TabLeft + DP_SETTINGS_TAB_DISPLAY_WIDTH + DP_SETTINGS_TAB_GAP

    'Create or retrieve the Behavior tab label
        Set Lbl_TabBehavior = UF_Ensure_FrameLabel(ParentFrame, DP_SETTINGS_TAB_BEHAVIOR_NAME)
    'Apply layout and visual properties to the Behavior tab
        With Lbl_TabBehavior
            .Caption = DP_SETTINGS_TAB_BEHAVIOR_CAPTION
            .Left = TabLeft
            .Top = DP_SETTINGS_TAB_TOP
            .Width = DP_SETTINGS_TAB_BEHAVIOR_WIDTH
            .Height = DP_SETTINGS_TAB_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_SETTINGS_TAB_NORMAL_BACK_COLOR
            .ForeColor = DP_SETTINGS_TAB_NORMAL_FORE_COLOR
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = DP_SETTINGS_TAB_BORDER_COLOR
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .MousePointer = fmMousePointerCustom
            .ControlTipText = DP_SETTINGS_PAGE_BEHAVIOR_CAPTION
        End With
    'Assign the clean body font to the Behavior tab
        Set Lbl_TabBehavior.Font = BodyFont
    'Remove any previous Behavior tab hook key when rebuilding independently
        On Error Resume Next
        mSettingsPanelHooks.Remove Lbl_TabBehavior.Name
        Err.Clear
        On Error GoTo ErrorHandler
    'Create the Behavior tab click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the Behavior tab to the Behavior Settings page
        LabelHook.Initialize Me, Lbl_TabBehavior, "SHOW_SETTINGS_BEHAVIOR", "SETTINGS"
    'Store the hook so that the click / hover events remain alive
        mSettingsPanelHooks.Add LabelHook, Lbl_TabBehavior.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT INTEGRATION TAB
'------------------------------------------------------------------------------
    'Advance the tab left position
        TabLeft = TabLeft + DP_SETTINGS_TAB_BEHAVIOR_WIDTH + DP_SETTINGS_TAB_GAP

    'Create or retrieve the Integration tab label
        Set Lbl_TabIntegration = UF_Ensure_FrameLabel(ParentFrame, DP_SETTINGS_TAB_INTEGRATION_NAME)
    'Apply layout and visual properties to the Integration tab
        With Lbl_TabIntegration
            .Caption = DP_SETTINGS_TAB_INTEGRATION_CAPTION
            .Left = TabLeft
            .Top = DP_SETTINGS_TAB_TOP
            .Width = DP_SETTINGS_TAB_INTEGRATION_WIDTH
            .Height = DP_SETTINGS_TAB_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_SETTINGS_TAB_NORMAL_BACK_COLOR
            .ForeColor = DP_SETTINGS_TAB_NORMAL_FORE_COLOR
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = DP_SETTINGS_TAB_BORDER_COLOR
            .SpecialEffect = fmSpecialEffectFlat
            .TextAlign = fmTextAlignCenter
            .WordWrap = False
            .Enabled = True
            .Visible = True
            .MousePointer = fmMousePointerCustom
            .ControlTipText = DP_SETTINGS_PAGE_INTEGRATION_CAPTION
        End With
    'Assign the clean body font to the Integration tab
        Set Lbl_TabIntegration.Font = BodyFont
    'Remove any previous Integration tab hook key when rebuilding independently
        On Error Resume Next
        mSettingsPanelHooks.Remove Lbl_TabIntegration.Name
        Err.Clear
        On Error GoTo ErrorHandler
    'Create the Integration tab click / hover hook
        Set LabelHook = New cDatePickerLabelHook
    'Connect the Integration tab to the Integration Settings page
        LabelHook.Initialize Me, Lbl_TabIntegration, "SHOW_SETTINGS_INTEGRATION", "SETTINGS"
    'Store the hook so that the click / hover events remain alive
        mSettingsPanelHooks.Add LabelHook, Lbl_TabIntegration.Name

'------------------------------------------------------------------------------
' APPLY INITIAL VISUAL STATE
'------------------------------------------------------------------------------
    'Apply selected state to the tabs
        UF_SettingsTab_ApplyVisualState Lbl_TabDisplay, True
        UF_SettingsTab_ApplyVisualState Lbl_TabBehavior, False
        UF_SettingsTab_ApplyVisualState Lbl_TabIntegration, False

'------------------------------------------------------------------------------
' RESTORE LAYERING
'------------------------------------------------------------------------------
    'Move the tabs to the front
        Lbl_TabDisplay.ZOrder 0
        Lbl_TabBehavior.ZOrder 0
        Lbl_TabIntegration.ZOrder 0

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
        Err.Raise Err.Number, PROC_NAME, "Settings tab creation failed: " & Err.Description

End Sub

Private Sub UF_SettingsPanel_SelectPage(ByVal PageIndex As Long)

'
'------------------------------------------------------------------------------
'                           SELECT SETTINGS PAGE
'------------------------------------------------------------------------------
' PURPOSE
'   Selects one settings MultiPage page through the label-based tab strip
'
' WHY THIS EXISTS
'   The native MultiPage tabs are hidden. Page navigation is therefore routed
'   through fake tab labels
'
' INPUTS
'   PageIndex
'     Zero-based settings page index
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the requested page, resolves Fra_Settings and Mp_Settings safely,
'   changes the selected MultiPage page, clears settings-header hover state, and
'   refreshes the visual state of the fake tabs
'
' ERROR POLICY
'   Raises a descriptive runtime error if the settings panel, MultiPage, or page
'   index is invalid
'
' DEPENDENCIES
'   UF_SettingsPanel_HoverReset
'   UF_SettingsTabs_RefreshVisualState
'
' NOTES
'   This routine does not rebuild the settings panel
'
'   The settings panel currently supports three pages:
'     - 0 Display Settings
'     - 1 Behavior Settings
'     - 2 Integration Settings
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_SettingsPanel_SelectPage"
    
    Const SETTINGS_PAGE_FIRST   As Long = 0                 'First supported settings page index
    Const SETTINGS_PAGE_LAST    As Long = 2                 'Last supported settings page index

    Dim ExistingControl         As MSForms.control          'Control resolved by name
    Dim Fra_Settings            As MSForms.Frame            'Reusable settings panel
    Dim Mp_Settings             As MSForms.MultiPage        'Settings MultiPage

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject unsupported page indexes
        If PageIndex < SETTINGS_PAGE_FIRST Or PageIndex > SETTINGS_PAGE_LAST Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "PageIndex must be between " & VBA.CStr(SETTINGS_PAGE_FIRST) & _
                " and " & VBA.CStr(SETTINGS_PAGE_LAST)
        End If

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the settings panel
        On Error Resume Next
    'Retrieve the settings panel control
        Set ExistingControl = Me.Controls(DP_SETTINGS_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing settings panel
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Unable to resolve expected settings panel " & DP_SETTINGS_PANEL_NAME
        End If
    'Reject a name collision with a non-frame control
        If VBA.TypeName(ExistingControl) <> "Frame" Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Control '" & DP_SETTINGS_PANEL_NAME & "' exists but is not an MSForms.Frame"
        End If
    'Use the resolved settings panel frame
        Set Fra_Settings = ExistingControl

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS MULTIPAGE
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the settings MultiPage
        On Error Resume Next
    'Retrieve the settings MultiPage control
        Set ExistingControl = Fra_Settings.Controls(DP_SETTINGS_MULTIPAGE_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing settings MultiPage
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Unable to resolve expected settings MultiPage " & DP_SETTINGS_MULTIPAGE_NAME
        End If
    'Reject a name collision with a non-MultiPage control
        If VBA.TypeName(ExistingControl) <> "MultiPage" Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Control '" & DP_SETTINGS_MULTIPAGE_NAME & "' exists but is not an MSForms.MultiPage"
        End If
    'Use the resolved settings MultiPage
        Set Mp_Settings = ExistingControl

'------------------------------------------------------------------------------
' VALIDATE MULTIPAGE STATE
'------------------------------------------------------------------------------
    'Reject page indexes outside the actual MultiPage page count
        If PageIndex > Mp_Settings.Pages.Count - 1 Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "PageIndex exceeds the current MultiPage page count"
        End If

'------------------------------------------------------------------------------
' RESET SETTINGS HOVER STATE
'------------------------------------------------------------------------------
    'Clear Save / Close hover state before switching settings page
        UF_SettingsPanel_HoverReset

'------------------------------------------------------------------------------
' SELECT PAGE
'------------------------------------------------------------------------------
    'Select the requested settings page
        Mp_Settings.Value = PageIndex

'------------------------------------------------------------------------------
' REFRESH TAB VISUAL STATE
'------------------------------------------------------------------------------
    'Refresh the fake settings tab visual state
        UF_SettingsTabs_RefreshVisualState PageIndex

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
        Err.Raise Err.Number, PROC_NAME, "Settings page selection failed: " & Err.Description

End Sub

Private Sub UF_SettingsTabs_RefreshVisualState( _
    Optional ByVal ActivePageIndex As Long = -1)

'
'------------------------------------------------------------------------------
'                       REFRESH SETTINGS TAB VISUAL STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes the selected / unselected formatting of the fake settings tabs
'
' WHY THIS EXISTS
'   The hidden MultiPage no longer displays native tabs. The label-based tabs
'   must therefore reflect the active page explicitly
'
' INPUTS
'   ActivePageIndex
'     Optional zero-based active page index. If omitted, the routine reads the
'     current value from Mp_Settings
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the settings panel, MultiPage, and fake tab labels, determines the
'   active settings page, applies selected formatting to the active tab, applies
'   normal formatting to the remaining tabs, and restores tab layering
'
' ERROR POLICY
'   Raises a descriptive runtime error if the settings panel, MultiPage, tab
'   labels, or active page index cannot be resolved
'
' DEPENDENCIES
'   UF_SettingsTab_ApplyVisualState
'
' NOTES
'   This routine changes only visual state
'
'   Developer-owned settings-tab constants are intentionally not validated here.
'   They should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_SettingsTabs_RefreshVisualState"
    
    Const SETTINGS_PAGE_FIRST   As Long = 0                     'First supported settings page index
    Const SETTINGS_PAGE_LAST    As Long = 2                     'Last supported settings page index

    Dim ExistingControl         As MSForms.control              'Control resolved by name
    Dim Fra_Settings            As MSForms.Frame                'Reusable settings panel
    Dim Mp_Settings             As MSForms.MultiPage            'Settings MultiPage
    Dim Lbl_TabDisplay          As MSForms.Label                'Display tab label
    Dim Lbl_TabBehavior         As MSForms.Label                'Behavior tab label
    Dim Lbl_TabIntegration      As MSForms.Label                'Integration tab label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the settings panel
        On Error Resume Next
    'Retrieve the settings panel control
        Set ExistingControl = Me.Controls(DP_SETTINGS_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing settings panel
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Unable to resolve expected settings panel " & DP_SETTINGS_PANEL_NAME
        End If
    'Reject a name collision with a non-frame control
        If VBA.TypeName(ExistingControl) <> "Frame" Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Control '" & DP_SETTINGS_PANEL_NAME & "' exists but is not an MSForms.Frame"
        End If
    'Use the resolved settings panel frame
        Set Fra_Settings = ExistingControl

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS MULTIPAGE
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the settings MultiPage
        On Error Resume Next
    'Retrieve the settings MultiPage control
        Set ExistingControl = Fra_Settings.Controls(DP_SETTINGS_MULTIPAGE_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing settings MultiPage
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Unable to resolve expected settings MultiPage " & DP_SETTINGS_MULTIPAGE_NAME
        End If
    'Reject a name collision with a non-MultiPage control
        If VBA.TypeName(ExistingControl) <> "MultiPage" Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Control '" & DP_SETTINGS_MULTIPAGE_NAME & "' exists but is not an MSForms.MultiPage"
        End If
    'Use the resolved settings MultiPage
        Set Mp_Settings = ExistingControl

'------------------------------------------------------------------------------
' RESOLVE ACTIVE PAGE
'------------------------------------------------------------------------------
    'Resolve the active page index from the MultiPage when not supplied
        If ActivePageIndex < SETTINGS_PAGE_FIRST Then
            ActivePageIndex = Mp_Settings.Value
        End If
    'Reject unsupported active page indexes
        If ActivePageIndex < SETTINGS_PAGE_FIRST Or ActivePageIndex > SETTINGS_PAGE_LAST Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "ActivePageIndex must be between " & VBA.CStr(SETTINGS_PAGE_FIRST) & _
                " and " & VBA.CStr(SETTINGS_PAGE_LAST)
        End If
    'Reject active page indexes outside the actual MultiPage page count
        If ActivePageIndex > Mp_Settings.Pages.Count - 1 Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "ActivePageIndex exceeds the current MultiPage page count"
        End If

'------------------------------------------------------------------------------
' RETRIEVE DISPLAY TAB
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the Display tab
        On Error Resume Next
    'Retrieve the Display tab label
        Set ExistingControl = Fra_Settings.Controls(DP_SETTINGS_TAB_DISPLAY_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing Display tab
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 519, PROC_NAME, _
                "Unable to resolve expected settings tab " & DP_SETTINGS_TAB_DISPLAY_NAME
        End If
    'Reject a name collision with a non-label control
        If VBA.TypeName(ExistingControl) <> "Label" Then
            Err.Raise vbObjectError + 520, PROC_NAME, _
                "Control '" & DP_SETTINGS_TAB_DISPLAY_NAME & "' exists but is not an MSForms.Label"
        End If
    'Use the resolved Display tab label
        Set Lbl_TabDisplay = ExistingControl

'------------------------------------------------------------------------------
' RETRIEVE BEHAVIOR TAB
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the Behavior tab
        On Error Resume Next
    'Retrieve the Behavior tab label
        Set ExistingControl = Fra_Settings.Controls(DP_SETTINGS_TAB_BEHAVIOR_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing Behavior tab
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 521, PROC_NAME, _
                "Unable to resolve expected settings tab " & DP_SETTINGS_TAB_BEHAVIOR_NAME
        End If
    'Reject a name collision with a non-label control
        If VBA.TypeName(ExistingControl) <> "Label" Then
            Err.Raise vbObjectError + 522, PROC_NAME, _
                "Control '" & DP_SETTINGS_TAB_BEHAVIOR_NAME & "' exists but is not an MSForms.Label"
        End If
    'Use the resolved Behavior tab label
        Set Lbl_TabBehavior = ExistingControl

'------------------------------------------------------------------------------
' RETRIEVE INTEGRATION TAB
'------------------------------------------------------------------------------
    'Suppress lookup errors while resolving the Integration tab
        On Error Resume Next
    'Retrieve the Integration tab label
        Set ExistingControl = Fra_Settings.Controls(DP_SETTINGS_TAB_INTEGRATION_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing Integration tab
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 523, PROC_NAME, _
                "Unable to resolve expected settings tab " & DP_SETTINGS_TAB_INTEGRATION_NAME
        End If
    'Reject a name collision with a non-label control
        If VBA.TypeName(ExistingControl) <> "Label" Then
            Err.Raise vbObjectError + 524, PROC_NAME, _
                "Control '" & DP_SETTINGS_TAB_INTEGRATION_NAME & "' exists but is not an MSForms.Label"
        End If
    'Use the resolved Integration tab label
        Set Lbl_TabIntegration = ExistingControl

'------------------------------------------------------------------------------
' APPLY TAB VISUAL STATE
'------------------------------------------------------------------------------
    'Refresh the tabs visual state
        UF_SettingsTab_ApplyVisualState Lbl_TabDisplay, (ActivePageIndex = 0)
        UF_SettingsTab_ApplyVisualState Lbl_TabBehavior, (ActivePageIndex = 1)
        UF_SettingsTab_ApplyVisualState Lbl_TabIntegration, (ActivePageIndex = 2)

'------------------------------------------------------------------------------
' RESTORE LAYERING
'------------------------------------------------------------------------------
    'Move the tab to the front
        Lbl_TabDisplay.ZOrder 0
        Lbl_TabBehavior.ZOrder 0
        Lbl_TabIntegration.ZOrder 0

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
        Err.Raise Err.Number, PROC_NAME, "Settings tab visual refresh failed: " & Err.Description

End Sub

Private Sub UF_SettingsTab_ApplyVisualState( _
    ByVal TabLabel As MSForms.Label, _
    ByVal IsSelected As Boolean)

'
'------------------------------------------------------------------------------
'                       APPLY SETTINGS TAB VISUAL STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Applies selected or normal formatting to one fake settings tab
'
' WHY THIS EXISTS
'   The fake tab strip needs deterministic formatting independent from the native
'   MultiPage renderer
'
' INPUTS
'   TabLabel
'     Tab label to format
'
'   IsSelected
'     True to apply selected-tab formatting
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Applies common tab formatting and then selected or normal colors
'
' ERROR POLICY
'   Raises a descriptive runtime error if TabLabel is missing or cannot be
'   formatted
'
' DEPENDENCIES
'   None
'
' NOTES
'   This routine does not change the MultiPage value
'
'   Developer-owned tab color constants are intentionally not validated here.
'   They should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_DatePicker.UF_SettingsTab_ApplyVisualState"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing tab label
        If TabLabel Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "TabLabel cannot be Nothing"
        End If

'------------------------------------------------------------------------------
' APPLY COMMON TAB STATE
'------------------------------------------------------------------------------
    'Apply shared tab formatting
        With TabLabel
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = DP_SETTINGS_TAB_BORDER_COLOR
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = True
            .Visible = True
        End With

'------------------------------------------------------------------------------
' APPLY SELECTED OR NORMAL COLORS
'------------------------------------------------------------------------------
    'Apply selected-tab colors when requested
        If IsSelected Then
            With TabLabel
                .BackColor = DP_SETTINGS_TAB_SELECTED_BACK_COLOR
                .ForeColor = DP_SETTINGS_TAB_SELECTED_FORE_COLOR
            End With
        Else
            With TabLabel
                .BackColor = DP_SETTINGS_TAB_NORMAL_BACK_COLOR
                .ForeColor = DP_SETTINGS_TAB_NORMAL_FORE_COLOR
            End With
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
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, "Settings tab visual-state application failed: " & Err.Description

End Sub

Private Sub UF_SettingsPage_BackgroundBuild( _
    ByVal ParentMultiPage As MSForms.MultiPage, _
    ByVal PageIndex As Long, _
    ByVal BackgroundName As String)

'
'------------------------------------------------------------------------------
'                       BUILD SETTINGS PAGE BACKGROUND
'------------------------------------------------------------------------------
' PURPOSE
'   Creates a white background label inside a settings MultiPage page
'
' WHY THIS EXISTS
'   Excel MSForms.MultiPage pages are rendered with a native grey surface and do
'   not reliably expose a writable BackColor property
'
'   A disabled opaque label placed behind the page controls provides a stable
'   white page surface without relying on unsupported Page formatting
'
'   The routine receives the parent MultiPage and page index instead of a Page
'   object because runtime Page references can be fragile when controls are
'   created dynamically
'
' INPUTS
'   ParentMultiPage
'     MultiPage that owns the target page
'
'   PageIndex
'     Zero-based page index inside the MultiPage
'
'   BackgroundName
'     Name of the background label to create or reuse
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the parent MultiPage, validates the page index, retrieves the
'   target page, creates or reuses a label inside that page, formats it as a
'   passive opaque background surface, and sends it behind the page controls
'
' ERROR POLICY
'   Raises a descriptive runtime error if the MultiPage is missing, the page
'   index is invalid, the background name is blank, or the background label
'   cannot be created or formatted
'
' DEPENDENCIES
'   UF_Ensure_PageLabel
'   MSForms.MultiPage
'   MSForms.Page
'   MSForms.Label
'
' NOTES
'   The label is intentionally disabled so it does not intercept mouse events
'
'   The label dimensions are intentionally constant-driven because MSForms.Page
'   does not expose reliable inside-width / inside-height properties in all
'   Excel VBA environments
'
'   Developer-owned settings-page layout constants are intentionally not
'   validated here. They should be checked by a dedicated debug or regression
'   routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_SettingsPage_BackgroundBuild"
    
    Dim ParentPage          As MSForms.Page         'Resolved target page
    Dim Lbl_Background      As MSForms.Label        'Page background label
    Dim EffectiveName       As String               'Normalized background label name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing parent MultiPage
        If ParentMultiPage Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentMultiPage cannot be Nothing"
        End If
    'Reject a parent MultiPage with no pages
        If ParentMultiPage.Pages.Count = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ParentMultiPage must contain at least one page"
        End If
    'Reject an invalid page index
        If PageIndex < 0 Or PageIndex > ParentMultiPage.Pages.Count - 1 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "PageIndex must be between 0 and " & VBA.CStr(ParentMultiPage.Pages.Count - 1)
        End If
    'Normalize the background label name
        EffectiveName = VBA.Trim$(BackgroundName)
    'Reject an empty background name
        If VBA.Len(EffectiveName) = 0 Then
            Err.Raise vbObjectError + 516, PROC_NAME, "BackgroundName cannot be empty"
        End If

'------------------------------------------------------------------------------
' RETRIEVE TARGET PAGE
'------------------------------------------------------------------------------
    'Retrieve the target MultiPage page
        Set ParentPage = ParentMultiPage.Pages(PageIndex)

'------------------------------------------------------------------------------
' CREATE / FORMAT BACKGROUND LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the page background label
        Set Lbl_Background = UF_Ensure_PageLabel(ParentPage, EffectiveName)
    'Apply layout and visual properties to the page background label
        With Lbl_Background
            .Caption = VBA.vbNullString
            .Left = DP_SETTINGS_PAGE_BACK_LEFT
            .Top = DP_SETTINGS_PAGE_BACK_TOP
            .Width = DP_SETTINGS_PAGE_BACK_WIDTH
            .Height = DP_SETTINGS_PAGE_BACK_HEIGHT
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_SETTINGS_PAGE_BACK_COLOR
            .ForeColor = DP_SETTINGS_PAGE_BACK_COLOR
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = False
            .Visible = True
        End With
    'Send the background label behind the page controls
        Lbl_Background.ZOrder 1

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
        Err.Raise Err.Number, PROC_NAME, "Settings page background creation failed: " & Err.Description

End Sub

Private Sub UF_SettingsCheckBox_Format( _
    ByVal TargetCheckBox As MSForms.CheckBox, _
    ByVal CaptionText As String, _
    ByVal ControlLeft As Single, _
    ByVal ControlTop As Single)

'
'------------------------------------------------------------------------------
'                           FORMAT SETTINGS CHECKBOX
'------------------------------------------------------------------------------
' PURPOSE
'   Applies deterministic formatting to one settings CheckBox
'
' WHY THIS EXISTS
'   Runtime-created and reused MSForms CheckBox controls can retain stale font
'   states when their existing Font object is edited in place or when a shared
'   font object is reused
'
'   Assigning a fresh cloned StdFont object to each CheckBox removes stale
'   Italic, Underline, and Strikethrough states consistently
'
' INPUTS
'   TargetCheckBox
'     CheckBox control to format
'
'   CaptionText
'     Caption to display
'
'   ControlLeft
'     Left position inside the parent page
'
'   ControlTop
'     Top position inside the parent page
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the caption, applies layout and visual state, assigns a fresh font,
'   restores the caption, and preserves the current CheckBox value
'
' ERROR POLICY
'   Raises a descriptive runtime error if TargetCheckBox is missing or the
'   CheckBox cannot be formatted
'
' DEPENDENCIES
'   UF_SettingsCheckBoxFont_Create
'   DP_SETTINGS_CHECKBOX_WIDTH
'   DP_SETTINGS_CHECKBOX_HEIGHT
'
' NOTES
'   The caption is assigned after the fresh font is assigned
'
'   The CheckBox Value property is intentionally not changed here. Value refresh
'   is handled separately by UF_SettingsPanel_RefreshCaptions
'
'   Developer-owned CheckBox layout constants are intentionally not validated
'   here. They should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_SettingsCheckBox_Format"

    Dim CheckBoxFont        As Object       'Fresh CheckBox font object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing CheckBox
        If TargetCheckBox Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "TargetCheckBox cannot be Nothing"
        End If

'------------------------------------------------------------------------------
' CREATE FRESH FONT
'------------------------------------------------------------------------------
    'Create a fresh independent CheckBox font
        Set CheckBoxFont = UF_SettingsCheckBoxFont_Create()

'------------------------------------------------------------------------------
' FORMAT CHECKBOX
'------------------------------------------------------------------------------
    'Apply deterministic CheckBox formatting
        With TargetCheckBox
            .Caption = VBA.vbNullString
            .Left = ControlLeft
            .Top = ControlTop
            .Width = DP_SETTINGS_CHECKBOX_WIDTH
            .Height = DP_SETTINGS_CHECKBOX_HEIGHT
            .AutoSize = False
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_FOOTER_FORE_COLOR
            .TripleState = False
            .SpecialEffect = fmSpecialEffectFlat
            .TabStop = False
            .Enabled = True
            .Visible = True
            .ControlTipText = CaptionText
        End With
    'Assign the fresh font object
        Set TargetCheckBox.Font = CheckBoxFont
    'Apply the caption after assigning the fresh font
        TargetCheckBox.Caption = CaptionText

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
        Err.Raise Err.Number, PROC_NAME, "Settings CheckBox formatting failed: " & Err.Description

End Sub

Private Function UF_SettingsCheckBoxFont_Create() As Object

'
'------------------------------------------------------------------------------
'                       CREATE SETTINGS CHECKBOX FONT
'------------------------------------------------------------------------------
' PURPOSE
'   Creates a fresh StdFont object for one settings CheckBox
'
' WHY THIS EXISTS
'   Reused MSForms CheckBox controls can retain stale font states such as Italic,
'   Underline, or Strikethrough
'
'   Assigning a newly created font object to each CheckBox gives each control a
'   clean independent font state
'
' INPUTS
'   None
'
' RETURNS
'   Fresh StdFont object configured for settings CheckBoxes
'
' BEHAVIOR
'   Creates and returns a new StdFont object with deterministic font properties
'
' ERROR POLICY
'   Raises a descriptive runtime error if the font object cannot be created or
'   configured
'
' DEPENDENCIES
'   StdFont
'   DP_FORM_FONT_NAME
'   DP_SETTINGS_CHECKBOX_FONT_SIZE
'
' NOTES
'   Do not reuse the returned font object across multiple CheckBoxes
'
'   Developer-owned font constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_DatePicker.UF_SettingsCheckBoxFont_Create"
    
    Dim NewFont         As Object       'Fresh CheckBox font object

    Dim SavedErrNumber          As Long     'Captured original error number
    Dim SavedErrDescription     As String   'Captured original error description

    Dim CleanupErrNumber        As Long     'Captured cleanup error number
    Dim CleanupErrDescription   As String   'Captured cleanup error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CREATE FONT
'------------------------------------------------------------------------------
    'Create a fresh font object
        Set NewFont = CreateObject("StdFont")
    'Configure the fresh font object
        With NewFont
            .Name = DP_FORM_FONT_NAME
            .Size = DP_SETTINGS_CHECKBOX_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

'------------------------------------------------------------------------------
' RETURN FONT
'------------------------------------------------------------------------------
    'Return the fresh font object
        Set UF_SettingsCheckBoxFont_Create = NewFont

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error before any cleanup runs
    '
    'Every On Error statement resets the Err object, so the On Error Resume Next
    'below already destroys the cause on its own, and On Error GoTo 0 destroys it
    'again. At v1.2.0 the raise that followed read Err.Number and Err.Description
    'after both, so this handler always reported error 0 with a blank cause rather
    'than the failure that actually occurred. Nothing may read the live Err object
    'below this point: see #48
        SavedErrNumber = Err.Number
        SavedErrDescription = "Settings CheckBox font creation failed: " & Err.Description
    'Suppress cleanup errors
        On Error Resume Next
    'Release any partially configured font object
        Set NewFont = Nothing
    'Capture a cleanup failure separately rather than letting it replace the
    'primary failure
        If Err.Number <> 0 Then
            CleanupErrNumber = Err.Number
            CleanupErrDescription = Err.Description
        End If
    'Restore normal error handling
        On Error GoTo 0
    'Append cleanup diagnostics when cleanup also failed
        If CleanupErrNumber <> 0 Then
            SavedErrDescription = SavedErrDescription & _
                " Cleanup also failed while releasing the font object: " & _
                CleanupErrDescription
        End If
    'Raise the original error after best-effort cleanup
        Err.Raise SavedErrNumber, PROC_NAME, SavedErrDescription

End Function

'
'------------------------------------------------------------------------------
'
'                     PUBLIC REFRESH AND CLOCK ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub UF_DP_AfterSuccessfulSelection(ByVal SelectedDate As Date)

'
'------------------------------------------------------------------------------
'                           AFTER SUCCESSFUL SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Synchronizes the open DatePicker form after a successful date write-back
'
' WHY THIS EXISTS
'   When close-after-selection is disabled, the form remains open after the user
'   clicks a day label. The form must clear the old selected-day visual state and
'   apply the new selected-day visual state without rebuilding the full calendar
'   grid while the label click event is still active
'
' INPUTS
'   SelectedDate
'     Date successfully written to the Excel target
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Captures the previous keyboard-selected date, stores the new selected date as
'   the active keyboard date, hides the picker panel when present, clears
'   transient hover state, restores only the previous selected day when visible,
'   and refreshes only the newly selected day when visible
'
' ERROR POLICY
'   Best-effort UI synchronization. Suppresses synchronization failures because
'   the Excel write-back has already succeeded and the user should not receive a
'   runtime error after a successful selection
'
' DEPENDENCIES
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_PickerPanel_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'   UF_DayCell_RefreshVisibleDate
'
' NOTES
'   This routine intentionally does not call UF_DayGrid_Populate
'
'   Rebuilding the full grid from inside a label click event can raise MSForms
'   run-time error 5 in some environments
'
'   The routine updates the open form only. The canonical selected-date state is
'   maintained by the companion module during write-back
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_DP_AfterSuccessfulSelection"

    Dim ExistingControl         As MSForms.control          'Control resolved by picker-panel name
    Dim Fra_PickerPanel         As MSForms.Frame            'Reusable picker panel
    
    Dim PreviousDate            As Date                     'Previously selected keyboard date
    Dim HasPreviousDate         As Boolean                  'True when previous keyboard date exists
    Dim SelectedDateOnly        As Date                     'Selected date without time

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' CAPTURE PREVIOUS STATE
'------------------------------------------------------------------------------
    'Capture the previous keyboard-selected date when available
        If mHasKeyboardDate Then
            PreviousDate = VBA.DateValue(mKeyboardDate)
            HasPreviousDate = True
        End If
    'Resolve the selected date without time
        SelectedDateOnly = VBA.DateValue(SelectedDate)

'------------------------------------------------------------------------------
' STORE NEW KEYBOARD STATE
'------------------------------------------------------------------------------
    'Store the selected date as the active keyboard navigation date
        mKeyboardDate = SelectedDateOnly
    'Mark keyboard navigation date as available
        mHasKeyboardDate = True

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Clear active picker hover before hiding the picker panel
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Suppress missing picker-panel lookup errors
        On Error Resume Next
    'Retrieve the picker panel control when available
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore fail-safe error handling
        On Error GoTo FailSafe
    'Hide the picker panel when it exists and is a frame
        If Not ExistingControl Is Nothing Then
            If VBA.TypeName(ExistingControl) = "Frame" Then
                Set Fra_PickerPanel = ExistingControl
                Fra_PickerPanel.Visible = False
            End If
        End If

    'Clear picker-panel mode
        mPickerPanelMode = 0
    'Clear picker-panel hover state
        mHoveredPickerItemIndex = 0

'------------------------------------------------------------------------------
' CLEAR TRANSIENT HOVER STATE
'------------------------------------------------------------------------------
    'Clear active day-cell hover after selection
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Clear active header hover after selection
        If VBA.Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Clear active footer hover after selection
        If VBA.Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset
    'Clear active settings-panel hover after selection
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then UF_SettingsPanel_HoverReset
        
'------------------------------------------------------------------------------
' REFRESH ONLY AFFECTED DAY CELLS
'------------------------------------------------------------------------------
    'Restore the previous selected day when it is visible and changed
        If HasPreviousDate Then
            If PreviousDate <> SelectedDateOnly Then
                UF_DayCell_RefreshVisibleDate PreviousDate
            End If
        End If

    'Apply the selected visual state to the newly selected day when visible
        UF_DayCell_RefreshVisibleDate SelectedDateOnly

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the fail-safe handler
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Suppress secondary cleanup errors
        On Error Resume Next

    'Clear transient hover trackers
        mHoveredDayCellIndex = 0
        mHoveredHeaderLabelName = vbNullString
        mHoveredPickerItemIndex = 0
        mHoveredFooterActionName = vbNullString
        mHoveredSettingsPanelLabelName = vbNullString

    'Clear picker-panel mode defensively
        mPickerPanelMode = 0

    'Restore normal error handling
        On Error GoTo 0

End Sub
'------------------------------------------------------------------------------
' PUBLIC REFRESH AND CLOCK ENTRY POINTS
'------------------------------------------------------------------------------

Public Sub UF_DP_RefreshFromExternalSelection( _
    ByVal SelectedDate As Date, _
    ByVal HasSelectedDate As Boolean)

'
'------------------------------------------------------------------------------
'                       REFRESH FROM EXTERNAL SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes the open DatePicker form after the user selects another worksheet
'   cell while the form remains open
'
' WHY THIS EXISTS
'   When close-after-selection is disabled, the form should remain available for
'   repeated use across different worksheet cells. Changing ActiveCell should
'   update or clear the selected-date visual state instead of leaving stale
'   highlights in the calendar grid
'
' INPUTS
'   SelectedDate
'     Date resolved from the new ActiveCell, or ignored when HasSelectedDate is
'     False
'
'   HasSelectedDate
'     True when the new ActiveCell contains a valid date
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   If ActiveCell contains a date, displays that date's month and highlights it.
'   If ActiveCell does not contain a date, clears selected-date / keyboard-date
'   state and displays the current system month. In both cases, clears transient
'   hover state, hides overlay panels, and repopulates the day grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if the form cannot be refreshed
'
' DEPENDENCIES
'   UF_DisplayPeriod_Initialize
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_PickerPanel_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'   UF_SettingsPanel_Hide
'   UF_DayGrid_Populate
'
' NOTES
'   This routine does not write to Excel
'
'   The selected-date state is aligned with the external worksheet selection so
'   that day-cell rendering does not fall back to stale global selection state
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_DP_RefreshFromExternalSelection"

    Dim ExistingControl     As MSForms.control      'Control resolved by picker-panel name
    Dim Fra_PickerPanel     As MSForms.Frame        'Reusable picker panel
    
    Dim DisplayDate         As Date                 'Date used to resolve displayed month
    Dim HandlerStep         As String               'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' RESET TRANSIENT HOVER STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Reset transient hover state"

    'Remove current day-cell hover formatting
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Remove current header hover formatting
        If VBA.Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Remove current picker-panel hover formatting
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Remove current footer hover formatting
        If VBA.Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset
    'Remove current settings-panel hover formatting
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then UF_SettingsPanel_HoverReset

'------------------------------------------------------------------------------
' RESOLVE DISPLAY DATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve display date"

    'Use the selected date when ActiveCell contains a valid date
        If HasSelectedDate Then
            'Normalize the selected date to a date-only value
                DisplayDate = VBA.DateValue(SelectedDate)
        Else
            'Use the current system date when ActiveCell has no date
                DisplayDate = VBA.Date
        End If

'------------------------------------------------------------------------------
' STORE SELECTION STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Store selection state"

    'Store selected state when the new ActiveCell contains a date
        If HasSelectedDate Then
            'Store the selected date as the active keyboard navigation date
                mKeyboardDate = DisplayDate
            'Mark keyboard navigation date as available
                mHasKeyboardDate = True
            'Align the global selected-date state with the external selection
                gDP_SelectedDate = DisplayDate
            'Mark the global selected-date state as available
                gDP_HasSelectedDate = True
        Else

            'Clear the keyboard navigation date
                mKeyboardDate = 0
            'Mark keyboard navigation date as unavailable
                mHasKeyboardDate = False
            'Clear the global selected-date value
                gDP_SelectedDate = 0
            'Mark the global selected-date state as unavailable
                gDP_HasSelectedDate = False

        End If

'------------------------------------------------------------------------------
' STORE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Initialize display period"

    'Initialize the displayed month and year from the resolved display date
        UF_DisplayPeriod_Initialize DisplayDate

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Hide picker panel"

    'Suppress lookup errors when the picker panel is not available
        On Error Resume Next
    'Retrieve the picker panel control when available
        Set ExistingControl = Me.Controls(DP_PICKER_PANEL_NAME)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Hide the picker panel when it exists
        If Not ExistingControl Is Nothing Then
            'Reject a name collision with a non-frame control
                If VBA.TypeName(ExistingControl) <> "Frame" Then
                    Err.Raise vbObjectError + 513, PROC_NAME, _
                        "Control '" & DP_PICKER_PANEL_NAME & "' exists but is not an MSForms.Frame"
                End If
            'Use the resolved picker panel
                Set Fra_PickerPanel = ExistingControl
            'Hide the picker panel
                Fra_PickerPanel.Visible = False
        End If
    'Clear picker-panel mode
        mPickerPanelMode = 0
    'Clear picker-panel hover state
        mHoveredPickerItemIndex = 0

'------------------------------------------------------------------------------
' HIDE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Hide settings panel"

    'Hide the settings panel after external worksheet selection
        UF_SettingsPanel_Hide

'------------------------------------------------------------------------------
' CLEAR STORED HOVER TRACKERS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Clear stored hover trackers"

    'Clear the day-cell hover tracker defensively
        mHoveredDayCellIndex = 0
    'Clear the header hover tracker defensively
        mHoveredHeaderLabelName = vbNullString
    'Clear the picker-panel hover tracker defensively
        mHoveredPickerItemIndex = 0
    'Clear the footer hover tracker defensively
        mHoveredFooterActionName = vbNullString
    'Clear the settings-panel hover tracker defensively
        mHoveredSettingsPanelLabelName = vbNullString

'------------------------------------------------------------------------------
' REFRESH CALENDAR GRID
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh day grid"

    'Refresh the calendar grid from the new worksheet selection
        UF_DayGrid_Populate mDisplayYear, mDisplayMonth

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
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "External selection refresh failed: " & Err.Description

End Sub
'------------------------------------------------------------------------------
' PUBLIC REFRESH AND CLOCK ENTRY POINTS
'------------------------------------------------------------------------------

Public Sub UF_DP_RefreshSettings()

'
'------------------------------------------------------------------------------
'                       REFRESH SETTINGS-DEPENDENT CAPTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes DatePicker captions affected by saved display settings
'
' WHY THIS EXISTS
'   Changing UseLocalNames, FirstDayOfWeek, weekend highlighting, compact layout,
'   or related display settings while the form is already open must refresh the
'   visible captions and calendar-grid visual state without rebuilding the whole
'   UserForm
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures the display period is usable, clears active hover state, refreshes
'   header labels, weekday labels, Today footer value, and the visible day grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if any settings-dependent refresh step
'   fails
'
' DEPENDENCIES
'   UF_DisplayPeriod_Initialize
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_PickerPanel_HoverReset
'   UF_Footer_HoverReset
'   UF_SettingsPanel_HoverReset
'   UF_Header_BuildLabels
'   UF_WeekdayRow_Build
'   M_Caption_GetDate
'   UF_DayGrid_Populate
'
' NOTES
'   This routine deliberately avoids recreating the full footer section
'
'   Time / live-clock behavior is handled separately by M_Timer_ApplyClockMode
'
'   Developer-owned layout constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_DP_RefreshSettings"

    Dim ExistingControl     As MSForms.control      'Control resolved by name
    Dim Lbl_Today           As MSForms.Label        'Footer Today value label
    
    Dim HandlerStep         As String               'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' ENSURE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure display period"

    'Initialize the display period when it is not usable
        If mDisplayYear = 0 Or mDisplayMonth < 1 Or mDisplayMonth > 12 Then
            UF_DisplayPeriod_Initialize VBA.Date
        End If

'------------------------------------------------------------------------------
' RESET TRANSIENT HOVER STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Reset transient hover state"

    'Clear active day-cell hover before refreshing the grid
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Clear active header hover before refreshing header labels
        If VBA.Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Clear active picker-panel hover before refreshing display state
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Clear active footer hover before refreshing footer value captions
        If VBA.Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset
    'Clear active settings-panel hover before refreshing settings-dependent UI
        If VBA.Len(mHoveredSettingsPanelLabelName) <> 0 Then UF_SettingsPanel_HoverReset

'------------------------------------------------------------------------------
' REFRESH HEADER CAPTIONS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh header labels"

    'Refresh the month and year header labels
        UF_Header_BuildLabels

'------------------------------------------------------------------------------
' REFRESH DAY-OF-WEEK CAPTIONS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh weekday row"

    'Refresh the weekday header row
        UF_WeekdayRow_Build

'------------------------------------------------------------------------------
' RETRIEVE TODAY VALUE LABEL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve Today footer label"

    'Suppress lookup errors while resolving the Today value label
        On Error Resume Next
    'Retrieve the Today value label
        Set ExistingControl = Me.Controls("Lbl_Today")
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Reject a missing Today value label
        If ExistingControl Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Unable to resolve expected footer label Lbl_Today"
        End If
    'Reject a name collision with a non-label control
        If VBA.TypeName(ExistingControl) <> "Label" Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Control 'Lbl_Today' exists but is not an MSForms.Label"
        End If
    'Use the resolved Today value label
        Set Lbl_Today = ExistingControl

'------------------------------------------------------------------------------
' REFRESH TODAY CAPTION
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh Today footer caption"

    'Refresh the Today value label
        Lbl_Today.Caption = M_Caption_GetDate(VBA.Date, gDP_UseLocalNames)

'------------------------------------------------------------------------------
' REFRESH DAY GRID
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh day grid"

    'Refresh the day-label grid
        UF_DayGrid_Populate mDisplayYear, mDisplayMonth

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
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Settings-dependent DatePicker refresh failed: " & Err.Description

End Sub

Public Sub UF_DP_UpdateLiveClock()

'
'------------------------------------------------------------------------------
'                           UPDATE LIVE CLOCK
'------------------------------------------------------------------------------
' PURPOSE
'   Updates the DatePicker footer time caption
'
' WHY THIS EXISTS
'   The live clock timer should update only the time text without rebuilding,
'   repainting, refreshing, or repeatedly resolving UserForm controls
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Uses the cached Lbl_Time reference when available, restores the cache from
'   the UserForm controls collection only when needed, and updates the caption
'   only when the displayed time value has changed
'
' ERROR POLICY
'   Best-effort UI update. Suppresses update failures because timer callbacks may
'   occur while the form is being hidden, unloaded, or rebuilt
'
' DEPENDENCIES
'   UserForm.Controls collection
'   Lbl_Time
'   mLbl_Time
'
' NOTES
'   This routine deliberately does not call Repaint
'
'   This routine deliberately does not rebuild the footer section
'
'   Missing or stale time-label references are treated as safe no-ops
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ExistingControl     As MSForms.control      'Control resolved by name
    Dim NewTimeCaption      As String               'Current formatted time caption

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' BUILD TIME CAPTION
'------------------------------------------------------------------------------
    'Build the new time caption
        NewTimeCaption = VBA.Format$(VBA.Time, "hh:nn:ss")

'------------------------------------------------------------------------------
' RESTORE CACHE WHEN NEEDED
'------------------------------------------------------------------------------
    'Restore the cached Time label reference only when missing
        If mLbl_Time Is Nothing Then
            'Suppress lookup errors while resolving the Time value label
                On Error Resume Next
            'Retrieve the Time value label
                Set ExistingControl = Me.Controls("Lbl_Time")
            'Clear any suppressed lookup error
                Err.Clear
            'Restore fail-safe error handling
                On Error GoTo FailSafe

            'Exit when the Time value label is not available
                If ExistingControl Is Nothing Then Exit Sub
            'Exit when the resolved control is not a label
                If VBA.TypeName(ExistingControl) <> "Label" Then Exit Sub
            'Cache the resolved Time value label
                Set mLbl_Time = ExistingControl
        End If

'------------------------------------------------------------------------------
' UPDATE TIME LABEL
'------------------------------------------------------------------------------
    'Update the label only if the caption changed
        If mLbl_Time.Caption <> NewTimeCaption Then
            mLbl_Time.Caption = NewTimeCaption
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the fail-safe handler
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Suppress timer-update failures
        On Error Resume Next
    'Clear the stale cached Time label reference
        Set mLbl_Time = Nothing
    'Clear any pending timer-update error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub


'
'------------------------------------------------------------------------------
'
'                   LOW-LEVEL CREATE / LAYOUT / VALIDATION HELPERS
'
'------------------------------------------------------------------------------
'


Private Function UF_Ensure_Label(ByVal ControlName As String) As MSForms.Label

'
'------------------------------------------------------------------------------
'                               ENSURE LABEL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns an existing MSForms.Label control or creates it if missing
'
' WHY THIS EXISTS
'   Many DatePicker UI elements are runtime-created labels. Centralizing the
'   create-or-reuse logic avoids repeated boilerplate across routines
'
' INPUTS
'   ControlName
'     Name of the label control to retrieve or create
'
' RETURNS
'   The requested MSForms.Label control
'
' BEHAVIOR
'   Normalizes the requested control name, reuses an existing MSForms.Label when
'   available, or creates a new label directly on the UserForm when missing
'
' ERROR POLICY
'   Raises a descriptive runtime error if ControlName is blank, if an existing
'   control with the same name is not a label, or if the label cannot be created
'
' DEPENDENCIES
'   UserForm.Controls collection
'
' NOTES
'   This routine creates controls on the UserForm itself, not inside a Frame,
'   Page, or MultiPage
'
'   This routine does not format the label. Formatting remains the responsibility
'   of the caller
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Ensure_Label"

    Dim EffectiveName       As String               'Normalized control name
    
    Dim ExistingControl     As MSForms.control      'Existing runtime control
    Dim Lbl                 As MSForms.Label        'Target label control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Normalize the requested control name
        EffectiveName = VBA.Trim$(ControlName)
    'Reject an empty control name
        If VBA.Len(EffectiveName) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ControlName cannot be empty"
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Suppress lookup errors when the control does not exist yet
        On Error Resume Next
    'Try to retrieve an existing control with the requested name
        Set ExistingControl = Me.Controls(EffectiveName)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING LABEL
'------------------------------------------------------------------------------
    'Reuse the existing control when it is available
        If Not ExistingControl Is Nothing Then
            'Reject an existing non-label control with the same name
                If VBA.TypeName(ExistingControl) <> "Label" Then
                    Err.Raise vbObjectError + 514, PROC_NAME, _
                        "Control '" & EffectiveName & "' already exists but is not an MSForms.Label"
                End If
            'Return the existing label
                Set UF_Ensure_Label = ExistingControl
            'Exit after returning the existing label
                Exit Function
        End If

'------------------------------------------------------------------------------
' CREATE LABEL IF MISSING
'------------------------------------------------------------------------------
    'Create the label directly on the UserForm
        Set ExistingControl = Me.Controls.Add("Forms.Label.1", EffectiveName, True)
    'Reject an unexpected created control type
        If VBA.TypeName(ExistingControl) <> "Label" Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Control '" & EffectiveName & "' was created but is not an MSForms.Label"
        End If
    'Use the created label
        Set Lbl = ExistingControl

'------------------------------------------------------------------------------
' RETURN LABEL
'------------------------------------------------------------------------------
    'Return the created label
        Set UF_Ensure_Label = Lbl

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
        Err.Raise Err.Number, PROC_NAME, "UserForm label ensure failed: " & Err.Description

End Function

Private Function UF_Ensure_FrameLabel( _
    ByVal ParentFrame As MSForms.Frame, _
    ByVal ControlName As String) As MSForms.Label

'
'------------------------------------------------------------------------------
'                           ENSURE FRAME LABEL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns an existing MSForms.Label inside a frame or creates it if missing
'
' WHY THIS EXISTS
'   The month/year picker panel, settings panel, footer groups, and other
'   overlay surfaces use labels hosted inside MSForms.Frame controls. Those
'   labels must be created in the frame's Controls collection, not directly on
'   the UserForm
'
' INPUTS
'   ParentFrame
'     Frame that owns the label
'
'   ControlName
'     Name of the label control to retrieve or create
'
' RETURNS
'   The requested MSForms.Label control
'
' BEHAVIOR
'   Normalizes the requested control name, reuses an existing MSForms.Label
'   inside the supplied frame when available, or creates a new label when missing
'
' ERROR POLICY
'   Raises a descriptive runtime error if ParentFrame is missing, ControlName is
'   blank, an existing control with the same name is not a label, or the label
'   cannot be created
'
' DEPENDENCIES
'   MSForms.Frame
'   MSForms.Label
'
' NOTES
'   This routine creates controls inside the supplied frame
'
'   This routine does not format the label. Formatting remains the responsibility
'   of the caller
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Ensure_FrameLabel"

    Dim EffectiveName       As String               'Normalized control name
    
    Dim ExistingControl     As MSForms.control      'Existing frame control
    Dim Lbl                 As MSForms.Label        'Target label control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing parent frame
        If ParentFrame Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentFrame cannot be Nothing"
        End If
    'Normalize the requested control name
        EffectiveName = VBA.Trim$(ControlName)
    'Reject an empty control name
        If VBA.Len(EffectiveName) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ControlName cannot be empty"
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Suppress lookup errors when the control does not exist yet
        On Error Resume Next
    'Try to retrieve an existing control with the requested name
        Set ExistingControl = ParentFrame.Controls(EffectiveName)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING LABEL
'------------------------------------------------------------------------------
    'Reuse the existing control when it is available
        If Not ExistingControl Is Nothing Then
            'Reject an existing non-label control with the same name
                If VBA.TypeName(ExistingControl) <> "Label" Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Control '" & EffectiveName & "' already exists but is not an MSForms.Label"
                End If
            'Use the existing label
                Set Lbl = ExistingControl
            'Return the existing label
                Set UF_Ensure_FrameLabel = Lbl
            'Exit after returning the existing label
                Exit Function
        End If

'------------------------------------------------------------------------------
' CREATE LABEL IF MISSING
'------------------------------------------------------------------------------
    'Create the label inside the supplied frame
        Set ExistingControl = ParentFrame.Controls.Add("Forms.Label.1", EffectiveName, True)
    'Reject an unexpected created control type
        If VBA.TypeName(ExistingControl) <> "Label" Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Control '" & EffectiveName & "' was created but is not an MSForms.Label"
        End If
    'Use the created label
        Set Lbl = ExistingControl

'------------------------------------------------------------------------------
' RETURN LABEL
'------------------------------------------------------------------------------
    'Return the created label
        Set UF_Ensure_FrameLabel = Lbl

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
        Err.Raise Err.Number, PROC_NAME, "Frame label ensure failed: " & Err.Description

End Function

Private Function UF_Ensure_FrameMultiPage( _
    ByVal ParentFrame As MSForms.Frame, _
    ByVal ControlName As String) As MSForms.MultiPage

'
'------------------------------------------------------------------------------
'                           ENSURE FRAME MULTIPAGE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns an existing MSForms.MultiPage inside a frame or creates it if missing
'
' WHY THIS EXISTS
'   The settings panel hosts a runtime MultiPage control inside Fra_Settings
'
' INPUTS
'   ParentFrame
'     Frame that owns the MultiPage
'
'   ControlName
'     Name of the MultiPage control to retrieve or create
'
' RETURNS
'   The requested MSForms.MultiPage control
'
' BEHAVIOR
'   Normalizes the requested control name, reuses an existing MSForms.MultiPage
'   inside the supplied frame when available, or creates a new MultiPage when
'   missing
'
' ERROR POLICY
'   Raises a descriptive runtime error if ParentFrame is missing, ControlName is
'   blank, an existing control with the same name is not a MultiPage, or the
'   MultiPage cannot be created
'
' DEPENDENCIES
'   MSForms.Frame
'   MSForms.MultiPage
'
' NOTES
'   This routine creates controls inside the supplied frame
'
'   This routine does not format the MultiPage. Formatting remains the
'   responsibility of the caller
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Ensure_FrameMultiPage"

    Dim EffectiveName       As String                   'Normalized control name
    
    Dim ExistingControl     As MSForms.control          'Existing frame control
    Dim Mp                  As MSForms.MultiPage        'Target MultiPage control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing parent frame
        If ParentFrame Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentFrame cannot be Nothing"
        End If
    'Normalize the requested control name
        EffectiveName = VBA.Trim$(ControlName)
    'Reject an empty control name
        If VBA.Len(EffectiveName) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ControlName cannot be empty"
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Suppress lookup errors when the control does not exist yet
        On Error Resume Next
    'Try to retrieve an existing control with the requested name
        Set ExistingControl = ParentFrame.Controls(EffectiveName)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING MULTIPAGE
'------------------------------------------------------------------------------
    'Reuse the existing control when it is available
        If Not ExistingControl Is Nothing Then
            'Reject an existing non-MultiPage control with the same name
                If VBA.TypeName(ExistingControl) <> "MultiPage" Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Control '" & EffectiveName & "' already exists but is not an MSForms.MultiPage"
                End If
            'Use the existing MultiPage
                Set Mp = ExistingControl
            'Return the existing MultiPage
                Set UF_Ensure_FrameMultiPage = Mp
            'Exit after returning the existing MultiPage
                Exit Function
        End If

'------------------------------------------------------------------------------
' CREATE MULTIPAGE IF MISSING
'------------------------------------------------------------------------------
    'Create the MultiPage inside the supplied frame
        Set ExistingControl = ParentFrame.Controls.Add("Forms.MultiPage.1", EffectiveName, True)
    'Reject an unexpected created control type
        If VBA.TypeName(ExistingControl) <> "MultiPage" Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Control '" & EffectiveName & "' was created but is not an MSForms.MultiPage"
        End If
    'Use the created MultiPage
        Set Mp = ExistingControl

'------------------------------------------------------------------------------
' RETURN MULTIPAGE
'------------------------------------------------------------------------------
    'Return the created MultiPage
        Set UF_Ensure_FrameMultiPage = Mp

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
        Err.Raise Err.Number, PROC_NAME, "Frame MultiPage ensure failed: " & Err.Description

End Function

Private Function UF_Ensure_PageLabel( _
    ByVal ParentPage As MSForms.Page, _
    ByVal ControlName As String) As MSForms.Label

'
'------------------------------------------------------------------------------
'                           ENSURE PAGE LABEL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns an existing MSForms.Label inside a MultiPage page or creates it if
'   missing
'
' WHY THIS EXISTS
'   Runtime settings pages need labels hosted inside MSForms.Page controls.
'   Centralizing the create-or-reuse logic avoids repeated boilerplate across
'   settings-panel build routines
'
' INPUTS
'   ParentPage
'     Page that owns the label
'
'   ControlName
'     Name of the label control to retrieve or create
'
' RETURNS
'   The requested MSForms.Label control
'
' BEHAVIOR
'   Normalizes the requested control name, reuses an existing MSForms.Label
'   inside the supplied MultiPage page when available, or creates a new label
'   when missing
'
' ERROR POLICY
'   Raises a descriptive runtime error if ParentPage is missing, ControlName is
'   blank, an existing control with the same name is not a label, or the label
'   cannot be created
'
' DEPENDENCIES
'   MSForms.Page
'   MSForms.Label
'
' NOTES
'   This routine creates controls inside the supplied MultiPage page
'
'   This routine does not format the label. Formatting remains the responsibility
'   of the caller
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Ensure_PageLabel"

    Dim EffectiveName       As String               'Normalized control name
    
    Dim ExistingControl     As MSForms.control      'Existing page control
    Dim Lbl                 As MSForms.Label        'Target label control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing parent page
        If ParentPage Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentPage cannot be Nothing"
        End If
    'Normalize the requested control name
        EffectiveName = VBA.Trim$(ControlName)
    'Reject an empty control name
        If VBA.Len(EffectiveName) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ControlName cannot be empty"
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Suppress lookup errors when the control does not exist yet
        On Error Resume Next
    'Try to retrieve an existing control with the requested name
        Set ExistingControl = ParentPage.Controls(EffectiveName)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING LABEL
'------------------------------------------------------------------------------
    'Reuse the existing control when it is available
        If Not ExistingControl Is Nothing Then
            'Reject an existing non-label control with the same name
                If VBA.TypeName(ExistingControl) <> "Label" Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Control '" & EffectiveName & "' already exists but is not an MSForms.Label"
                End If
            'Use the existing label
                Set Lbl = ExistingControl
            'Return the existing label
                Set UF_Ensure_PageLabel = Lbl
            'Exit after returning the existing label
                Exit Function
        End If

'------------------------------------------------------------------------------
' CREATE LABEL IF MISSING
'------------------------------------------------------------------------------
    'Create the label inside the supplied MultiPage page
        Set ExistingControl = ParentPage.Controls.Add("Forms.Label.1", EffectiveName, True)
    'Reject an unexpected created control type
        If VBA.TypeName(ExistingControl) <> "Label" Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Control '" & EffectiveName & "' was created but is not an MSForms.Label"
        End If
    'Use the created label
        Set Lbl = ExistingControl

'------------------------------------------------------------------------------
' RETURN LABEL
'------------------------------------------------------------------------------
    'Return the created label
        Set UF_Ensure_PageLabel = Lbl

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
        Err.Raise Err.Number, PROC_NAME, "Page label ensure failed: " & Err.Description

End Function

Private Function UF_Ensure_PageComboBox( _
    ByVal ParentPage As MSForms.Page, _
    ByVal ControlName As String) As MSForms.ComboBox

'
'------------------------------------------------------------------------------
'                           ENSURE PAGE COMBOBOX
'------------------------------------------------------------------------------
' PURPOSE
'   Returns an existing MSForms.ComboBox inside a MultiPage page or creates it if
'   missing
'
' WHY THIS EXISTS
'   Runtime settings pages need ComboBox controls hosted inside MSForms.Page
'   controls. Centralizing the create-or-reuse logic avoids repeated boilerplate
'   across settings-panel build routines
'
' INPUTS
'   ParentPage
'     Page that owns the ComboBox
'
'   ControlName
'     Name of the ComboBox control to retrieve or create
'
' RETURNS
'   The requested MSForms.ComboBox control
'
' BEHAVIOR
'   Normalizes the requested control name, reuses an existing MSForms.ComboBox
'   inside the supplied MultiPage page when available, or creates a new ComboBox
'   when missing
'
' ERROR POLICY
'   Raises a descriptive runtime error if ParentPage is missing, ControlName is
'   blank, an existing control with the same name is not a ComboBox, or the
'   ComboBox cannot be created
'
' DEPENDENCIES
'   MSForms.Page
'   MSForms.ComboBox
'
' NOTES
'   This routine creates controls inside the supplied MultiPage page
'
'   This routine does not format the ComboBox. Formatting remains the
'   responsibility of the caller
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Ensure_PageComboBox"

    Dim EffectiveName       As String               'Normalized control name
    
    Dim ExistingControl     As MSForms.control      'Existing page control
    Dim Cbo                 As MSForms.ComboBox     'Target ComboBox control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing parent page
        If ParentPage Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentPage cannot be Nothing"
        End If
    'Normalize the requested control name
        EffectiveName = VBA.Trim$(ControlName)
    'Reject an empty control name
        If VBA.Len(EffectiveName) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ControlName cannot be empty"
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Suppress lookup errors when the control does not exist yet
        On Error Resume Next
    'Try to retrieve an existing control with the requested name
        Set ExistingControl = ParentPage.Controls(EffectiveName)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING COMBOBOX
'------------------------------------------------------------------------------
    'Reuse the existing control when it is available
        If Not ExistingControl Is Nothing Then
            'Reject an existing non-ComboBox control with the same name
                If VBA.TypeName(ExistingControl) <> "ComboBox" Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Control '" & EffectiveName & "' already exists but is not an MSForms.ComboBox"
                End If
            'Use the existing ComboBox
                Set Cbo = ExistingControl
            'Return the existing ComboBox
                Set UF_Ensure_PageComboBox = Cbo
            'Exit after returning the existing ComboBox
                Exit Function
        End If

'------------------------------------------------------------------------------
' CREATE COMBOBOX IF MISSING
'------------------------------------------------------------------------------
    'Create the ComboBox inside the supplied MultiPage page
        Set ExistingControl = ParentPage.Controls.Add("Forms.ComboBox.1", EffectiveName, True)
    'Reject an unexpected created control type
        If VBA.TypeName(ExistingControl) <> "ComboBox" Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Control '" & EffectiveName & "' was created but is not an MSForms.ComboBox"
        End If
    'Use the created ComboBox
        Set Cbo = ExistingControl

'------------------------------------------------------------------------------
' RETURN COMBOBOX
'------------------------------------------------------------------------------
    'Return the created ComboBox
        Set UF_Ensure_PageComboBox = Cbo

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
        Err.Raise Err.Number, PROC_NAME, "Page ComboBox ensure failed: " & Err.Description

End Function

Private Function UF_Ensure_PageCheckBox( _
    ByVal ParentPage As MSForms.Page, _
    ByVal ControlName As String) As MSForms.CheckBox

'
'------------------------------------------------------------------------------
'                           ENSURE PAGE CHECKBOX
'------------------------------------------------------------------------------
' PURPOSE
'   Returns an existing MSForms.CheckBox inside a MultiPage page or creates it if
'   missing
'
' WHY THIS EXISTS
'   Runtime settings pages need CheckBox controls hosted inside MSForms.Page
'   controls. Centralizing the create-or-reuse logic avoids repeated boilerplate
'   across settings-panel build routines
'
' INPUTS
'   ParentPage
'     Page that owns the CheckBox
'
'   ControlName
'     Name of the CheckBox control to retrieve or create
'
' RETURNS
'   The requested MSForms.CheckBox control
'
' BEHAVIOR
'   Normalizes the requested control name, reuses an existing MSForms.CheckBox
'   inside the supplied MultiPage page when available, or creates a new CheckBox
'   when missing
'
' ERROR POLICY
'   Raises a descriptive runtime error if ParentPage is missing, ControlName is
'   blank, an existing control with the same name is not a CheckBox, or the
'   CheckBox cannot be created
'
' DEPENDENCIES
'   MSForms.Page
'   MSForms.CheckBox
'
' NOTES
'   This routine creates controls inside the supplied MultiPage page
'
'   This routine does not format the CheckBox. Formatting remains the
'   responsibility of the caller
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DatePicker.UF_Ensure_PageCheckBox"

    Dim EffectiveName       As String                   'Normalized control name
    
    Dim ExistingControl     As MSForms.control          'Existing page control
    Dim Chk                 As MSForms.CheckBox         'Target CheckBox control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing parent page
        If ParentPage Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentPage cannot be Nothing"
        End If
    'Normalize the requested control name
        EffectiveName = VBA.Trim$(ControlName)
    'Reject an empty control name
        If VBA.Len(EffectiveName) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ControlName cannot be empty"
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Suppress lookup errors when the control does not exist yet
        On Error Resume Next
    'Try to retrieve an existing control with the requested name
        Set ExistingControl = ParentPage.Controls(EffectiveName)
    'Clear any suppressed lookup error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING CHECKBOX
'------------------------------------------------------------------------------
    'Reuse the existing control when it is available
        If Not ExistingControl Is Nothing Then
            'Reject an existing non-CheckBox control with the same name
                If VBA.TypeName(ExistingControl) <> "CheckBox" Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Control '" & EffectiveName & "' already exists but is not an MSForms.CheckBox"
                End If
            'Use the existing CheckBox
                Set Chk = ExistingControl
            'Return the existing CheckBox
                Set UF_Ensure_PageCheckBox = Chk
            'Exit after returning the existing CheckBox
                Exit Function
        End If

'------------------------------------------------------------------------------
' CREATE CHECKBOX IF MISSING
'------------------------------------------------------------------------------
    'Create the CheckBox inside the supplied MultiPage page
        Set ExistingControl = ParentPage.Controls.Add("Forms.CheckBox.1", EffectiveName, True)
    'Reject an unexpected created control type
        If VBA.TypeName(ExistingControl) <> "CheckBox" Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Control '" & EffectiveName & "' was created but is not an MSForms.CheckBox"
        End If
    'Use the created CheckBox
        Set Chk = ExistingControl

'------------------------------------------------------------------------------
' RETURN CHECKBOX
'------------------------------------------------------------------------------
    'Return the created CheckBox
        Set UF_Ensure_PageCheckBox = Chk

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
        Err.Raise Err.Number, PROC_NAME, "Page CheckBox ensure failed: " & Err.Description

End Function

Private Function UF_CalendarGrid_GetWidth() As Single

'
'------------------------------------------------------------------------------
'                           GET CALENDAR GRID WIDTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the rendered width of the seven-column DatePicker day grid
'
' WHY THIS EXISTS
'   Header, footer, divider, picker-panel, and settings-panel layout routines
'   need one shared calculation for the horizontal span occupied by the calendar
'   day grid
'
' INPUTS
'   None
'
' RETURNS
'   Calendar grid width
'
' BEHAVIOR
'   Calculates the total grid width from the day-label width, the horizontal
'   spacing between day labels, and the number of visible day columns
'
' ERROR POLICY
'   Does not raise errors. The calculation depends only on developer-owned layout
'   constants
'
' DEPENDENCIES
'   DP_DAY_LABEL_WIDTH
'   DP_DAY_LABEL_HORIZONTAL_STEP
'   DP_DAY_LABELS_PER_ROW
'
' NOTES
'   Assumes a regular grid with a fixed horizontal step
'
'   Developer-owned layout constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim GridWidth       As Single       'Calculated calendar grid width

'------------------------------------------------------------------------------
' CALCULATE WIDTH
'------------------------------------------------------------------------------
    'Calculate the full width occupied by the day grid
        GridWidth = VBA.CSng( _
            DP_DAY_LABEL_WIDTH + _
            (DP_DAY_LABEL_HORIZONTAL_STEP * (DP_DAY_LABELS_PER_ROW - 1)))

'------------------------------------------------------------------------------
' RETURN WIDTH
'------------------------------------------------------------------------------
    'Return the calculated calendar grid width
        UF_CalendarGrid_GetWidth = GridWidth

End Function

Private Function UF_CalendarGrid_GetBottom() As Single

'
'------------------------------------------------------------------------------
'                           GET CALENDAR GRID BOTTOM
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the bottom coordinate of the DatePicker day grid
'
' WHY THIS EXISTS
'   Footer, divider, picker-panel, and settings-panel layout routines need one
'   shared calculation for the vertical bottom edge of the calendar day grid
'
' INPUTS
'   None
'
' RETURNS
'   Bottom coordinate of the day grid
'
' BEHAVIOR
'   Calculates the bottom coordinate from the day-grid start position, the
'   vertical spacing between rows, the number of visible rows, and the day-label
'   height
'
' ERROR POLICY
'   Does not raise errors. The calculation depends only on developer-owned layout
'   constants
'
' DEPENDENCIES
'   DP_DAY_GRID_START_TOP
'   DP_DAY_LABEL_VERTICAL_STEP
'   DP_DAY_GRID_ROWS
'   DP_DAY_LABEL_HEIGHT
'
' NOTES
'   Assumes a regular calendar grid with a fixed vertical step
'
'   Developer-owned layout constants are intentionally not validated here. They
'   should be checked by a dedicated debug or regression routine
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim GridBottom      As Single       'Calculated calendar grid bottom coordinate

'------------------------------------------------------------------------------
' CALCULATE BOTTOM
'------------------------------------------------------------------------------
    'Calculate the bottom coordinate of the last day-label row
        GridBottom = VBA.CSng( _
            DP_DAY_GRID_START_TOP + _
            (DP_DAY_LABEL_VERTICAL_STEP * (DP_DAY_GRID_ROWS - 1)) + _
            DP_DAY_LABEL_HEIGHT)

'------------------------------------------------------------------------------
' RETURN BOTTOM
'------------------------------------------------------------------------------
    'Return the calculated calendar grid bottom coordinate
        UF_CalendarGrid_GetBottom = GridBottom

End Function

Private Sub UF_Validate_CalendarLayoutConstants(ByVal CallerName As String)

'
'------------------------------------------------------------------------------
'                       VALIDATE CALENDAR LAYOUT CONSTANTS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the shared DatePicker calendar layout constants
'
' WHY THIS EXISTS
'   Calendar layout constants are developer-owned configuration. They should not
'   be repeatedly validated inside high-frequency runtime routines, but a
'   dedicated diagnostic routine is useful when changing the DatePicker visual
'   layout, preparing a release, or running regression checks
'
' INPUTS
'   CallerName
'     Name of the calling routine, used as the propagated error source
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the core weekday-row and day-grid layout constants used to build
'   the DatePicker calendar surface
'
' ERROR POLICY
'   Raises a descriptive runtime error if CallerName is blank or if one or more
'   calendar layout constants are internally inconsistent
'
' DEPENDENCIES
'   DatePicker calendar layout constants in this UserForm module
'
' NOTES
'   This routine is intended for debug / regression usage only
'
'   Do not call this routine from normal UserForm build, populate, hover, or
'   keyboard-navigation paths
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_DatePicker.UF_Validate_CalendarLayoutConstants"

    Dim EffectiveCallerName     As String       'Normalized caller name

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the caller name
        EffectiveCallerName = VBA.Trim$(CallerName)
    'Reject an empty caller name
        If VBA.Len(EffectiveCallerName) = 0 Then
            Err.Raise vbObjectError + 512, PROC_NAME, "CallerName cannot be empty"
        End If

'------------------------------------------------------------------------------
' VALIDATE DAY LABEL DIMENSIONS
'------------------------------------------------------------------------------
    'Reject invalid day-label width
        If DP_DAY_LABEL_WIDTH <= 0 Then
            Err.Raise vbObjectError + 513, EffectiveCallerName, _
                "DP_DAY_LABEL_WIDTH must be greater than zero"
        End If
    'Reject invalid day-label height
        If DP_DAY_LABEL_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 514, EffectiveCallerName, _
                "DP_DAY_LABEL_HEIGHT must be greater than zero"
        End If

'------------------------------------------------------------------------------
' VALIDATE DAY LABEL SPACING
'------------------------------------------------------------------------------
    'Reject invalid day-label horizontal spacing
        If DP_DAY_LABEL_HORIZONTAL_STEP <= 0 Then
            Err.Raise vbObjectError + 515, EffectiveCallerName, _
                "DP_DAY_LABEL_HORIZONTAL_STEP must be greater than zero"
        End If
    'Reject invalid day-label vertical spacing
        If DP_DAY_LABEL_VERTICAL_STEP <= 0 Then
            Err.Raise vbObjectError + 516, EffectiveCallerName, _
                "DP_DAY_LABEL_VERTICAL_STEP must be greater than zero"
        End If

'------------------------------------------------------------------------------
' VALIDATE WEEKDAY LABEL DIMENSIONS
'------------------------------------------------------------------------------
    'Reject invalid weekday-label width
        If DP_DOW_LABEL_WIDTH <= 0 Then
            Err.Raise vbObjectError + 517, EffectiveCallerName, _
                "DP_DOW_LABEL_WIDTH must be greater than zero"
        End If
    'Reject invalid weekday-label height
        If DP_DOW_LABEL_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 518, EffectiveCallerName, _
                "DP_DOW_LABEL_HEIGHT must be greater than zero"
        End If

'------------------------------------------------------------------------------
' VALIDATE WEEKDAY LABEL SPACING
'------------------------------------------------------------------------------
    'Reject invalid weekday-label horizontal spacing
        If DP_DOW_LABEL_HORIZONTAL_STEP <= 0 Then
            Err.Raise vbObjectError + 519, EffectiveCallerName, _
                "DP_DOW_LABEL_HORIZONTAL_STEP must be greater than zero"
        End If

'------------------------------------------------------------------------------
' VALIDATE GRID STRUCTURE
'------------------------------------------------------------------------------
    'Reject invalid day-grid row count
        If DP_DAY_GRID_ROWS <= 0 Then
            Err.Raise vbObjectError + 520, EffectiveCallerName, _
                "DP_DAY_GRID_ROWS must be greater than zero"
        End If
    'Reject invalid day-grid column count
        If DP_DAY_LABELS_PER_ROW <= 0 Then
            Err.Raise vbObjectError + 521, EffectiveCallerName, _
                "DP_DAY_LABELS_PER_ROW must be greater than zero"
        End If
    'Reject invalid total day-label count
        If DP_DAY_LABEL_COUNT <= 0 Then
            Err.Raise vbObjectError + 522, EffectiveCallerName, _
                "DP_DAY_LABEL_COUNT must be greater than zero"
        End If
    'Reject inconsistent total day-label count
        If DP_DAY_GRID_ROWS * DP_DAY_LABELS_PER_ROW <> DP_DAY_LABEL_COUNT Then
            Err.Raise vbObjectError + 523, EffectiveCallerName, _
                "DP_DAY_LABEL_COUNT must equal DP_DAY_GRID_ROWS * DP_DAY_LABELS_PER_ROW"
        End If

'------------------------------------------------------------------------------
' VALIDATE DAY CELL DIMENSIONS
'------------------------------------------------------------------------------
    'Reject invalid day-cell background width
        If DP_DAY_CELL_WIDTH <= 0 Then
            Err.Raise vbObjectError + 524, EffectiveCallerName, _
                "DP_DAY_CELL_WIDTH must be greater than zero"
        End If
    'Reject invalid day-cell background height
        If DP_DAY_CELL_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 525, EffectiveCallerName, _
                "DP_DAY_CELL_HEIGHT must be greater than zero"
        End If

End Sub





