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
'   Provides the runtime UserForm implementation for the DatePicker UI
'
' WHY THIS EXISTS
'   The DatePicker form is built mostly at runtime so that the visual layout,
'   event routing, date-grid rendering, picker-panel behavior, footer shortcuts,
'   and footer clock remain centralized, consistent, and independent from
'   design-time controls
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Manages:
'     - UserForm initialization, activation, formatting, keyboard routing, and
'       termination
'     - runtime creation and reuse of header, weekday, day-grid, picker-panel,
'       divider, footer, and footer-icon controls
'     - month and year navigation through clickable header labels
'     - month and year selection through the reusable picker panel
'     - calendar-grid population across a fixed 6 x 7 layout
'     - first-day-of-week and local-name dependent captions
'     - today, selected-date, keyboard-date, outside-month, disabled, and hover
'       visual states
'     - day-label selection and delegation to the shared DatePicker write-back
'       routine
'     - footer shortcut actions:
'         - Today writes today's date without time
'         - Time writes today's date with the current system time
'     - keyboard shortcuts for navigation, selection, Today, Now, month panel,
'       year panel, and close
'     - live footer clock updates through the shared timer infrastructure
'     - cleanup of timers, runtime event hooks, hover state, keyboard state, and
'       picker-panel state
'
' ERROR POLICY
'   Event procedures and helper routines raise descriptive runtime errors unless
'   they are explicit cleanup paths, where errors are intentionally suppressed
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
'   M_Settings_Load
'   M_Settings_IsValidFirstDayOfWeek
'
'   M_FormBridge_ConsumeInitialDate
'
'   M_Caption_GetMonth
'   M_Caption_GetDate
'
'   M_Window_RemoveTitleBar
'   M_Window_MoveFormToMouse
'   M_Platform_ShouldUseWinAPI
'
'   M_Timer_ApplyClockMode
'   M_Timer_Stop
'
'   M_DatePolicy_CanSelectDate
'   M_Picker_SelectDate
'   DP_Today
'   DP_Now
'   DP_Close
'
'   gDP_HasSelectedDate
'   gDP_SelectedDate
'
'   gDP_FirstDayOfWeek
'   gDP_UseLocalNames
'   gDP_SizeMode
'   gDP_HighlightWeekends
'
' NOTES
'   Runtime-created controls are not available as direct VBA member variables
'   unless they also exist at design time. Use Me.Controls or stored object
'   references returned by UF_Ensure_Label / UF_Ensure_FrameLabel
'
'   The picker panel is intentionally reusable. It displays either months or
'   years depending on mPickerPanelMode
'
'   The picker panel does not contain year-scroll arrows. Year scrolling is
'   handled by the header year-arrow labels when the year panel is visible
'
'   The live clock must update only Lbl_Time.Caption and must not repaint,
'   rebuild, or reformat footer controls
'
'   Settings-dependent refresh updates both captions and the calendar grid when
'   the first-day-of-week setting changes
'
'   Date selection is delegated to the shared DatePicker module through
'   M_Picker_SelectDate after the selected date has been validated by
'   M_DatePolicy_CanSelectDate
'
'   The footer layout is intentionally configured by constants. Today is placed
'   on the left because it is the primary date shortcut. Time is placed on the
'   right because it is a secondary date-time shortcut
'
'   This is UserForm code-behind. Paste it into UF_DatePicker; do not import it
'   as a standard module unless you intentionally adapt it
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------

    '--------------------------------FORM--------------------------------------
    Private Const DP_FORM_CAPTION                       As String = "DATETIME PICKER"    'UserForm caption
    Private Const DP_FORM_WIDTH                         As Single = 223                  'UserForm width
    Private Const DP_FORM_HEIGHT                        As Single = 269                  'UserForm height
    Private Const DP_FORM_HEIGHT_COMPACT                As Single = 210                  'Compact form height
    Private Const DP_FORM_TITLEBAR_HEIGHT_COMPENSATION  As Single = 24                   'Additional height when title bar remains visible

    Private Const DP_FORM_BACK_COLOR                    As Long = vbWhite                'UserForm background color
    Private Const DP_FORM_FORE_COLOR                    As Long = vbButtonText           'UserForm foreground color
    Private Const DP_FORM_FONT_NAME                     As String = "Segoe UI"           'UserForm font name
    Private Const DP_FORM_FONT_SIZE                     As Single = 9                    'UserForm font size
    Private Const DP_FORM_STARTUP_POSITION              As Long = 0                      'UserForm startup position

    '-----------------------------KEYBOARD-------------------------------------
    Private Const DP_KEY_CTRL_MASK                      As Integer = 2                   'Ctrl key mask

    '-------------------------------HEADER-------------------------------------
    Private Const DP_HEADER_TOP                         As Single = 0                    'Header banner top position
    Private Const DP_HEADER_HEIGHT                      As Single = 52                   'Header banner height
    Private Const DP_HEADER_BACK_COLOR                  As Long = 9985057                'Header banner background color
    Private Const DP_HEADER_FORE_COLOR                  As Long = vbWhite                'Header banner foreground color

    Private Const DP_HEADER_MONTH_REL_LEFT              As Single = 24                   'Header month label relative left
    Private Const DP_HEADER_MONTH_TOP                   As Single = 8                    'Header month label top
    Private Const DP_HEADER_MONTH_WIDTH                 As Single = 84                   'Header month label width
    Private Const DP_HEADER_MONTH_HEIGHT                As Single = 20                   'Header month label height

    Private Const DP_HEADER_YEAR_REL_LEFT               As Single = 152                  'Header year label relative left
    Private Const DP_HEADER_YEAR_TOP                    As Single = 9                    'Header year label top
    Private Const DP_HEADER_YEAR_WIDTH                  As Single = 42                   'Header year label width
    Private Const DP_HEADER_YEAR_HEIGHT                 As Single = 20                   'Header year label height

    Private Const DP_HEADER_PREVMONTH_ARROW_LEFT        As Single = 4                    'Previous-month arrow relative left
    Private Const DP_HEADER_NEXTMONTH_ARROW_LEFT        As Single = 114                  'Next-month arrow relative left
    Private Const DP_HEADER_MONTH_ARROW_TOP             As Single = 10                   'Header month-arrow top
    Private Const DP_HEADER_MONTH_ARROW_WIDTH           As Single = 14                   'Header month-arrow width
    Private Const DP_HEADER_MONTH_ARROW_HEIGHT          As Single = 18                   'Header month-arrow height

    Private Const DP_HEADER_YEAR_ARROW_LEFT             As Single = 195                  'Year-arrow relative left
    Private Const DP_HEADER_YEAR_ARROW_TOP              As Single = 7                    'Year-arrow top
    Private Const DP_HEADER_YEAR_ARROW_WIDTH            As Single = 11                   'Year-arrow width
    Private Const DP_HEADER_YEAR_ARROW_HEIGHT           As Single = 11                   'Year-arrow height

    Private Const DP_HEADER_HOVER_BACK_COLOR            As Long = &HAE7546               'Header hover background color

    '-----------------------------DAY OF WEEK----------------------------------
    Private Const DP_DOW_LABEL_WIDTH                    As Single = 30                   'Weekday-label width
    Private Const DP_DOW_LABEL_HEIGHT                   As Single = 16                   'Weekday-label height
    Private Const DP_DOW_LABEL_HORIZONTAL_STEP          As Single = 30                   'Weekday-label horizontal spacing

    '--------------------------------DAYS--------------------------------------
    Private Const DP_DAY_LABEL_WIDTH                    As Single = 18                   'Day-label width
    Private Const DP_DAY_LABEL_HEIGHT                   As Single = 15                   'Day-label height
    Private Const DP_DAY_LABEL_HORIZONTAL_STEP          As Single = 30                   'Day-label horizontal spacing
    Private Const DP_DAY_LABEL_VERTICAL_STEP            As Single = 25                   'Day-label vertical spacing
    
    Private Const DP_DAY_CELL_WIDTH                     As Single = 24                   'Day-cell background width
    Private Const DP_DAY_CELL_HEIGHT                    As Single = 22                   'Day-cell background height
    
    Private Const DP_DAY_GRID_START_LEFT                As Single = 10                   'Calendar grid left position
    Private Const DP_DAY_GRID_START_TOP                 As Single = 60                   'First day-label row top position
    Private Const DP_DAY_GRID_ROWS                      As Long = 6                      'Calendar rows

    Private Const DP_DAY_LABEL_COUNT                    As Long = 42                     'Total day labels
    Private Const DP_DAY_LABELS_PER_ROW                 As Long = 7                      'Calendar columns

    Private Const DP_DAY_NORMAL_BACK_COLOR              As Long = vbWhite                'Normal day-label background
    Private Const DP_DAY_NORMAL_FORE_COLOR              As Long = vbButtonText           'Normal day-label foreground
    Private Const DP_DAY_HOVER_BACK_COLOR               As Long = &HF2F2F2               'Hover day-label background
    Private Const DP_DAY_HOVER_BORDER_COLOR             As Long = &HC8C8C8               'Hover day-label border color

    Private Const DP_DAY_OUTSIDE_MONTH_FORE_COLOR       As Long = &H808080               'Outside-month day foreground
    Private Const DP_DAY_CURRENT_MONTH_FORE_COLOR       As Long = vbButtonText           'Current-month day foreground

    Private Const DP_DAY_TODAY_BACK_COLOR               As Long = &HF2F7FC               'Today day-label background
    Private Const DP_DAY_TODAY_BORDER_COLOR             As Long = 9985057                'Today day-label border color
    Private Const DP_DAY_SELECTED_BACK_COLOR            As Long = 9985057                'Selected day-label background
    Private Const DP_DAY_SELECTED_FORE_COLOR            As Long = vbWhite                'Selected day-label foreground

    '-------------------------------DIVIDER------------------------------------
    Private Const DP_DIVIDER_HEIGHT                     As Single = 2                    'Divider height
    Private Const DP_DIVIDER_SIDE_MARGIN                As Single = 0                    'Divider left/right margin
    Private Const DP_DIVIDER_COLOR                      As Long = &HE6E6E6               'Divider border color
    Private Const DP_DIVIDER_BACK_COLOR                 As Long = vbWhite                'Divider background color

    '-------------------------------FOOTER-------------------------------------
    Private Const DP_FOOTER_TOP_GAP                     As Single = 10                   'Gap between day grid and footer
    Private Const DP_FOOTER_HEIGHT                      As Single = 50                   'Footer banner height
    Private Const DP_FOOTER_SIDE_PADDING                As Single = 8                    'Footer banner horizontal padding
    Private Const DP_FOOTER_BACK_COLOR                  As Long = &HF7F7F7               'Footer banner background color
    Private Const DP_FOOTER_FORE_COLOR                  As Long = vbButtonText           'Footer banner foreground color

    '--------------------------FOOTER GROUPS-----------------------------------
    Private Const DP_FOOTER_GROUP_WIDTH                 As Single = 100                  'Footer shortcut group width
    Private Const DP_FOOTER_GROUP_HEIGHT                As Single = 38                   'Footer shortcut group height
    Private Const DP_FOOTER_GROUP_LEFT_REL_LEFT         As Single = 4                    'Left shortcut group relative left
    Private Const DP_FOOTER_GROUP_RIGHT_REL_LEFT        As Single = 94                   'Right shortcut group relative left

    Private Const DP_FOOTER_GROUP_ICON_OFFSET_LEFT      As Single = 11                   'Footer group icon offset left
    Private Const DP_FOOTER_GROUP_TEXT_OFFSET_LEFT      As Single = 22                   'Footer group text offset left
    Private Const DP_FOOTER_GROUP_TEXT_WIDTH            As Single = 76                   'Footer group text width

    '--------------------------FOOTER LEFT: TODAY------------------------------
    Private Const DP_FOOTER_TODAY_ICON_REL_LEFT         As Single = DP_FOOTER_GROUP_LEFT_REL_LEFT + DP_FOOTER_GROUP_ICON_OFFSET_LEFT     'Footer Today icon relative left
    Private Const DP_FOOTER_TODAY_CAPTION_REL_LEFT      As Single = DP_FOOTER_GROUP_LEFT_REL_LEFT + DP_FOOTER_GROUP_TEXT_OFFSET_LEFT     'Footer Today caption relative left
    Private Const DP_FOOTER_TODAY_VALUE_REL_LEFT        As Single = DP_FOOTER_GROUP_LEFT_REL_LEFT + DP_FOOTER_GROUP_TEXT_OFFSET_LEFT     'Footer Today value relative left

    Private Const DP_FOOTER_TODAY_CAPTION_TOP           As Single = 10                   'Footer Today caption top
    Private Const DP_FOOTER_TODAY_CAPTION_WIDTH         As Single = DP_FOOTER_GROUP_TEXT_WIDTH 'Footer Today caption width
    Private Const DP_FOOTER_TODAY_CAPTION_HEIGHT        As Single = 16                   'Footer Today caption height

    Private Const DP_FOOTER_TODAY_VALUE_TOP             As Single = 30                   'Footer Today value top
    Private Const DP_FOOTER_TODAY_VALUE_WIDTH           As Single = DP_FOOTER_GROUP_TEXT_WIDTH 'Footer Today value width
    Private Const DP_FOOTER_TODAY_VALUE_HEIGHT          As Single = 16                   'Footer Today value height

    '---------------------------FOOTER RIGHT: TIME------------------------------
    Private Const DP_FOOTER_TIME_ICON_REL_LEFT          As Single = DP_FOOTER_GROUP_RIGHT_REL_LEFT + DP_FOOTER_GROUP_ICON_OFFSET_LEFT    'Footer Time icon relative left
    Private Const DP_FOOTER_TIME_CAPTION_REL_LEFT       As Single = DP_FOOTER_GROUP_RIGHT_REL_LEFT + DP_FOOTER_GROUP_TEXT_OFFSET_LEFT    'Footer Time caption relative left
    Private Const DP_FOOTER_TIME_VALUE_REL_LEFT         As Single = DP_FOOTER_GROUP_RIGHT_REL_LEFT + DP_FOOTER_GROUP_TEXT_OFFSET_LEFT    'Footer Time value relative left
    
    Private Const DP_FOOTER_TIME_CAPTION_TOP            As Single = 10                   'Footer Time caption top
    Private Const DP_FOOTER_TIME_CAPTION_WIDTH          As Single = 66                   'Footer Time caption width
    Private Const DP_FOOTER_TIME_CAPTION_HEIGHT         As Single = 16                   'Footer Time caption height
    
    Private Const DP_FOOTER_TIME_VALUE_TOP              As Single = 30                   'Footer Time value top
    Private Const DP_FOOTER_TIME_VALUE_WIDTH            As Single = 66                   'Footer Time value width
    Private Const DP_FOOTER_TIME_VALUE_HEIGHT           As Single = 16                   'Footer Time value height
    
    Private Const DP_FOOTER_TIME_HALO_REL_LEFT          As Single = DP_FOOTER_GROUP_RIGHT_REL_LEFT 'Time halo relative left
    Private Const DP_FOOTER_TIME_HALO_WIDTH             As Single = 88                   'Time halo width

    '-----------------------------FOOTER ICONS---------------------------------
    Private Const DP_FOOTER_ICON_FONT_NAME              As String = "Segoe MDL2 Assets"  'Footer icon font name
    Private Const DP_FOOTER_ICON_FORE_COLOR             As Long = &H808000               'Footer icon teal / green color
    Private Const DP_FOOTER_ICON_TOP                    As Single = 17                   'Footer icon top
    Private Const DP_FOOTER_ICON_WIDTH                  As Single = 18                   'Footer icon width
    Private Const DP_FOOTER_ICON_HEIGHT                 As Single = 18                   'Footer icon height
    Private Const DP_FOOTER_ICON_FONT_SIZE              As Single = 14                   'Footer icon font size

    Private Const DP_FOOTER_ICON_CLOCK_CODEPOINT        As Long = &HE121                 'Segoe MDL2 clock glyph
    Private Const DP_FOOTER_ICON_CALENDAR_CODEPOINT     As Long = &HE8BF                 'Segoe MDL2 calendar-day glyph

    '--------------------------FOOTER SETTINGS--------------------------------
    Private Const DP_FOOTER_SEPARATOR_REL_LEFT          As Single = 180                 'Footer settings separator relative left
    Private Const DP_FOOTER_SEPARATOR_TOP               As Single = 9                   'Footer settings separator top
    Private Const DP_FOOTER_SEPARATOR_WIDTH             As Single = 1                   'Footer settings separator width
    Private Const DP_FOOTER_SEPARATOR_HEIGHT            As Single = 32                  'Footer settings separator height
    Private Const DP_FOOTER_SEPARATOR_COLOR             As Long = &HE0E0E0              'Footer settings separator color
    
    Private Const DP_FOOTER_SETTINGS_ICON_REL_LEFT      As Single = 190                 'Footer settings icon relative left
    Private Const DP_FOOTER_SETTINGS_ICON_TOP           As Single = 17                  'Footer settings icon top
    Private Const DP_FOOTER_SETTINGS_ICON_WIDTH         As Single = 18                  'Footer settings icon width
    Private Const DP_FOOTER_SETTINGS_ICON_HEIGHT        As Single = 18                  'Footer settings icon height
    Private Const DP_FOOTER_SETTINGS_ICON_FONT_NAME     As String = "Segoe MDL2 Assets" 'Footer settings icon font name
    Private Const DP_FOOTER_SETTINGS_ICON_FONT_SIZE     As Single = 14                  'Footer settings icon font size
    Private Const DP_FOOTER_SETTINGS_ICON_CODEPOINT     As Long = &HE713                'Segoe MDL2 settings glyph
    Private Const DP_FOOTER_SETTINGS_ICON_CAPTION       As String = "Settings"          'Footer Settings icon tooltip / accessibility caption
    
    '-----------------------------FOOTER HOVER---------------------------------
    Private Const DP_FOOTER_TODAY_HALO_REL_LEFT         As Single = DP_FOOTER_GROUP_LEFT_REL_LEFT 'Today halo relative left
    Private Const DP_FOOTER_TODAY_HALO_WIDTH            As Single = DP_FOOTER_GROUP_WIDTH 'Today halo width
    
    Private Const DP_FOOTER_HALO_TOP                    As Single = 7                    'Footer halo top
    Private Const DP_FOOTER_HALO_HEIGHT                 As Single = DP_FOOTER_GROUP_HEIGHT 'Footer halo height
    Private Const DP_FOOTER_HALO_BACK_COLOR             As Long = DP_HEADER_HOVER_BACK_COLOR 'Footer halo background color
    Private Const DP_FOOTER_HOVER_FORE_COLOR            As Long = vbWhite               'Footer hover foreground color

    Private Const DP_FOOTER_SETTINGS_HALO_REL_LEFT      As Single = 184                 'Settings halo relative left
    Private Const DP_FOOTER_SETTINGS_HALO_WIDTH         As Single = 30                  'Settings halo width
    Private Const DP_FOOTER_SETTINGS_HALO_TOP           As Single = DP_FOOTER_HALO_TOP  'Settings halo top
    Private Const DP_FOOTER_SETTINGS_HALO_HEIGHT        As Single = DP_FOOTER_HALO_HEIGHT 'Settings halo height

    '-----------------------------PICKER PANEL---------------------------------
    Private Const DP_PICKER_PANEL_NAME                  As String = "Fra_PickerPanel"    'Reusable picker panel name
    Private Const DP_PICKER_PANEL_WIDTH                 As Single = 218                  'Picker panel width
    Private Const DP_PICKER_PANEL_HEIGHT                As Single = 154                  'Picker panel height
    Private Const DP_PICKER_PANEL_BACK_COLOR            As Long = vbWhite                'Picker panel background color

    Private Const DP_PICKER_ITEM_COUNT                  As Long = 12                     'Picker item count
    Private Const DP_PICKER_ITEMS_PER_ROW               As Long = 3                      'Picker items per row
    Private Const DP_PICKER_ITEM_TOP                    As Single = 9                    'Picker item first top
    Private Const DP_PICKER_ITEM_WIDTH                  As Single = 58                   'Picker item width
    Private Const DP_PICKER_ITEM_HEIGHT                 As Single = 28                   'Picker item height
    Private Const DP_PICKER_ITEM_HORIZONTAL_STEP        As Single = 66                   'Picker item horizontal step
    Private Const DP_PICKER_ITEM_VERTICAL_STEP          As Single = 34                   'Picker item vertical step
    Private Const DP_PICKER_ITEM_TEXT_TOP_OFFSET        As Single = 9                    'Picker item text vertical offset
    Private Const DP_PICKER_ITEM_TEXT_HEIGHT            As Single = 16                   'Picker item text label height
    Private Const DP_PICKER_ITEM_BACK_COLOR             As Long = vbWhite                'Picker item background color
    Private Const DP_PICKER_ITEM_BORDER_COLOR           As Long = &HE6E6E6               'Picker item border color
    Private Const DP_PICKER_ITEM_FORE_COLOR             As Long = vbButtonText           'Picker item text color
    Private Const DP_PICKER_ITEM_HOVER_BACK_COLOR       As Long = &HF2F2F2               'Picker item hover background color
    Private Const DP_PICKER_ITEM_HOVER_BORDER_COLOR     As Long = &HC8C8C8               'Picker item hover border color

    '----------------------------SETTINGS PANEL-------------------------------
    Private Const DP_SETTINGS_PANEL_NAME                As String = "Fra_Settings"       'Reusable settings panel name
    Private Const DP_SETTINGS_PANEL_BACK_COLOR          As Long = vbWhite                'Settings panel background color

    Private Const DP_SETTINGS_TITLE_CAPTION             As String = "SETTINGS"           'Settings panel title caption
    Private Const DP_SETTINGS_TITLE_TOP                 As Single = 10                   'Settings title top
    Private Const DP_SETTINGS_TITLE_LEFT                As Single = 0                    'Settings title left
    Private Const DP_SETTINGS_TITLE_WIDTH               As Single = DP_PICKER_PANEL_WIDTH 'Settings title width
    Private Const DP_SETTINGS_TITLE_HEIGHT              As Single = 18                   'Settings title height

    Private Const DP_SETTINGS_OPTION_LEFT               As Single = 18                   'Settings option label left
    Private Const DP_SETTINGS_OPTION_WIDTH              As Single = 182                  'Settings option label width
    Private Const DP_SETTINGS_OPTION_HEIGHT             As Single = 17                   'Settings option label height
    Private Const DP_SETTINGS_OPTION_TOP_1              As Single = 39                   'Settings option 1 top
    Private Const DP_SETTINGS_OPTION_TOP_2              As Single = 61                   'Settings option 2 top
    Private Const DP_SETTINGS_OPTION_TOP_3              As Single = 83                   'Settings option 3 top
    Private Const DP_SETTINGS_OPTION_TOP_4              As Single = 105                  'Settings option 4 top

    Private Const DP_SETTINGS_HINT_TOP                  As Single = 130                  'Settings hint top
    Private Const DP_SETTINGS_HINT_LEFT                 As Single = 12                   'Settings hint left
    Private Const DP_SETTINGS_HINT_WIDTH                As Single = 194                  'Settings hint width
    Private Const DP_SETTINGS_HINT_HEIGHT               As Single = 14                   'Settings hint height
    Private Const DP_SETTINGS_HINT_FORE_COLOR           As Long = &H808080               'Settings hint foreground color

    '--------------------------SETTINGS HEADER ICONS----------------------------
    Private Const DP_SETTINGS_HEADER_ICON_FONT_NAME     As String = "Segoe MDL2 Assets" 'Settings header icon font name
    Private Const DP_SETTINGS_HEADER_ICON_HEIGHT        As Single = 16                  'Settings header icon height
    Private Const DP_SETTINGS_HEADER_ICON_TOP           As Single = 7                   'Settings header icon top
    Private Const DP_SETTINGS_HEADER_ICON_WIDTH         As Single = 16                  'Settings header icon width
    Private Const DP_SETTINGS_HEADER_ICON_FORE_COLOR    As Long = &H808080              'Settings header icon color

    Private Const DP_SETTINGS_SAVE_FONT_SIZE            As Single = 11                  'Save icon font size
    Private Const DP_SETTINGS_SAVE_LEFT                 As Single = 176                 'Settings save icon left
    Private Const DP_SETTINGS_SAVE_WIDTH                As Single = 18                  'Save icon width
    Private Const DP_SETTINGS_SAVE_HEIGHT               As Single = 16                  'Save icon height
    Private Const DP_SETTINGS_SAVE_CODEPOINT            As Long = &HE74E                'Segoe MDL2 save glyph
    Private Const DP_SETTINGS_SAVE_TOOLTIP              As String = "Save settings"     'Settings save tooltip
   
    Private Const DP_SETTINGS_CLOSE_FONT_SIZE           As Single = 9                   'Close icon font size
    Private Const DP_SETTINGS_CLOSE_LEFT                As Single = 194                 'Settings close icon left
    Private Const DP_SETTINGS_CLOSE_CODEPOINT           As Long = &HE8BB                'Segoe MDL2 close glyph
    Private Const DP_SETTINGS_CLOSE_TOOLTIP             As String = "Close settings"    'Settings close tooltip
    

    '--------------------------SETTINGS MULTIPAGE-------------------------------
    Private Const DP_SETTINGS_MULTIPAGE_NAME            As String = "Mp_Settings"        'Settings MultiPage name
    Private Const DP_SETTINGS_MULTIPAGE_LEFT            As Single = 7                    'Settings MultiPage left
    Private Const DP_SETTINGS_MULTIPAGE_TOP             As Single = 50                   'Settings MultiPage top below fake tabs
    Private Const DP_SETTINGS_MULTIPAGE_WIDTH           As Single = 204                  'Settings MultiPage width
    Private Const DP_SETTINGS_MULTIPAGE_HEIGHT          As Single = 96                   'Settings MultiPage height

    Private Const DP_SETTINGS_PAGE_BACK_LEFT            As Single = 0                    'Settings page background left
    Private Const DP_SETTINGS_PAGE_BACK_TOP             As Single = 0                    'Settings page background top
    Private Const DP_SETTINGS_PAGE_BACK_WIDTH           As Single = 206                  'Settings page background width
    Private Const DP_SETTINGS_PAGE_BACK_HEIGHT          As Single = 96                   'Settings page background height
    Private Const DP_SETTINGS_PAGE_BACK_COLOR           As Long = vbWhite                'Settings page background color

    Private Const DP_SETTINGS_PAGE_DISPLAY_CAPTION      As String = "Display Settings"   'Display settings page caption
    Private Const DP_SETTINGS_PAGE_BEHAVIOR_CAPTION     As String = "Behavior Settings"  'Behavior settings page caption
    Private Const DP_SETTINGS_PAGE_INTEGRATION_CAPTION  As String = "Integration Settings" 'Integration settings page caption

    '--------------------------SETTINGS FAKE TABS-------------------------------
    Private Const DP_SETTINGS_TAB_DISPLAY_NAME          As String = "Lbl_SettingsTabDisplay" 'Display tab label name
    Private Const DP_SETTINGS_TAB_BEHAVIOR_NAME         As String = "Lbl_SettingsTabBehavior" 'Behavior tab label name
    Private Const DP_SETTINGS_TAB_INTEGRATION_NAME      As String = "Lbl_SettingsTabIntegration" 'Integration tab label name

    Private Const DP_SETTINGS_TAB_LEFT                  As Single = 7                    'Settings tab strip left
    Private Const DP_SETTINGS_TAB_TOP                   As Single = 32                   'Settings tab strip top
    Private Const DP_SETTINGS_TAB_HEIGHT                As Single = 18                   'Settings tab height
    Private Const DP_SETTINGS_TAB_GAP                   As Single = 1                    'Settings tab gap

    Private Const DP_SETTINGS_TAB_DISPLAY_WIDTH         As Single = 60                   'Display tab width
    Private Const DP_SETTINGS_TAB_BEHAVIOR_WIDTH        As Single = 66                   'Behavior tab width
    Private Const DP_SETTINGS_TAB_INTEGRATION_WIDTH     As Single = 76                   'Integration tab width

    Private Const DP_SETTINGS_TAB_DISPLAY_CAPTION       As String = "Display"            'Display tab caption
    Private Const DP_SETTINGS_TAB_BEHAVIOR_CAPTION      As String = "Behavior"           'Behavior tab caption
    Private Const DP_SETTINGS_TAB_INTEGRATION_CAPTION   As String = "Integration"        'Integration tab caption

    Private Const DP_SETTINGS_TAB_NORMAL_BACK_COLOR     As Long = &HF2F2F2               'Inactive tab background color
    Private Const DP_SETTINGS_TAB_SELECTED_BACK_COLOR   As Long = vbWhite                'Active tab background color
    Private Const DP_SETTINGS_TAB_BORDER_COLOR          As Long = &HD9D9D9               'Tab border color
    Private Const DP_SETTINGS_TAB_NORMAL_FORE_COLOR     As Long = vbButtonText           'Inactive tab text color
    Private Const DP_SETTINGS_TAB_SELECTED_FORE_COLOR   As Long = DP_HEADER_BACK_COLOR   'Active tab text color

    '----------------------SETTINGS DISPLAY PAGE CONTROLS-----------------------
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
    Private Const DP_SETTINGS_CHECKBOX_HEIGHT           As Single = 14                   'Settings CheckBox height
    Private Const DP_SETTINGS_CHECKBOX_FONT_SIZE        As Single = 8.25                 'Settings CheckBox font size
    Private Const DP_SETTINGS_CHECKBOX_TOP_1            As Single = 24                   'Settings first CheckBox top
    Private Const DP_SETTINGS_CHECKBOX_VERTICAL_STEP    As Single = 17                   'Settings CheckBox vertical step

    Private Const DP_SETTINGS_CHECK_LOCAL_NAMES_CAPTION As String = "Use local names"       'Local names CheckBox caption
    Private Const DP_SETTINGS_CHECK_COMPACT_CAPTION     As String = "Compact layout"        'Compact layout CheckBox caption
    Private Const DP_SETTINGS_CHECK_WEEKENDS_CAPTION    As String = "Highlight weekends"    'Weekend highlight CheckBox caption
    Private Const DP_SETTINGS_CHECK_LIVE_CLOCK_CAPTION  As String = "Live clock"            'Live clock checkbox caption
    
    Private Const DP_SETTINGS_PLACEHOLDER_LEFT          As Single = 10                      'Placeholder label left
    Private Const DP_SETTINGS_PLACEHOLDER_TOP           As Single = 18                      'Placeholder label top
    Private Const DP_SETTINGS_PLACEHOLDER_WIDTH         As Single = 176                     'Placeholder label width
    Private Const DP_SETTINGS_PLACEHOLDER_HEIGHT        As Single = 36                      'Placeholder label height

    '--------------------------SETTINGS BEHAVIOR PAGE---------------------------
    Private Const DP_SETTINGS_CHECK_ALLOW_OUTSIDE_CAPTION   As String = "Allow outside-month selection" 'Outside-month selection caption
    Private Const DP_SETTINGS_CHECK_CLOSE_AFTER_CAPTION     As String = "Close after selection"         'Close-after-selection caption

    '-------------------------SETTINGS INTEGRATION PAGE-------------------------
    Private Const DP_SETTINGS_CHECK_RIGHT_CLICK_CAPTION     As String = "Right-click menu"              'Right-click menu caption
    Private Const DP_SETTINGS_CHECK_IN_GRID_ICON_CAPTION    As String = "In-grid icon"                  'In-grid icon caption
    Private Const DP_SETTINGS_CHECK_WINAPI_STYLE_CAPTION    As String = "Use WinAPI styling"            'WinAPI styling caption
    
    '-----------------------------DATE BOUNDS----------------------------------
    Private Const DP_MIN_YEAR                           As Long = 100                    'Minimum supported year
    Private Const DP_MAX_YEAR                           As Long = 9999                   'Maximum supported year
    Private Const DP_YEAR_PANEL_MAX_START               As Long = 9988                   'Maximum 12-year panel start

    '-----------------------------POSITIONING----------------------------------
    Private Const DP_FORM_MOUSE_OFFSET_XPX              As Long = 10                     'Mouse-position X offset in pixels
    Private Const DP_FORM_MOUSE_OFFSET_YPX              As Long = 0                      'Mouse-position Y offset in pixels
    Private Const DP_FORM_CENTER_ON_MOUSE               As Boolean = False               'True to center form on mouse

'------------------------------------------------------------------------------
' PRIVATE VARIABLES
'------------------------------------------------------------------------------
    Private mDayLabelHooks                              As Collection                   'Runtime day-label event hooks
    Private mHeaderLabelHooks                           As Collection                   'Header label click hooks
    Private mPickerPanelHooks                           As Collection                   'Picker panel click hooks
    Private mFooterLabelHooks                           As Collection                   'Footer label click hooks
    Private mSettingsPanelHooks                         As Collection                   'Settings panel click hooks
    
    Private mHoveredDayLabelName                        As String                       'Currently hovered day label name
    Private mHoveredHeaderLabelName                     As String                       'Currently hovered clickable header label name
    Private mHoveredPickerItemIndex                     As Long                         'Currently hovered picker-panel item index
    Private mHoveredFooterActionName                    As String                       'Currently hovered footer action name
    Private mHoveredDayCellIndex                        As Long                         'Currently hovered day-cell index
    
    Private mDayTextLabels(1 To DP_DAY_LABEL_COUNT)     As MSForms.Label                'Cached day text labels
    Private mDayBackLabels(1 To DP_DAY_LABEL_COUNT)     As MSForms.Label                'Cached day background labels
    Private mPickerTextLabels(1 To DP_PICKER_ITEM_COUNT) As MSForms.Label               'Cached picker-panel text labels
    Private mPickerBackLabels(1 To DP_PICKER_ITEM_COUNT) As MSForms.Label               'Cached picker-panel background labels
    Private mDayCellDates(1 To DP_DAY_LABEL_COUNT)      As Date                         'Cached date represented by each day cell
    Private mDayCellHasDate(1 To DP_DAY_LABEL_COUNT)    As Boolean                      'True when cached day-cell date is available
    
    Private mDisplayYear                                As Long                         'Currently displayed year
    Private mDisplayMonth                               As Long                         'Currently displayed month

    Private mPickerPanelMode                            As Long                         '1 = months, 2 = years
    Private mYearPanelStart                             As Long                         'First year shown in year panel
    Private mHasActivated                               As Boolean                      'Tracks first activation to run one-time post-show initialization

    Private mKeyboardDate                               As Date                         'Date currently selected by keyboard navigation
    Private mHasKeyboardDate                            As Boolean                      'True when keyboard date is initialized
    
    Private mDayFontNormal                              As Object                       'Cached normal day-label font
    Private mDayFontWeekend                             As Object                       'Cached weekend day-label font

'------------------------------------------------------------------------------
' FORM LIFECYCLE
'------------------------------------------------------------------------------

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
'   must load saved settings, format the shell, initialize the displayed
'   month/year, create runtime controls, populate the visible calendar grid, and
'   apply the borderless window styling
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Loads saved DatePicker settings
'   - Formats the UserForm shell
'   - Initializes the displayed month and year
'   - Initializes keyboard navigation state
'   - Creates runtime header, weekday, day-grid, picker-panel, and footer UI
'   - Attempts title-bar removal during initialization
'
' ERROR POLICY
'   Raises a descriptive runtime error to the caller if any initialization step
'   fails
'
' DEPENDENCIES
'   M_Settings_Load
'   UF_Form_Format
'   M_FormBridge_ConsumeInitialDate
'   UF_DisplayPeriod_Initialize
'   UF_KeyboardDate_Initialize
'   UF_Header_Build
'   UF_WeekdayRow_Build
'   UF_DayGrid_Build
'   UF_DayGrid_Populate
'   UF_PickerPanel_Build
'   UF_Footer_Build
'   M_Window_RemoveTitleBar
'
' NOTES
'   The display period must be initialized before header labels are created
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InitialDate     As Date     'Initial date consumed from M_DATEPICKER

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Reset the one-time activation guard
        mHasActivated = False
    'Use the configured startup position before the first paint
        Me.StartUpPosition = DP_FORM_STARTUP_POSITION
    'Load saved DatePicker settings
        M_Settings_Load

'------------------------------------------------------------------------------
' FORMAT USERFORM
'------------------------------------------------------------------------------
    'Format the existing DatePicker UserForm instance
        UF_Form_Format

'------------------------------------------------------------------------------
' RESOLVE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Resolve the initial display date from the companion module
        If M_FormBridge_ConsumeInitialDate(InitialDate) Then
            UF_DisplayPeriod_Initialize InitialDate
        Else
            InitialDate = VBA.Date
            UF_DisplayPeriod_Initialize InitialDate
        End If
    'Initialize keyboard navigation date
        UF_KeyboardDate_Initialize InitialDate

'------------------------------------------------------------------------------
' CREATE RUNTIME UI
'------------------------------------------------------------------------------
    'Create the header banner and its labels
        UF_Header_Build
    'Create the locale-dependent / fixed day-of-week header row
        UF_WeekdayRow_Build
    'Create the 6 x 7 day-label grid
        UF_DayGrid_Build
    'Create the hidden month/year picker panel
        UF_PickerPanel_Build
    'Create the hidden settings panel
        UF_SettingsPanel_Build
    'Create the footer banner and its labels
        UF_Footer_Build
    'Populate the day grid for the initialized display period
        UF_DayGrid_Populate mDisplayYear, mDisplayMonth
        
'------------------------------------------------------------------------------
' REMOVE TITLE BAR
'------------------------------------------------------------------------------
    'Attempt to remove the UserForm title bar during initialization
        M_Window_RemoveTitleBar Me

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
        Err.Raise Err.Number, "UF_DatePicker.UserForm_Initialize", Err.Description

End Sub

Private Sub UserForm_Activate()

'
'==============================================================================
'                           ACTIVATE USERFORM
'------------------------------------------------------------------------------
' PURPOSE
'   Runs one-time post-show initialization for the DatePicker UserForm
'
' WHY THIS EXISTS
'   Some UI operations are more reliable after the UserForm window has been
'   created and activated. The activation event provides a fallback point for
'   title-bar removal, mouse-based positioning, and runtime behavior that depends
'   on the displayed form
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits immediately if post-show initialization has already run
'   - Retries title-bar removal after the form window exists
'   - Moves the form so its top-left window position corresponds to the mouse
'     position
'   - Applies the configured static/live clock mode
'   - Marks post-show initialization as complete
'
' ERROR POLICY
'   Raises a descriptive runtime error if post-show initialization fails
'
' DEPENDENCIES
'   M_Window_RemoveTitleBar
'   M_Window_MoveFormToMouse
'   M_Timer_ApplyClockMode
'
' NOTES
'   UserForm_Activate can fire more than once during a form lifetime. The
'   mHasActivated guard prevents repeated positioning, timer initialization, and
'   post-show work
'
'   Exact mouse-position placement requires WinAPI support. If WinAPI features
'   are disabled or unavailable, M_Window_MoveFormToMouse exits safely
'
' UPDATED
'   2026-04-28
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Exit if one-time post-show initialization has already run
        If mHasActivated Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' POST-SHOW INITIALIZATION
'------------------------------------------------------------------------------
    'Retry title-bar removal now that the form window should exist
        M_Window_RemoveTitleBar Me

    'Move the form top-left corner to the current mouse position
        M_Window_MoveFormToMouse _
            Me, _
            DP_FORM_MOUSE_OFFSET_XPX, _
            DP_FORM_MOUSE_OFFSET_YPX, _
            DP_FORM_CENTER_ON_MOUSE

    'Apply the configured static or live clock mode
        M_Timer_ApplyClockMode

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
        Err.Raise Err.Number, "UF_DatePicker.UserForm_Activate", Err.Description

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
'   Users should be able to navigate and select dates without using the mouse
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
'   Supports:
'     - Arrow keys for day and week movement
'     - PageUp / PageDown for month movement
'     - Ctrl + PageUp / Ctrl + PageDown for year movement
'     - Home / End for first / last day of displayed month
'     - Enter / Space to write the keyboard-selected date
'     - Esc to close the form
'     - M to open the month picker panel
'     - Y to open the year picker panel
'     - T to write Today without time
'     - N to write Now with time
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
'   DatePicker selection remains date-only. The N shortcut is a direct Now
'   command and does not change normal calendar-selection behavior
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UserForm_KeyDown"            'Current procedure name

    Dim CtrlIsPressed          As Boolean                                'True when Ctrl modifier is pressed

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Resolve Ctrl key state
        CtrlIsPressed = ((Shift And DP_KEY_CTRL_MASK) <> 0)

'------------------------------------------------------------------------------
' ROUTE KEYBOARD ACTION
'------------------------------------------------------------------------------
    'Route the pressed key
        Select Case KeyCode

            Case vbKeyEscape
                'Close the DatePicker
                    DP_Close
                'Consume the key
                    KeyCode = 0

            Case vbKeyReturn, vbKeySpace
                'Select the keyboard-highlighted date when available and selectable
                    If mHasKeyboardDate Then
                        If M_DatePolicy_CanSelectDate(mKeyboardDate, mDisplayYear, mDisplayMonth) Then
                            M_Picker_SelectDate mKeyboardDate
                        End If
                    End If
                'Consume the key
                    KeyCode = 0

            Case vbKeyLeft
                'Move one day backward
                    UF_KeyboardDate_MoveDays -1
                'Consume the key
                    KeyCode = 0

            Case vbKeyRight
                'Move one day forward
                    UF_KeyboardDate_MoveDays 1
                'Consume the key
                    KeyCode = 0

            Case vbKeyUp
                'Move one week backward
                    UF_KeyboardDate_MoveDays -7
                'Consume the key
                    KeyCode = 0

            Case vbKeyDown
                'Move one week forward
                    UF_KeyboardDate_MoveDays 7
                'Consume the key
                    KeyCode = 0

            Case vbKeyPageUp
                'Move one year or one month backward
                    If CtrlIsPressed Then
                        UF_KeyboardDate_MoveMonths -12
                    Else
                        UF_KeyboardDate_MoveMonths -1
                    End If
                'Consume the key
                    KeyCode = 0

            Case vbKeyPageDown
                'Move one year or one month forward
                    If CtrlIsPressed Then
                        UF_KeyboardDate_MoveMonths 12
                    Else
                        UF_KeyboardDate_MoveMonths 1
                    End If
                'Consume the key
                    KeyCode = 0

            Case vbKeyHome
                'Move to the first day of the displayed month
                    UF_KeyboardDate_Set VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
                'Consume the key
                    KeyCode = 0

            Case vbKeyEnd
                'Move to the last day of the displayed month
                    UF_KeyboardDate_Set VBA.DateSerial(mDisplayYear, mDisplayMonth + 1, 0)
                'Consume the key
                    KeyCode = 0

            Case vbKeyM
                'Show the month picker panel
                    UF_PickerPanel_ShowMonths
                'Consume the key
                    KeyCode = 0

            Case vbKeyY
                'Show the year picker panel
                    UF_PickerPanel_ShowYears
                'Consume the key
                    KeyCode = 0
            
            Case vbKeyT
                'Write today without time
                    DP_Today
                'Consume the key
                    KeyCode = 0

            Case vbKeyN
                'Write today with the current system time
                    DP_Now
                'Consume the key
                    KeyCode = 0

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
'   Resets only the hover areas that currently have active hover state
'
' ERROR POLICY
'   Propagates unexpected runtime errors raised by the hover reset routines
'
' DEPENDENCIES
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_PickerPanel_HoverReset
'   UF_Footer_HoverReset
'
' NOTES
'   Guarding the reset calls avoids repeated no-op reset work while the mouse is
'   moving across the UserForm background
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESET DAY-CELL HOVER STATE
'------------------------------------------------------------------------------
    'Reset day-cell hover only when a day cell is currently hovered
        If mHoveredDayCellIndex <> 0 Then
            UF_DayCell_HoverReset
        End If

'------------------------------------------------------------------------------
' RESET HEADER HOVER STATE
'------------------------------------------------------------------------------
    'Reset header hover only when a header label is currently hovered
        If Len(mHoveredHeaderLabelName) <> 0 Then
            UF_Header_HoverReset
        End If

'------------------------------------------------------------------------------
' RESET PICKER-PANEL HOVER STATE
'------------------------------------------------------------------------------
    'Reset picker-panel hover only when a picker-panel item is currently hovered
        If mHoveredPickerItemIndex <> 0 Then
            UF_PickerPanel_HoverReset
        End If

'------------------------------------------------------------------------------
' RESET FOOTER HOVER STATE
'------------------------------------------------------------------------------
    'Reset footer hover only when a footer action is currently hovered
        If Len(mHoveredFooterActionName) <> 0 Then
            UF_Footer_HoverReset
        End If


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
'   Runtime label event hooks are stored in collections to keep their events
'   alive while the form is open. Those references should be released when the
'   form is destroyed
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stops timers, releases hook collections, and clears transient state
'
' ERROR POLICY
'   Suppresses cleanup errors because the form is already terminating
'
' DEPENDENCIES
'   M_Timer_Stop
'
' NOTES
'   Visual hover reset is not required during termination because the UserForm is
'   being destroyed
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' STOP TIMER
'------------------------------------------------------------------------------
    'Stop the live clock timer if active
        M_Timer_Stop

'------------------------------------------------------------------------------
' RELEASE HOOKS
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
    'Release cached normal day font
        Set mDayFontNormal = Nothing
    'Release cached weekend day font
        Set mDayFontWeekend = Nothing

'------------------------------------------------------------------------------
' CLEAR HOVER STATE
'------------------------------------------------------------------------------
    'Clear the current day-label hover state
        mHoveredDayLabelName = vbNullString
    'Clear the current header-label hover state
        mHoveredHeaderLabelName = vbNullString
    'Clear the current picker-panel hover state
        mHoveredPickerItemIndex = 0
    'Clear the current footer hover state
        mHoveredFooterActionName = vbNullString
    'Clear the current day-cell hover index
        mHoveredDayCellIndex = 0

'------------------------------------------------------------------------------
' CLEAR PANEL AND ACTIVATION STATE
'------------------------------------------------------------------------------
    'Clear the activation guard
        mHasActivated = False
    'Clear the picker-panel mode
        mPickerPanelMode = 0
    'Clear the year-panel start
        mYearPanelStart = 0

'------------------------------------------------------------------------------
' CLEAR KEYBOARD STATE
'------------------------------------------------------------------------------
    'Clear keyboard navigation state
        mKeyboardDate = 0
    'Clear keyboard navigation availability
        mHasKeyboardDate = False

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' USERFORM SHELL
'------------------------------------------------------------------------------

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
'   controls are created or refreshed
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Formats the current UserForm instance using module-level form constants and
'   applies effective height according to size mode and title-bar availability
'
' ERROR POLICY
'   Raises a descriptive runtime error if invalid layout or font constants are
'   configured
'
' DEPENDENCIES
'   M_Platform_ShouldUseWinAPI
'   gDP_SizeMode
'
' NOTES
'   Runtime labels inherit the UserForm font if created after this routine
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Form_Format"                'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE FORM CONSTANTS
'------------------------------------------------------------------------------
    'Reject an empty form caption
        If Len(Trim$(DP_FORM_CAPTION)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "DP_FORM_CAPTION cannot be empty."
        End If
    'Reject invalid form width
        If DP_FORM_WIDTH <= 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "DP_FORM_WIDTH must be greater than zero."
        End If
    'Reject invalid form height
        If DP_FORM_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 515, PROC_NAME, "DP_FORM_HEIGHT must be greater than zero."
        End If
    'Reject invalid compact form height
        If DP_FORM_HEIGHT_COMPACT <= 0 Then
            Err.Raise vbObjectError + 516, PROC_NAME, "DP_FORM_HEIGHT_COMPACT must be greater than zero."
        End If
    'Reject compact height greater than normal height
        If DP_FORM_HEIGHT_COMPACT > DP_FORM_HEIGHT Then
            Err.Raise vbObjectError + 517, PROC_NAME, "DP_FORM_HEIGHT_COMPACT cannot be greater than DP_FORM_HEIGHT."
        End If
    'Reject invalid title-bar height compensation
        If DP_FORM_TITLEBAR_HEIGHT_COMPENSATION < 0 Then
            Err.Raise vbObjectError + 518, PROC_NAME, "DP_FORM_TITLEBAR_HEIGHT_COMPENSATION cannot be negative."
        End If
    'Reject an empty form font name
        If Len(Trim$(DP_FORM_FONT_NAME)) = 0 Then
            Err.Raise vbObjectError + 519, PROC_NAME, "DP_FORM_FONT_NAME cannot be empty."
        End If
    'Reject invalid form font size
        If DP_FORM_FONT_SIZE <= 0 Then
            Err.Raise vbObjectError + 520, PROC_NAME, "DP_FORM_FONT_SIZE must be greater than zero."
        End If
    'Reject invalid startup position
        If DP_FORM_STARTUP_POSITION < 0 Or DP_FORM_STARTUP_POSITION > 3 Then
            Err.Raise vbObjectError + 521, PROC_NAME, "DP_FORM_STARTUP_POSITION must be between 0 and 3."
        End If

'------------------------------------------------------------------------------
' FORMAT USERFORM SHELL
'------------------------------------------------------------------------------
    'Apply the standard DatePicker shell properties
        With Me
            .Caption = DP_FORM_CAPTION
            .Width = DP_FORM_WIDTH
            .Height = DP_FORM_HEIGHT
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
    'Apply compact height when compact mode is enabled
        If gDP_SizeMode = DP_SizeMode_Compact Then
            Me.Width = DP_FORM_WIDTH
            Me.Height = DP_FORM_HEIGHT_COMPACT
        Else
            Me.Width = DP_FORM_WIDTH
            Me.Height = DP_FORM_HEIGHT
        End If

    'Increase the form height when the native title bar remains visible
        If Not M_Platform_ShouldUseWinAPI Then
            Me.Height = Me.Height + DP_FORM_TITLEBAR_HEIGHT_COMPENSATION
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

'------------------------------------------------------------------------------
' HEADER
'------------------------------------------------------------------------------

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
'   current month and year
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds or refreshes the header banner and clickable header labels
'
' ERROR POLICY
'   Propagates unexpected runtime errors raised by the called routines
'
' DEPENDENCIES
'   UF_Header_BuildBanner
'   UF_Header_BuildLabels
'
' NOTES
'   This routine coordinates header creation only
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

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
'   distinct top section
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses Lbl_HeaderBanner and formats it as a full-width top band
'
' ERROR POLICY
'   Raises a descriptive runtime error if the banner cannot be created or
'   formatted
'
' DEPENDENCIES
'   UF_Ensure_Label
'
' NOTES
'   The banner is sized to DP_FORM_WIDTH
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Header_BuildBanner"        'Current procedure name

    Dim Lbl_HeaderBanner       As MSForms.Label                          'Header banner control

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
'   navigation controls used to move across months and years or open the
'   reusable picker panel
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
'
'   Formats each label and registers its click hook in mHeaderLabelHooks
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
'
' NOTES
'   mHeaderLabelHooks is reset each time the header labels are refreshed so that
'   runtime label events remain connected to the current controls
'
'   Header label click actions are routed through UF_PickerPanel_HandleAction by
'   cDatePickerLabelHook
'
'   Runtime MSForms labels may retain old font states when reused. This routine
'   therefore assigns clean StdFont objects instead of changing only selected
'   Font properties
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Header_BuildLabels"        'Current procedure name

    Dim LabelHook              As cDatePickerLabelHook                   'Runtime click hook
    Dim HeaderFont             As Object                                 'Clean header-label font object
    Dim BannerLeft             As Single                                 'Header banner left position
    Dim Lbl_HeaderMonth        As MSForms.Label                          'Header month label
    Dim Lbl_HeaderYear         As MSForms.Label                          'Header year label
    Dim Lbl_PrevMonth          As MSForms.Label                          'Header previous-month label
    Dim Lbl_NextMonth          As MSForms.Label                          'Header next-month label
    Dim Lbl_PrevYear           As MSForms.Label                          'Header previous-year label
    Dim Lbl_NextYear           As MSForms.Label                          'Header next-year label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Reset header label click hooks
        Set mHeaderLabelHooks = New Collection

'------------------------------------------------------------------------------
' ENSURE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Initialize the display period if it is not already available
        If mDisplayYear = 0 Or mDisplayMonth = 0 Then
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
        End With

    'Create a clean header-label font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the month label font explicitly
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

    'Register the month label click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the month label to the month-panel action
        LabelHook.Initialize Me, Lbl_HeaderMonth, "SHOW_MONTH_PANEL", "HEADER"

    'Store the hook so that the click event remains alive
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
        End With

    'Create a clean header-label font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the year label font explicitly
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

    'Register the year label click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the year label to the year-panel action
        LabelHook.Initialize Me, Lbl_HeaderYear, "SHOW_YEAR_PANEL", "HEADER"

    'Store the hook so that the click event remains alive
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

    'Create a clean header-label font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the previous-year label font explicitly
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

    'Register the previous-year click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the previous-year label to the previous-year action
        LabelHook.Initialize Me, Lbl_PrevYear, "PREV_YEAR", "HEADER"

    'Store the hook so that the click event remains alive
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

    'Create a clean header-label font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the next-year label font explicitly
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

    'Register the next-year click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the next-year label to the next-year action
        LabelHook.Initialize Me, Lbl_NextYear, "NEXT_YEAR", "HEADER"

    'Store the hook so that the click event remains alive
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

    'Create a clean header-label font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the previous-month label font explicitly
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

    'Register the previous-month click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the previous-month label to the previous-month action
        LabelHook.Initialize Me, Lbl_PrevMonth, "PREV_MONTH", "HEADER"

    'Store the hook so that the click event remains alive
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

    'Create a clean header-label font object
        Set HeaderFont = CreateObject("StdFont")

    'Configure the next-month label font explicitly
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

    'Register the next-month click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the next-month label to the next-month action
        LabelHook.Initialize Me, Lbl_NextMonth, "NEXT_MONTH", "HEADER"

    'Store the hook so that the click event remains alive
        mHeaderLabelHooks.Add LabelHook, Lbl_NextMonth.Name

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
'------------------------------------------------------------------------------
' DAY OF WEEK
'------------------------------------------------------------------------------

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
'   The calendar grid needs a weekday header row that respects:
'     - the saved first-day-of-week setting
'     - the saved local-dependent / local-independent caption setting
'     - the shared calendar layout constants
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates or reuses:
'     - Lbl_DayOfWeek1 to Lbl_DayOfWeek7
'
'   Weekday captions can be:
'     - locale-dependent, using WeekdayName
'     - locale-independent, using fixed English captions
'
'   The weekday labels are wider than the day labels and are centered above the
'   corresponding day-label columns
'
'   Assigns clean StdFont objects to all weekday labels so reused runtime labels
'   do not retain stale Italic, Underline, or Strikethrough states
'
' ERROR POLICY
'   Raises a descriptive runtime error if invalid layout constants or settings
'   are configured, or if a weekday label cannot be created or formatted
'
' DEPENDENCIES
'   UF_WeekdayCaption_Get
'   UF_Validate_CalendarLayoutConstants
'   M_Settings_IsValidFirstDayOfWeek
'   UF_Ensure_Label
'   MSForms.Label
'   StdFont
'   UserForm.Controls collection
'
' NOTES
'   This routine must stay in the UserForm module because it uses Me.Controls
'
'   It is Public so external refresh routines can update the weekday header when
'   settings are changed while the form is already loaded
'
'   Runtime MSForms labels may retain old font states when reused. This routine
'   therefore assigns clean StdFont objects instead of changing only selected
'   Font properties
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_WeekdayRow_Build"  'Current procedure name

    Dim Index                  As Long                                   'Sequential weekday label index
    Dim ControlName            As String                                 'Runtime control name
    Dim HeaderLeft             As Single                                 'Calculated weekday row left position
    Dim HeaderTop              As Single                                 'Calculated weekday row top position
    Dim DayCaption             As String                                 'Resolved weekday caption
    Dim Lbl_DayOfWeek          As MSForms.Label                          'Weekday label control
    Dim WeekdayFont            As Object                                 'Clean weekday-label font object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE LAYOUT AND SETTINGS
'------------------------------------------------------------------------------
    'Validate the shared calendar layout constants
        UF_Validate_CalendarLayoutConstants PROC_NAME

    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(gDP_FirstDayOfWeek) Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "gDP_FirstDayOfWeek must be vbSunday or vbMonday."
        End If

    'Reject layouts where the header row would be outside the visible form area
        If DP_DAY_GRID_START_TOP - DP_DAY_LABEL_VERTICAL_STEP < 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "DP_DAY_GRID_START_TOP minus DP_DAY_LABEL_VERTICAL_STEP cannot be negative."
        End If

'------------------------------------------------------------------------------
' CALCULATE HEADER POSITION
'------------------------------------------------------------------------------
    'Calculate the left position that centers wider weekday labels above day labels
        HeaderLeft = DP_DAY_GRID_START_LEFT + _
            ((DP_DAY_LABEL_WIDTH - DP_DOW_LABEL_WIDTH) / 2)

    'Calculate the top position of the weekday header row
        HeaderTop = DP_DAY_GRID_START_TOP - DP_DAY_LABEL_VERTICAL_STEP

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

            'Create a clean weekday-label font object
                Set WeekdayFont = CreateObject("StdFont")

            'Configure the weekday label font explicitly
                With WeekdayFont
                    .Name = DP_FORM_FONT_NAME
                    .Size = DP_FORM_FONT_SIZE
                    .Bold = True
                    .Italic = False
                    .Underline = False
                    .Strikethrough = False
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
' ERROR POLICY
'   Raises a descriptive runtime error if Index or FirstDayOfWeek is invalid
'
' DEPENDENCIES
'   M_Settings_IsValidFirstDayOfWeek
'   WeekdayName
'   UF_WeekdayCaption_GetFixedEnglish
'
' NOTES
'   Local captions depend on the user's Windows or Office language environment
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_WeekdayCaption_Get"       'Current procedure name

    Dim DayNumber              As Long                                   'Resolved absolute weekday number
    Dim DayCaption             As String                                 'Resolved weekday caption

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
                "Index must be between 1 and " & CStr(DP_DAY_LABELS_PER_ROW) & "."
        End If

    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(FirstDayOfWeek) Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "FirstDayOfWeek must be vbSunday or vbMonday."
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
            DayCaption = WeekdayName(DayNumber, True, vbSunday)
        Else
            DayCaption = UF_WeekdayCaption_GetFixedEnglish(DayNumber)
        End If

'------------------------------------------------------------------------------
' RETURN CAPTION
'------------------------------------------------------------------------------
    'Return the normalized caption
        UF_WeekdayCaption_Get = UCase$(Trim$(DayCaption))

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Function

'------------------------------------------------------------------------------
' DAY GRID
'------------------------------------------------------------------------------

Private Sub UF_DayGrid_Build()

'
'==============================================================================
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
'   Creates or reuses:
'     - Lbl_DayBg1 to Lbl_DayBg42
'     - Lbl_Day1 to Lbl_Day42
'
'   Both labels in each pair are connected to the same routed action so that
'   hovering or clicking either surface behaves as one calendar day cell
'
' ERROR POLICY
'   Raises a descriptive runtime error if invalid layout constants are configured,
'   if a label cannot be created or formatted, or if an event hook cannot be
'   registered
'
' DEPENDENCIES
'   UF_Validate_CalendarLayoutConstants
'   UF_Ensure_Label
'   MSForms.Label
'   cDatePickerLabelHook
'   mDayLabelHooks
'
' NOTES
'   The background label is placed behind the foreground text label
'
'   The text label remains small and transparent to simulate vertical alignment
'   inside the larger day-cell background
'
' UPDATED
'   2026-04-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DayGrid_Build"       'Current procedure name

    Dim Index                  As Long                                   'Sequential day-cell index
    Dim RowIndex               As Long                                   'Zero-based row index
    Dim ColIndex               As Long                                   'Zero-based column index
    Dim TextLeft               As Single                                 'Day text label left position
    Dim TextTop                As Single                                 'Day text label top position
    Dim CellLeft               As Single                                 'Day background label left position
    Dim CellTop                As Single                                 'Day background label top position
    Dim TextControlName        As String                                 'Runtime day text label name
    Dim BgControlName          As String                                 'Runtime day background label name
    Dim ActionName             As String                                 'Runtime action routed by the label hook
    Dim Lbl_Day                As MSForms.Label                          'Day text label control
    Dim Lbl_DayBg              As MSForms.Label                          'Day background label control
    Dim LabelHook              As cDatePickerLabelHook                   'Runtime label event hook

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

    'Clear the current day-label hover state
        mHoveredDayLabelName = vbNullString

    'Clear the current day-cell hover index
        mHoveredDayCellIndex = 0

'------------------------------------------------------------------------------
' VALIDATE LAYOUT CONSTANTS
'------------------------------------------------------------------------------
    'Validate the shared calendar layout constants
        UF_Validate_CalendarLayoutConstants PROC_NAME

    'Reject invalid day-cell background width
        If DP_DAY_CELL_WIDTH <= 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DP_DAY_CELL_WIDTH must be greater than zero."
        End If

    'Reject invalid day-cell background height
        If DP_DAY_CELL_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "DP_DAY_CELL_HEIGHT must be greater than zero."
        End If

    'Reject invalid grid row count
        If DP_DAY_GRID_ROWS <= 0 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "DP_DAY_GRID_ROWS must be greater than zero."
        End If

    'Reject inconsistent day-label count
        If DP_DAY_GRID_ROWS * DP_DAY_LABELS_PER_ROW <> DP_DAY_LABEL_COUNT Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "DP_DAY_LABEL_COUNT must equal DP_DAY_GRID_ROWS * DP_DAY_LABELS_PER_ROW."
        End If

'------------------------------------------------------------------------------
' CREATE / FORMAT DAY CELLS
'------------------------------------------------------------------------------
    'Loop through the full 6 x 7 calendar grid
        For Index = 1 To DP_DAY_LABEL_COUNT

            'Calculate the zero-based row index
                RowIndex = (Index - 1) \ DP_DAY_LABELS_PER_ROW

            'Calculate the zero-based column index
                ColIndex = (Index - 1) Mod DP_DAY_LABELS_PER_ROW

            'Calculate the original day text position
                TextLeft = DP_DAY_GRID_START_LEFT + DP_DAY_LABEL_HORIZONTAL_STEP * ColIndex

            'Calculate the original day text top position
                TextTop = DP_DAY_GRID_START_TOP + DP_DAY_LABEL_VERTICAL_STEP * RowIndex

            'Center the larger background cell around the text label
                CellLeft = TextLeft - ((DP_DAY_CELL_WIDTH - DP_DAY_LABEL_WIDTH) / 2)

            'Center the larger background cell around the text label
                CellTop = TextTop - ((DP_DAY_CELL_HEIGHT - DP_DAY_LABEL_HEIGHT) / 2)

            'Build the runtime day text label name
                TextControlName = "Lbl_Day" & CStr(Index)

            'Build the runtime day background label name
                BgControlName = "Lbl_DayBg" & CStr(Index)

            'Build the action name routed by both day labels
                ActionName = "DAY_PICKED_" & CStr(Index)

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

            'Register the background-label event hook
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
                    .Caption = CStr(Index)
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

            'Register the text-label event hook
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Creates the cached normal and weekend day fonts when they are missing
'
' ERROR POLICY
'   Raises a descriptive runtime error if the font cache cannot be initialized
'
' DEPENDENCIES
'   StdFont
'   DP_FORM_FONT_NAME
'   DP_FORM_FONT_SIZE
'
' NOTES
'   The cached fonts intentionally reset Italic, Underline, and Strikethrough
'   so reused runtime labels do not inherit stale font states
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DayGrid_EnsureFonts"       'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CREATE NORMAL DAY FONT
'------------------------------------------------------------------------------
    'Create the normal day font when missing
        If mDayFontNormal Is Nothing Then

            'Create a clean font object
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

            'Create a clean font object
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
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub

Public Sub UF_DayGrid_Populate( _
    Optional ByVal DisplayYear As Long = 0, _
    Optional ByVal DisplayMonth As Long = 0)

'
'------------------------------------------------------------------------------
'                           POPULATE DAY LABELS GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Populates the 42 DatePicker day labels for the requested month and year
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
'   Populates Lbl_Day1 to Lbl_Day42 and refreshes the month/year header labels
'
' ERROR POLICY
'   Raises a descriptive runtime error if the requested month/year is invalid,
'   if the first-day setting is invalid, if the grid constants are inconsistent,
'   or if one or more day labels cannot be found
'
' DEPENDENCIES
'   M_Settings_IsValidFirstDayOfWeek
'   M_Caption_GetMonth
'   UF_DayCell_ApplyDateStateByIndex
'
' NOTES
'   Cached StdFont objects are used to avoid repeated font-object creation while
'   still enforcing deterministic normal and weekend font states
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DayGrid_Populate"     'Current procedure name
    
    Const VBA_DATE_MIN_YEAR    As Long = 100                             'Minimum year supported by VBA Date
    Const VBA_DATE_MAX_YEAR    As Long = 9999                            'Maximum year supported by VBA Date

    Dim EffectiveYear          As Long                                   'Validated display year
    Dim EffectiveMonth         As Long                                   'Validated display month
    Dim EffectiveFirstDay      As VbDayOfWeek                            'Validated first day of week
    
    Dim StartOfMonth           As Date                                   'First day of display month
    Dim TrackingDate           As Date                                   'Current date written to the grid
    Dim StartOfMonthDay        As Long                                   'Weekday index of first month day
    
    Dim RowIndex               As Long                                   'Calendar row index
    Dim ColIndex               As Long                                   'Calendar column index
    Dim LabelIndex             As Long                                   'Sequential label index
    
    Dim WeekdaySunBasis        As Long                                   'Sunday-based weekday number
    Dim IsWeekend              As Boolean                                'True for Saturday or Sunday
    
    Dim Lbl_Day                As MSForms.Label                          'Day label control
    Dim Lbl_DayBg              As MSForms.Label                          'Day background label control
        
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Ensure cached day-label references are available
        UF_DayGrid_EnsureCache
    'Ensure cached day-grid fonts are available
        UF_DayGrid_EnsureFonts
    'Clear cached day-cell date state before repopulating the grid
        UF_DayGrid_ClearDateCache

'------------------------------------------------------------------------------
' VALIDATE GRID CONSTANTS
'------------------------------------------------------------------------------
    'Reject invalid grid row count
        If DP_DAY_GRID_ROWS <= 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DP_DAY_GRID_ROWS must be greater than zero."
        End If
    'Reject invalid grid column count
        If DP_DAY_LABELS_PER_ROW <= 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "DP_DAY_LABELS_PER_ROW must be greater than zero."
        End If
    'Reject inconsistent day-label count
        If DP_DAY_GRID_ROWS * DP_DAY_LABELS_PER_ROW <> DP_DAY_LABEL_COUNT Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "DP_DAY_LABEL_COUNT must equal DP_DAY_GRID_ROWS * DP_DAY_LABELS_PER_ROW."
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
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "DisplayYear must be between " & CStr(DP_MIN_YEAR) & " and " & CStr(DP_MAX_YEAR) & "."
        End If
    'Reject invalid display months
        If EffectiveMonth < 1 Or EffectiveMonth > 12 Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "DisplayMonth must be between 1 and 12."
        End If
    'Reject display periods that cannot render a complete 6-week grid
        If EffectiveYear = VBA_DATE_MIN_YEAR And EffectiveMonth = 1 Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "January " & CStr(VBA_DATE_MIN_YEAR) & " cannot render a complete DatePicker grid."
        End If
    'Reject display periods that cannot render a complete 6-week grid
        If EffectiveYear = VBA_DATE_MAX_YEAR And EffectiveMonth = 12 Then
            Err.Raise vbObjectError + 519, PROC_NAME, _
                "December " & CStr(VBA_DATE_MAX_YEAR) & " cannot render a complete DatePicker grid."
        End If

'------------------------------------------------------------------------------
' VALIDATE SETTINGS
'------------------------------------------------------------------------------
    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(gDP_FirstDayOfWeek) Then
            Err.Raise vbObjectError + 520, PROC_NAME, _
                "gDP_FirstDayOfWeek must be vbSunday or vbMonday."
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
    'Back up to the first visible date in the 6 x 7 grid
        TrackingDate = VBA.DateAdd("d", -StartOfMonthDay + 1, StartOfMonth)

'------------------------------------------------------------------------------
' REFRESH HEADER CAPTIONS
'------------------------------------------------------------------------------
    'Update the header month caption
        Me.Controls("Lbl_HeaderMonth").Caption = _
            M_Caption_GetMonth(mDisplayMonth, gDP_UseLocalNames)
    'Update the header year caption
        Me.Controls("Lbl_HeaderYear").Caption = CStr(mDisplayYear)

'------------------------------------------------------------------------------
' POPULATE DAY LABELS
'------------------------------------------------------------------------------
    'Loop through the visible calendar rows
        For RowIndex = 1 To DP_DAY_GRID_ROWS

            'Loop through the visible calendar columns
                For ColIndex = 1 To DP_DAY_LABELS_PER_ROW
                        
                    'Resolve the sequential day-label index
                        LabelIndex = ((RowIndex - 1) * DP_DAY_LABELS_PER_ROW) + ColIndex
                    'Cache the date represented by this day cell
                        mDayCellDates(LabelIndex) = TrackingDate
                    'Mark the cached day-cell date as available
                        mDayCellHasDate(LabelIndex) = True
                    'Retrieve the cached day text label
                        Set Lbl_Day = mDayTextLabels(LabelIndex)
                    'Retrieve the cached day background label
                        Set Lbl_DayBg = mDayBackLabels(LabelIndex)
                    'Resolve Sunday-based weekday number for stable weekend detection
                        WeekdaySunBasis = VBA.Weekday(TrackingDate, vbSunday)
                    'Detect weekends independently from the displayed first-day setting
                        IsWeekend = _
                            (WeekdaySunBasis = vbSaturday Or WeekdaySunBasis = vbSunday)
                    'Apply day-cell data to the background label
                        With Lbl_DayBg
                            .Caption = vbNullString
                            .BackStyle = fmBackStyleOpaque
                            .BackColor = DP_DAY_NORMAL_BACK_COLOR
                            .BorderStyle = fmBorderStyleNone
                            .BorderColor = DP_DAY_NORMAL_BACK_COLOR
                            .SpecialEffect = fmSpecialEffectFlat
                            .Visible = True
                        End With
                    'Apply day-cell data to the text label
                        With Lbl_Day
                            .Caption = CStr(VBA.Day(TrackingDate))

                            .BackStyle = fmBackStyleTransparent
                            .BackColor = DP_DAY_NORMAL_BACK_COLOR
                            .BorderStyle = fmBorderStyleNone
                            .BorderColor = DP_DAY_NORMAL_BACK_COLOR
                            .SpecialEffect = fmSpecialEffectFlat
                            .Visible = True
                        End With

                    'Apply date-driven visual state to the paired day cell
                        UF_DayCell_ApplyDateStateByIndex LabelIndex, TrackingDate

                    'Assign the cached weekend font when weekend highlighting is enabled
                        If gDP_HighlightWeekends And IsWeekend Then
                            Set Lbl_Day.Font = mDayFontWeekend
                        Else
                            Set Lbl_Day.Font = mDayFontNormal
                        End If
                    'Advance to the next visible date
                        TrackingDate = VBA.DateAdd("d", 1, TrackingDate)

                Next ColIndex

        Next RowIndex

'------------------------------------------------------------------------------
' RESET HOVER STATE
'------------------------------------------------------------------------------
    'Clear any previous hover state after repopulating the grid
        mHoveredDayLabelName = vbNullString
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   resolved
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
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DayGrid_EnsureCache"       'Current procedure name

    Dim Index                  As Long                                   'Sequential day-cell index

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

            'Resolve the cached day text label when missing
                If mDayTextLabels(Index) Is Nothing Then
                    Set mDayTextLabels(Index) = Me.Controls("Lbl_Day" & CStr(Index))
                End If

            'Resolve the cached day background label when missing
                If mDayBackLabels(Index) Is Nothing Then
                    Set mDayBackLabels(Index) = Me.Controls("Lbl_DayBg" & CStr(Index))
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub

Private Sub UF_DayGrid_ClearDateCache()

'
'------------------------------------------------------------------------------
'                         CLEAR DAY GRID DATE CACHE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the cached date state for all visible DatePicker day cells
'
' WHY THIS EXISTS
'   The DatePicker stores the date represented by each visible day cell so hover,
'   selection, keyboard refresh, and targeted visual updates do not need to read
'   or parse the label Tag property repeatedly
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears mDayCellDates and mDayCellHasDate for all day-cell indexes
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   mDayCellDates
'   mDayCellHasDate
'
' NOTES
'   This routine clears only cached date state, not cached label references
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index       As Long     'Day-cell index

'------------------------------------------------------------------------------
' CLEAR CACHE
'------------------------------------------------------------------------------
    'Loop through all cached day-cell date slots
        For Index = 1 To DP_DAY_LABEL_COUNT
            'Clear the cached date
                mDayCellDates(Index) = 0
            'Mark the cached date as unavailable
                mDayCellHasDate(Index) = False
        Next Index

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
'   terminating
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
'   Suppresses cleanup errors because the form is already terminating
'
' DEPENDENCIES
'   None
'
' NOTES
'   This routine does not delete controls
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index               As Long     'Sequential day-cell index

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

    'Loop through all cached day cells
        For Index = 1 To DP_DAY_LABEL_COUNT
            'Clear the cached day text label
                Set mDayTextLabels(Index) = Nothing
            'Clear the cached day background label
                Set mDayBackLabels(Index) = Nothing
        Next Index

    'Restore normal error handling
        On Error GoTo 0

End Sub


Private Sub UF_DayCell_ApplyDateStateByIndex( _
    ByVal DayIndex As Long, _
    ByVal CellDate As Date)

'
'------------------------------------------------------------------------------
'                       APPLY DAY CELL DATE STATE BY INDEX
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the normal, outside-month, disabled, today, and selected-date visual
'   state to one paired DatePicker day cell using its numeric index
'
' WHY THIS EXISTS
'   The original UF_DayCell_ApplyDateState resolves the day-cell index from a
'   label name and then retrieves the paired controls from Me.Controls
'
'   During grid population, hover reset, and minimal keyboard refresh, the caller
'   often already knows the day-cell index. This routine avoids repeated string
'   parsing and repeated Controls collection lookups
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
'   - Validates DayIndex
'   - Ensures cached day-label references are available
'   - Retrieves the paired text and background labels from cache
'   - Applies base visual state
'   - Applies outside-month state
'   - Applies today state
'   - Applies selected-date / keyboard-date state
'
' ERROR POLICY
'   Raises a descriptive runtime error if DayIndex is invalid, if cached labels
'   are unavailable, or if the day cell cannot be formatted
'
' DEPENDENCIES
'   UF_DayGrid_EnsureCache
'   M_DatePolicy_CanSelectDate
'   gDP_HasSelectedDate
'   gDP_SelectedDate
'   mDayTextLabels
'   mDayBackLabels
'
' NOTES
'   Selected-date formatting has priority over today formatting
'
'   This routine intentionally does not call ZOrder. Layering is established
'   when the grid is built and should not be repeated during state refresh
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DayCell_ApplyDateStateByIndex" 'Current procedure name

    Dim IsOutsideMonth         As Boolean                                'True for adjacent-month dates
    Dim IsToday                As Boolean                                'True when CellDate is today
    Dim IsSelected             As Boolean                                'True when CellDate is selected
    Dim IsSelectable           As Boolean                                'True when the date can be selected
    Dim Lbl_Text               As MSForms.Label                          'Day text label
    Dim Lbl_Bg                 As MSForms.Label                          'Day background label

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
                "DayIndex must be between 1 and " & CStr(DP_DAY_LABEL_COUNT) & "."
        End If

'------------------------------------------------------------------------------
' ENSURE CACHE
'------------------------------------------------------------------------------
    'Ensure cached day-label references are available
        UF_DayGrid_EnsureCache

'------------------------------------------------------------------------------
' RETRIEVE CACHED LABELS
'------------------------------------------------------------------------------
    'Retrieve the cached day text label
        Set Lbl_Text = mDayTextLabels(DayIndex)

    'Retrieve the cached day background label
        Set Lbl_Bg = mDayBackLabels(DayIndex)

    'Reject a missing cached text label
        If Lbl_Text Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Cached text label is missing for DayIndex " & CStr(DayIndex) & "."
        End If

    'Reject a missing cached background label
        If Lbl_Bg Is Nothing Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Cached background label is missing for DayIndex " & CStr(DayIndex) & "."
        End If

'------------------------------------------------------------------------------
' RESOLVE STATE
'------------------------------------------------------------------------------
    'Resolve whether the cell belongs to an adjacent month
        IsOutsideMonth = _
            (VBA.Year(CellDate) <> mDisplayYear Or VBA.Month(CellDate) <> mDisplayMonth)

    'Resolve whether the cell represents today
        IsToday = (VBA.DateValue(CellDate) = VBA.Date)

    'Resolve whether the cell represents the keyboard-selected date
        If mHasKeyboardDate Then
            IsSelected = (VBA.DateValue(CellDate) = VBA.DateValue(mKeyboardDate))
        ElseIf gDP_HasSelectedDate Then
            IsSelected = (VBA.DateValue(CellDate) = VBA.DateValue(gDP_SelectedDate))
        Else
            IsSelected = False
        End If

    'Resolve whether the date can be selected
        IsSelectable = M_DatePolicy_CanSelectDate(CellDate, mDisplayYear, mDisplayMonth)

'------------------------------------------------------------------------------
' APPLY BACKGROUND BASE STATE
'------------------------------------------------------------------------------
    'Apply normal background state
        With Lbl_Bg
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
            .BackStyle = fmBackStyleTransparent
            .ForeColor = DP_DAY_CURRENT_MONTH_FORE_COLOR
            .BorderStyle = fmBorderStyleNone
            .BorderColor = DP_DAY_NORMAL_BACK_COLOR
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = IsSelectable
            .Visible = True
        End With

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
    'Apply today background highlight when applicable
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

            'Apply selected background formatting
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Scans cached visible dates and refreshes the first matching day cell
'
' ERROR POLICY
'   Raises a descriptive runtime error if a matching visible day cell cannot be
'   refreshed
'
' DEPENDENCIES
'   mDayCellDates
'   mDayCellHasDate
'   UF_DayCell_ApplyDateStateByIndex
'
' NOTES
'   If TargetDate is not visible in the current 6 x 7 grid, this routine exits
'   without changing anything
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DayCell_RefreshVisibleDate" 'Current procedure name

    Dim Index               As Long     'Day-cell index
    Dim TargetDateOnly      As Date     'Target date without time

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
                    'Refresh the matching day cell
                        If VBA.DateValue(mDayCellDates(Index)) = TargetDateOnly Then
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub
Public Sub UF_DayCell_HoverApply(ByVal LabelName As String)

'
'==============================================================================
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
'   Resolves the day-cell index, resets the previous hover state, and applies
'   hover formatting to the paired day-cell background while preserving selected
'   date readability
'
' ERROR POLICY
'   Raises a descriptive runtime error if LabelName is empty, unsupported, or if
'   the target day cell cannot be formatted
'
' DEPENDENCIES
'   UF_DayCell_GetIndexFromLabelName
'   UF_DayCell_HoverReset
'   UserForm.Controls collection
'
' NOTES
'   Hover formatting is applied primarily to the background label
'
' UPDATED
'   2026-04-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DayCell_HoverApply"       'Current procedure name

    Dim DayIndex               As Long                                   'Resolved day-cell index
    Dim LabelDate              As Date                                   'Date represented by the label
    Dim HasLabelDate           As Boolean                                'True when the label Tag contains a date
    
    Dim IsSelected             As Boolean                                'True when the day cell is selected
    Dim IsOutsideMonth         As Boolean                                'True when the day cell belongs to an adjacent month
    
    Dim Lbl_Text               As MSForms.Label                          'Day text label
    Dim Lbl_Bg                 As MSForms.Label                          'Day background label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty label name
        If Len(Trim$(LabelName)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "LabelName cannot be empty."
        End If

'------------------------------------------------------------------------------
' RESOLVE DAY CELL
'------------------------------------------------------------------------------
    'Resolve the paired day-cell index
        DayIndex = UF_DayCell_GetIndexFromLabelName(LabelName)
    'Reject invalid day-cell indexes
        If DayIndex < 1 Or DayIndex > DP_DAY_LABEL_COUNT Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "LabelName does not resolve to a valid day-cell index."
        End If

'------------------------------------------------------------------------------
' EXIT IF ALREADY HOVERED
'------------------------------------------------------------------------------
    'Exit if the requested day cell is already highlighted
        If mHoveredDayCellIndex = DayIndex Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RESET CURRENT HOVER STATE
'------------------------------------------------------------------------------
    'Remove hover formatting from the previously highlighted day cell
        UF_DayCell_HoverReset

'------------------------------------------------------------------------------
' RETRIEVE TARGET LABELS
'------------------------------------------------------------------------------
    'Retrieve the day text label
        Set Lbl_Text = Me.Controls("Lbl_Day" & CStr(DayIndex))
    'Retrieve the day background label
        Set Lbl_Bg = Me.Controls("Lbl_DayBg" & CStr(DayIndex))

'------------------------------------------------------------------------------
' RESOLVE LABEL DATE STATE
'------------------------------------------------------------------------------
    'Resolve the label date from the cached day-cell date
        If mDayCellHasDate(DayIndex) Then
            LabelDate = mDayCellDates(DayIndex)
            HasLabelDate = True
        End If
    'Resolve whether the hovered day cell is the selected date
        If HasLabelDate Then
            If mHasKeyboardDate Then
                IsSelected = (VBA.DateValue(LabelDate) = VBA.DateValue(mKeyboardDate))
            ElseIf gDP_HasSelectedDate Then
                IsSelected = (VBA.DateValue(LabelDate) = VBA.DateValue(gDP_SelectedDate))
            End If
        End If
    'Resolve whether the hovered day cell belongs to an adjacent month
        If HasLabelDate Then
            IsOutsideMonth = _
                (VBA.Year(LabelDate) <> mDisplayYear Or VBA.Month(LabelDate) <> mDisplayMonth)
        End If

'------------------------------------------------------------------------------
' APPLY SELECTED HOVER STATE
'------------------------------------------------------------------------------
    'Preserve selected-date colors when hovering the selected day cell
        If IsSelected Then

            'Apply selected hover background formatting
                With Lbl_Bg
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = DP_DAY_SELECTED_BACK_COLOR
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = DP_DAY_SELECTED_BACK_COLOR
                    .SpecialEffect = fmSpecialEffectBump
                End With

            'Apply selected text formatting
                Lbl_Text.ForeColor = DP_DAY_SELECTED_FORE_COLOR
            'Store the current hovered day-cell index
                mHoveredDayCellIndex = DayIndex
            'Store the current hovered day-label name for compatibility
                mHoveredDayLabelName = "Lbl_Day" & CStr(DayIndex)
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
' RESTORE LAYERING
'------------------------------------------------------------------------------
    'Move the background label behind the text label
        Lbl_Bg.ZOrder 1
    'Move the text label to the front
        Lbl_Text.ZOrder 0

'------------------------------------------------------------------------------
' STORE HOVER STATE
'------------------------------------------------------------------------------
    'Store the current hovered day-cell index
        mHoveredDayCellIndex = DayIndex
    'Store the current hovered day-label name for compatibility
        mHoveredDayLabelName = "Lbl_Day" & CStr(DayIndex)

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
Private Sub UF_DayCell_HoverReset()

'
'==============================================================================
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
'
' ERROR POLICY
'   Silently ignores a missing previously-hovered day cell because controls may
'   have been rebuilt or removed during a refresh
'
' DEPENDENCIES
'   UserForm.Controls collection
'   UF_DayCell_ApplyDateState
'
' NOTES
'   This routine resets only the currently highlighted day cell, not all 42 cells
'
' UPDATED
'   2026-04-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_DayCell_HoverReset"     'Current procedure name

    Dim Lbl_Text            As MSForms.Label        'Previously highlighted day text label
    Dim Lbl_Bg              As MSForms.Label        'Previously highlighted day background label
    
    Dim LabelDate           As Date                 'Date represented by the day cell

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' EXIT IF NOTHING IS HOVERED
'------------------------------------------------------------------------------
    'Exit if no day cell is currently highlighted
        If mHoveredDayCellIndex = 0 Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RETRIEVE CURRENTLY HOVERED LABELS
'------------------------------------------------------------------------------
    'Temporarily ignore lookup errors if the previous labels no longer exist
        On Error Resume Next
    'Retrieve the previously highlighted day text label
        Set Lbl_Text = Me.Controls("Lbl_Day" & CStr(mHoveredDayCellIndex))
    'Retrieve the previously highlighted day background label
        Set Lbl_Bg = Me.Controls("Lbl_DayBg" & CStr(mHoveredDayCellIndex))
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESET VISUAL STATE
'------------------------------------------------------------------------------
    'Reset the day cell if the text label was found
        If Not Lbl_Text Is Nothing Then
            'Restore the full date-driven visual state when the cached date is available
                If mDayCellHasDate(mHoveredDayCellIndex) Then
                    'Read the cached day-cell date
                        LabelDate = mDayCellDates(mHoveredDayCellIndex)
                    'Restore the complete date-driven state
                        UF_DayCell_ApplyDateStateByIndex mHoveredDayCellIndex, LabelDate
                Else
                    'Apply a safe normal visual fallback to the background label
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
                    'Apply a safe normal visual fallback to the text label
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
        End If

'------------------------------------------------------------------------------
' CLEAR HOVER STATE
'------------------------------------------------------------------------------
    'Clear the current hovered day-label name
        mHoveredDayLabelName = vbNullString

    'Clear the current hovered day-cell index
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub
'------------------------------------------------------------------------------
' FOOTER
'------------------------------------------------------------------------------

Private Sub UF_Footer_Build()

'
'------------------------------------------------------------------------------
'                           CREATE FOOTER SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Creates the DatePicker footer divider, footer banner, and footer labels
'
' WHY THIS EXISTS
'   The footer area visually separates quick Today / Time shortcuts from the
'   calendar grid
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds or refreshes the full footer section
'
' ERROR POLICY
'   Propagates unexpected runtime errors raised by the called routines
'
' DEPENDENCIES
'   UF_Footer_BuildDivider
'   UF_Footer_BuildBanner
'   UF_Footer_BuildLabels
'
' NOTES
'   This routine coordinates footer creation only
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

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
    'Create or format the footer labels
        UF_Footer_BuildLabels

'------------------------------------------------------------------------------
' CREATE SETTINGS AREA
'------------------------------------------------------------------------------
    'Create or format the settings separator and icon
        UF_Footer_BuildSettingsArea
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
'   Creates or reuses Lbl_Divider
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
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Footer_BuildDivider"       'Current procedure name

    Dim DividerTop             As Single                                 'Divider top position
    Dim DividerWidth           As Single                                 'Divider width
    Dim Lbl_Divider            As MSForms.Label                          'Divider label control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE CONSTANTS
'------------------------------------------------------------------------------
    'Reject invalid divider height
        If DP_DIVIDER_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "DP_DIVIDER_HEIGHT must be greater than zero."
        End If
    'Reject invalid divider side margin
        If DP_DIVIDER_SIDE_MARGIN < 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "DP_DIVIDER_SIDE_MARGIN cannot be negative."
        End If
    'Reject invalid calculated divider width
        If DP_FORM_WIDTH - (DP_DIVIDER_SIDE_MARGIN * 2) <= 0 Then
            Err.Raise vbObjectError + 515, PROC_NAME, "Divider width must be greater than zero."
        End If

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
            .BackColor = DP_DIVIDER_BACK_COLOR
            .ForeColor = DP_DIVIDER_COLOR
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = DP_DIVIDER_COLOR
            .SpecialEffect = fmSpecialEffectFlat
            .Enabled = True
            .Visible = True
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
'   Creates or reuses Lbl_FooterBanner
'
' ERROR POLICY
'   Raises a descriptive runtime error if footer layout constants are invalid
'   or if the banner cannot be created or formatted
'
' DEPENDENCIES
'   UF_Ensure_Label
'   UF_CalendarGrid_GetWidth
'   UF_CalendarGrid_GetBottom
'
' NOTES
'   The footer banner is implemented as a runtime-created MSForms.Label
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Footer_BuildBanner"        'Current procedure name

    Dim BannerLeft             As Single                                 'Footer banner left position
    Dim BannerTop              As Single                                 'Footer banner top position
    Dim BannerWidth            As Single                                 'Footer banner width
    Dim Lbl_FooterBanner       As MSForms.Label                          'Footer banner control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE CONSTANTS
'------------------------------------------------------------------------------
    'Reject invalid footer height
        If DP_FOOTER_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "DP_FOOTER_HEIGHT must be greater than zero."
        End If

    'Reject invalid footer side padding
        If DP_FOOTER_SIDE_PADDING < 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "DP_FOOTER_SIDE_PADDING cannot be negative."
        End If

    'Reject invalid calculated banner width
        If UF_CalendarGrid_GetWidth + (DP_FOOTER_SIDE_PADDING * 2) <= 0 Then
            Err.Raise vbObjectError + 515, PROC_NAME, "Footer banner width must be greater than zero."
        End If

'------------------------------------------------------------------------------
' CALCULATE LAYOUT
'------------------------------------------------------------------------------
    'Calculate the banner left position
        BannerLeft = DP_DAY_GRID_START_LEFT - DP_FOOTER_SIDE_PADDING

    'Calculate the banner top position
        BannerTop = UF_CalendarGrid_GetBottom() + DP_FOOTER_TOP_GAP

    'Calculate the banner width
        BannerWidth = UF_CalendarGrid_GetWidth + (DP_FOOTER_SIDE_PADDING * 2)

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
            .Enabled = True
            .Visible = True
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
'   Creates or reuses:
'     - Lbl_TodayIcon
'     - Lbl_TodayCaption
'     - Lbl_Today
'     - Lbl_TimeIcon
'     - Lbl_TimeCaption
'     - Lbl_Time
'
'   Registers footer click hooks so that:
'     - clicking Today icon, caption, or value writes today's date without time
'     - clicking Time icon, caption, or value writes today's date with the
'       current system time
'
' ERROR POLICY
'   Raises a descriptive runtime error if footer constants are invalid, if one or
'   more labels cannot be created or formatted, or if one or more footer click
'   hooks cannot be registered
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
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_Footer_BuildLabels"     'Current procedure name

    Dim BannerLeft              As Single                               'Footer banner left position
    Dim BannerTop               As Single                               'Footer banner top position

    Dim Lbl_TodayIcon           As MSForms.Label                        'Footer Today icon label
    Dim Lbl_TodayCaption        As MSForms.Label                        'Footer Today caption label
    Dim Lbl_Today               As MSForms.Label                        'Footer Today value label

    Dim Lbl_TimeIcon            As MSForms.Label                        'Footer Time icon label
    Dim Lbl_TimeCaption         As MSForms.Label                        'Footer Time caption label
    Dim Lbl_Time                As MSForms.Label                        'Footer Time value label

    Dim Lbl_TodayHalo           As MSForms.Label                        'Footer Today halo label
    Dim Lbl_TimeHalo            As MSForms.Label                        'Footer Time halo label
    
    Dim Lbl_SettingsIcon        As MSForms.Label                        'Footer Settings icon label
    
    Dim LabelHook               As cDatePickerLabelHook                 'Runtime footer click hook
    Dim CaptionFont             As Object                               'Clean caption-label font object
    Dim IconFont                As Object                               'Clean icon font object
    Dim NewFont                 As Object                               'Clean value-label font object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Create or reset footer-label click hooks
        Set mFooterLabelHooks = New Collection

'------------------------------------------------------------------------------
' VALIDATE FOOTER CONSTANTS
'------------------------------------------------------------------------------
    'Reject invalid footer icon font name
        If Len(Trim$(DP_FOOTER_ICON_FONT_NAME)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DP_FOOTER_ICON_FONT_NAME cannot be empty."
        End If
    'Reject invalid footer icon dimensions
        If DP_FOOTER_ICON_WIDTH <= 0 Or DP_FOOTER_ICON_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "DP_FOOTER_ICON_WIDTH and DP_FOOTER_ICON_HEIGHT must be greater than zero."
        End If
    'Reject invalid footer icon font size
        If DP_FOOTER_ICON_FONT_SIZE <= 0 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "DP_FOOTER_ICON_FONT_SIZE must be greater than zero."
        End If
    'Reject invalid Today label widths
        If DP_FOOTER_TODAY_CAPTION_WIDTH <= 0 Or DP_FOOTER_TODAY_VALUE_WIDTH <= 0 Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Today footer label widths must be greater than zero."
        End If
    'Reject invalid Time label widths
        If DP_FOOTER_TIME_CAPTION_WIDTH <= 0 Or DP_FOOTER_TIME_VALUE_WIDTH <= 0 Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Time footer label widths must be greater than zero."
        End If

'------------------------------------------------------------------------------
' CALCULATE LAYOUT
'------------------------------------------------------------------------------
    'Calculate the footer banner left position
        BannerLeft = DP_DAY_GRID_START_LEFT - DP_FOOTER_SIDE_PADDING
    'Calculate the footer banner top position
        BannerTop = UF_CalendarGrid_GetBottom() + DP_FOOTER_TOP_GAP

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
            .SpecialEffect = fmSpecialEffectBump
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
            .SpecialEffect = fmSpecialEffectBump
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

    'Create a clean icon font object
        Set IconFont = CreateObject("StdFont")

    'Configure the clean icon font explicitly
        With IconFont
            .Name = DP_FOOTER_ICON_FONT_NAME
            .Size = DP_FOOTER_ICON_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean icon font to the Today icon label
        Set Lbl_TodayIcon.Font = IconFont

    'Apply the Today icon glyph after the clean font has been assigned
        Lbl_TodayIcon.Caption = ChrW$(DP_FOOTER_ICON_CALENDAR_CODEPOINT)

    'Move the Today icon label to the front
        Lbl_TodayIcon.ZOrder 0

    'Register the Today icon click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the Today icon to the write-today action
        LabelHook.Initialize Me, Lbl_TodayIcon, "WRITE_TODAY", "FOOTER"

    'Store the hook so that the click event remains alive
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
        End With

    'Create a clean caption-label font object
        Set CaptionFont = CreateObject("StdFont")

    'Configure the Today caption font explicitly
        With CaptionFont
            .Name = DP_FORM_FONT_NAME
            .Size = 10
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean font back to the Today caption label
        Set Lbl_TodayCaption.Font = CaptionFont

    'Apply the caption text after the clean font has been assigned
        Lbl_TodayCaption.Caption = "Today"

    'Move the Today caption label to the front
        Lbl_TodayCaption.ZOrder 0

    'Register the Today caption click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the Today caption to the write-today action
        LabelHook.Initialize Me, Lbl_TodayCaption, "WRITE_TODAY", "FOOTER"

    'Store the hook so that the click event remains alive
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
        End With

    'Create a clean value-label font object
        Set NewFont = CreateObject("StdFont")

    'Configure the Today value-label font explicitly
        With NewFont
            .Name = DP_FORM_FONT_NAME
            .Size = 8
            .Italic = False
            .Underline = False
            .Strikethrough = False
            .Bold = False
        End With

    'Assign the clean font back to the Today value label
        Set Lbl_Today.Font = NewFont

    'Apply the Today value caption after the clean font has been assigned
        Lbl_Today.Caption = M_Caption_GetDate(VBA.Date, gDP_UseLocalNames)

    'Move the Today value label to the front
        Lbl_Today.ZOrder 0

    'Register the Today value click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the Today value to the write-today action
        LabelHook.Initialize Me, Lbl_Today, "WRITE_TODAY", "FOOTER"

    'Store the hook so that the click event remains alive
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

    'Create a clean icon font object
        Set IconFont = CreateObject("StdFont")

    'Configure the clean icon font explicitly
        With IconFont
            .Name = DP_FOOTER_ICON_FONT_NAME
            .Size = DP_FOOTER_ICON_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean icon font to the Time icon label
        Set Lbl_TimeIcon.Font = IconFont

    'Apply the Time icon glyph after the clean font has been assigned
        Lbl_TimeIcon.Caption = ChrW$(DP_FOOTER_ICON_CLOCK_CODEPOINT)

    'Move the Time icon label to the front
        Lbl_TimeIcon.ZOrder 0

    'Register the Time icon click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the Time icon to the write-now action
        LabelHook.Initialize Me, Lbl_TimeIcon, "WRITE_NOW", "FOOTER"

    'Store the hook so that the click event remains alive
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
        End With

    'Create a clean caption-label font object
        Set CaptionFont = CreateObject("StdFont")

    'Configure the Time caption font explicitly
        With CaptionFont
            .Name = DP_FORM_FONT_NAME
            .Size = 10
            .Bold = True
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean font back to the Time caption label
        Set Lbl_TimeCaption.Font = CaptionFont

    'Apply the caption text after the clean font has been assigned
        Lbl_TimeCaption.Caption = "Time"

    'Move the Time caption label to the front
        Lbl_TimeCaption.ZOrder 0

    'Register the Time caption click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the Time caption to the write-now action
        LabelHook.Initialize Me, Lbl_TimeCaption, "WRITE_NOW", "FOOTER"

    'Store the hook so that the click event remains alive
        mFooterLabelHooks.Add LabelHook, Lbl_TimeCaption.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT TIME VALUE LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the Time value label
        Set Lbl_Time = UF_Ensure_Label("Lbl_Time")

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
        End With

    'Create a clean value-label font object
        Set NewFont = CreateObject("StdFont")

    'Configure the Time value-label font explicitly
        With NewFont
            .Name = DP_FORM_FONT_NAME
            .Size = 8
            .Italic = False
            .Underline = False
            .Strikethrough = False
            .Bold = False
        End With

    'Assign the clean font back to the Time value label
        Set Lbl_Time.Font = NewFont

    'Apply the Time value caption after the clean font has been assigned
        Lbl_Time.Caption = Format$(VBA.Time, "hh:nn:ss")

    'Move the Time value label to the front
        Lbl_Time.ZOrder 0

    'Register the Time value click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the Time value to the write-now action
        LabelHook.Initialize Me, Lbl_Time, "WRITE_NOW", "FOOTER"

    'Store the hook so that the click event remains alive
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Settings is visually smaller than Today and Time, but it should follow the
'   same hover model:
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
'   Creates or reuses:
'     - Lbl_SettingsSeparator
'     - Lbl_SettingsHalo
'     - Lbl_SettingsIcon
'
'   Positions the Settings group directly on the UserForm, aligned to the footer
'   banner geometry
'
' ERROR POLICY
'   Raises a descriptive runtime error if the labels cannot be created,
'   formatted, or hooked
'
' DEPENDENCIES
'   UF_Ensure_Label
'   UF_CalendarGrid_GetBottom
'   cDatePickerLabelHook
'
' NOTES
'   There is no footer frame in this form
'
'   Footer controls are positioned directly on the UserForm
'
'   The halo is disabled because it is purely visual. The icon is the event
'   surface for click and hover routing
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_Footer_BuildSettingsArea"  'Current procedure name

    Dim LabelHook               As cDatePickerLabelHook                   'Runtime footer click / hover hook
    Dim BannerLeft              As Single                                 'Footer banner left position
    Dim BannerTop               As Single                                 'Footer banner top position

    Dim Lbl_Separator           As MSForms.Label                          'Footer separator label
    Dim Lbl_SettingsHalo        As MSForms.Label                          'Footer Settings halo label
    Dim Lbl_SettingsIcon        As MSForms.Label                          'Footer Settings icon label

    Dim IconFont                As Object                                 'Clean settings icon font object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE SETTINGS CONSTANTS
'------------------------------------------------------------------------------
    'Reject invalid separator dimensions
        If DP_FOOTER_SEPARATOR_WIDTH <= 0 Or DP_FOOTER_SEPARATOR_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Settings separator width and height must be greater than zero."
        End If

    'Reject invalid settings halo dimensions
        If DP_FOOTER_SETTINGS_HALO_WIDTH <= 0 Or DP_FOOTER_SETTINGS_HALO_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Settings halo width and height must be greater than zero."
        End If

    'Reject invalid settings icon dimensions
        If DP_FOOTER_SETTINGS_ICON_WIDTH <= 0 Or DP_FOOTER_SETTINGS_ICON_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Settings icon width and height must be greater than zero."
        End If

    'Reject an empty settings icon font name
        If Len(Trim$(DP_FOOTER_SETTINGS_ICON_FONT_NAME)) = 0 Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "DP_FOOTER_SETTINGS_ICON_FONT_NAME cannot be empty."
        End If

    'Reject invalid settings icon font size
        If DP_FOOTER_SETTINGS_ICON_FONT_SIZE <= 0 Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "DP_FOOTER_SETTINGS_ICON_FONT_SIZE must be greater than zero."
        End If

'------------------------------------------------------------------------------
' CALCULATE FOOTER POSITION
'------------------------------------------------------------------------------
    'Calculate the footer banner left position
        BannerLeft = DP_DAY_GRID_START_LEFT - DP_FOOTER_SIDE_PADDING

    'Calculate the footer banner top position
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
            .SpecialEffect = fmSpecialEffectBump
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

    'Store the hook so that the event remains alive
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   hovered, resets the previous footer hover state, retrieves the current
'   group labels, validates required labels, reveals the halo, applies hover
'   foreground formatting, restores layering, and stores the new hover state
'
' ERROR POLICY
'   Raises a descriptive runtime error if:
'     - LabelName is blank
'     - LabelName does not resolve to a supported footer action
'     - the required halo or icon label is missing
'     - Today or Time is missing its caption or value label
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
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Footer_HoverApply"          'Current procedure name

    Dim ActionName             As String                                  'Resolved footer action name
    Dim RequiresTextLabels     As Boolean                                 'True when caption and value are required

    Dim Lbl_Halo               As MSForms.Label                           'Footer halo label
    Dim Lbl_Icon               As MSForms.Label                           'Footer icon label
    Dim Lbl_Caption            As MSForms.Label                           'Footer caption label
    Dim Lbl_Value              As MSForms.Label                           'Footer value label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty label name
        If Len(Trim$(LabelName)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "LabelName cannot be empty"
        End If

'------------------------------------------------------------------------------
' RESOLVE FOOTER ACTION
'------------------------------------------------------------------------------
    'Resolve the footer action from the hovered label name
        ActionName = UF_Footer_ActionFromLabelName(LabelName)

    'Reject unsupported footer labels
        If Len(ActionName) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "LabelName does not resolve to a supported footer action"
        End If

'------------------------------------------------------------------------------
' EXIT IF ALREADY HOVERED
'------------------------------------------------------------------------------
    'Exit if the requested footer action is already highlighted
        If StrComp(mHoveredFooterActionName, ActionName, vbBinaryCompare) = 0 Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RESET CURRENT HOVER STATE
'------------------------------------------------------------------------------
    'Remove hover formatting from the previously highlighted footer action
        UF_Footer_HoverReset

'------------------------------------------------------------------------------
' RETRIEVE FOOTER GROUP LABELS
'------------------------------------------------------------------------------
    'Suppress lookup errors so missing labels can be reported consistently
        On Error Resume Next

    'Retrieve labels for the requested footer action
        Select Case ActionName

            Case "WRITE_TODAY"
                Set Lbl_Halo = Me.Controls("Lbl_TodayHalo")
                Set Lbl_Icon = Me.Controls("Lbl_TodayIcon")
                Set Lbl_Caption = Me.Controls("Lbl_TodayCaption")
                Set Lbl_Value = Me.Controls("Lbl_Today")
                RequiresTextLabels = True

            Case "WRITE_NOW"
                Set Lbl_Halo = Me.Controls("Lbl_TimeHalo")
                Set Lbl_Icon = Me.Controls("Lbl_TimeIcon")
                Set Lbl_Caption = Me.Controls("Lbl_TimeCaption")
                Set Lbl_Value = Me.Controls("Lbl_Time")
                RequiresTextLabels = True

            Case "SHOW_SETTINGS"
                Set Lbl_Halo = Me.Controls("Lbl_SettingsHalo")
                Set Lbl_Icon = Me.Controls("Lbl_SettingsIcon")
                RequiresTextLabels = False

            Case Else
                Err.Raise vbObjectError + 515, PROC_NAME, _
                    "Unsupported footer action: " & ActionName

        End Select

    'Restore controlled error handling
        On Error GoTo ErrorHandler

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

    'Reject a missing footer caption when the action requires one
        If RequiresTextLabels Then
            If Lbl_Caption Is Nothing Then
                Err.Raise vbObjectError + 518, PROC_NAME, _
                    "Footer caption label is missing for action " & ActionName
            End If
        End If

    'Reject a missing footer value when the action requires one
        If RequiresTextLabels Then
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
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectBump
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Best-effort cleanup
'
'   Suppresses reset errors because hover reset should never interrupt UserForm
'   interaction
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
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentAction          As String                                  'Current hovered footer action

    Dim Lbl_Halo              As MSForms.Label                           'Footer halo label
    Dim Lbl_Icon              As MSForms.Label                           'Footer icon label
    Dim Lbl_Caption           As MSForms.Label                           'Footer caption label
    Dim Lbl_Value             As MSForms.Label                           'Footer value label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress reset errors
        On Error Resume Next

    'Capture the current hovered action
        CurrentAction = Trim$(mHoveredFooterActionName)

'------------------------------------------------------------------------------
' EXIT IF NOTHING IS HOVERED
'------------------------------------------------------------------------------
    'Exit if no footer action is currently highlighted
        If Len(CurrentAction) = 0 Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RETRIEVE FOOTER GROUP LABELS
'------------------------------------------------------------------------------
    'Retrieve labels for the currently hovered footer action
        Select Case CurrentAction

            Case "WRITE_TODAY"
                Set Lbl_Halo = Me.Controls("Lbl_TodayHalo")
                Set Lbl_Icon = Me.Controls("Lbl_TodayIcon")
                Set Lbl_Caption = Me.Controls("Lbl_TodayCaption")
                Set Lbl_Value = Me.Controls("Lbl_Today")

            Case "WRITE_NOW"
                Set Lbl_Halo = Me.Controls("Lbl_TimeHalo")
                Set Lbl_Icon = Me.Controls("Lbl_TimeIcon")
                Set Lbl_Caption = Me.Controls("Lbl_TimeCaption")
                Set Lbl_Value = Me.Controls("Lbl_Time")

            Case "SHOW_SETTINGS"
                Set Lbl_Halo = Me.Controls("Lbl_SettingsHalo")
                Set Lbl_Icon = Me.Controls("Lbl_SettingsIcon")

            Case Else
                mHoveredFooterActionName = vbNullString
                On Error GoTo 0
                Exit Sub

        End Select

'------------------------------------------------------------------------------
' HIDE GROUP HALO
'------------------------------------------------------------------------------
    'Hide the footer halo when available
        If Not Lbl_Halo Is Nothing Then
            Lbl_Halo.Visible = False
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
' CLEAR HOVER STATE
'------------------------------------------------------------------------------
    'Clear the current footer hover action
        mHoveredFooterActionName = vbNullString

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Function UF_Footer_ActionFromLabelName(ByVal LabelName As String) As String

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the supplied label name
        LabelName = Trim$(LabelName)

'------------------------------------------------------------------------------
' RETURN ACTION
'------------------------------------------------------------------------------
    'Resolve footer labels
        Select Case LabelName

            Case "Lbl_TodayIcon", _
                 "Lbl_TodayCaption", _
                 "Lbl_Today", _
                 "Lbl_TodayHalo"

                UF_Footer_ActionFromLabelName = "WRITE_TODAY"
                Exit Function

            Case "Lbl_TimeIcon", _
                 "Lbl_TimeCaption", _
                 "Lbl_Time", _
                 "Lbl_TimeHalo"

                UF_Footer_ActionFromLabelName = "WRITE_NOW"
                Exit Function

            Case "Lbl_SettingsIcon", _
                 "Lbl_SettingsHalo"

                UF_Footer_ActionFromLabelName = "SHOW_SETTINGS"
                Exit Function

        End Select

'------------------------------------------------------------------------------
' RETURN FALLBACK
'------------------------------------------------------------------------------
    'Return an empty action for unsupported labels
        UF_Footer_ActionFromLabelName = vbNullString

End Function

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
'   Creates or reuses Fra_PickerPanel and 12 background/text label pairs
'
' ERROR POLICY
'   Raises a descriptive runtime error if layout constants are invalid, if the
'   picker panel cannot be created, or if one or more labels cannot be created
'   or hooked
'
' DEPENDENCIES
'   UF_Ensure_FrameLabel
'   UF_CalendarGrid_GetWidth
'   cDatePickerLabelHook
'
' NOTES
'   Year navigation is handled by the header year-arrow labels
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_PickerPanel_Build"          'Current procedure name

    Dim Fra_PickerPanel        As MSForms.Frame                          'Reusable picker panel
    Dim LabelHook              As cDatePickerLabelHook                   'Runtime click hook
    Dim Index                  As Long                                   'Picker item index
    Dim RowIndex               As Long                                   'Zero-based row index
    Dim ColIndex               As Long                                   'Zero-based column index
    Dim ControlName            As String                                 'Runtime control name
    Dim ItemGridLeft           As Single                                 'Centered item-grid left position
    Dim Lbl_ItemBg             As MSForms.Label                          'Picker item background label
    Dim Lbl_Item               As MSForms.Label                          'Picker item text label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Reset picker-panel hooks
        Set mPickerPanelHooks = New Collection

'------------------------------------------------------------------------------
' VALIDATE CONSTANTS
'------------------------------------------------------------------------------
    'Reject invalid picker-panel width
        If DP_PICKER_PANEL_WIDTH <= 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DP_PICKER_PANEL_WIDTH must be greater than zero."
        End If
    'Reject invalid picker-panel height
        If DP_PICKER_PANEL_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "DP_PICKER_PANEL_HEIGHT must be greater than zero."
        End If
    'Reject invalid picker-item width
        If DP_PICKER_ITEM_WIDTH <= 0 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "DP_PICKER_ITEM_WIDTH must be greater than zero."
        End If
    'Reject invalid picker-item height
        If DP_PICKER_ITEM_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "DP_PICKER_ITEM_HEIGHT must be greater than zero."
        End If
    'Reject invalid picker-item horizontal step
        If DP_PICKER_ITEM_HORIZONTAL_STEP <= 0 Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "DP_PICKER_ITEM_HORIZONTAL_STEP must be greater than zero."
        End If
    'Reject invalid picker-item vertical step
        If DP_PICKER_ITEM_VERTICAL_STEP <= 0 Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "DP_PICKER_ITEM_VERTICAL_STEP must be greater than zero."
        End If

'------------------------------------------------------------------------------
' CREATE / RETRIEVE FRAME
'------------------------------------------------------------------------------
    'Temporarily ignore lookup errors when the frame does not exist yet
        On Error Resume Next
    'Try to retrieve the existing picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)
    'Restore controlled error handling
        On Error GoTo ErrorHandler
    'Create the picker panel if it does not already exist
        If Fra_PickerPanel Is Nothing Then
            Set Fra_PickerPanel = Me.Controls.Add("Forms.Frame.1", DP_PICKER_PANEL_NAME, True)
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
    'Calculate the centered left position of the 3-column item grid
        ItemGridLeft = (Fra_PickerPanel.Width - _
            ((DP_PICKER_ITEM_WIDTH * DP_PICKER_ITEMS_PER_ROW) + _
            ((DP_PICKER_ITEM_HORIZONTAL_STEP - DP_PICKER_ITEM_WIDTH) * _
            (DP_PICKER_ITEMS_PER_ROW - 1)))) / 2

'------------------------------------------------------------------------------
' CREATE 12 BACKGROUND AND TEXT LABEL PAIRS
'------------------------------------------------------------------------------
    'Loop through the picker items
        For Index = 1 To DP_PICKER_ITEM_COUNT

            'Calculate the zero-based row index
                RowIndex = (Index - 1) \ DP_PICKER_ITEMS_PER_ROW

            'Calculate the zero-based column index
                ColIndex = (Index - 1) Mod DP_PICKER_ITEMS_PER_ROW

            'Build the background control name
                ControlName = "Lbl_MonthYearBg" & CStr(Index)

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

            'Build the text control name
                ControlName = "Lbl_MonthYear" & CStr(Index)

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

            'Move the background behind the text label
                Lbl_ItemBg.ZOrder 1

            'Move the text label to the front
                Lbl_Item.ZOrder 0

            'Register the picker-item click hook on the background label
                Set LabelHook = New cDatePickerLabelHook

            'Connect the background label to its action
                LabelHook.Initialize Me, Lbl_ItemBg, "PICKER_ITEM_" & CStr(Index), "PICKER"

            'Store the hook so that the click event remains alive
                mPickerPanelHooks.Add LabelHook, Lbl_ItemBg.Name

            'Register the picker-item click hook on the text label
                Set LabelHook = New cDatePickerLabelHook

            'Connect the text label to its action
                LabelHook.Initialize Me, Lbl_Item, "PICKER_ITEM_" & CStr(Index), "PICKER"

            'Store the hook so that the click event remains alive
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Fra_PickerPanel
'   Lbl_MonthYear1 to Lbl_MonthYear12
'   Lbl_MonthYearBg1 to Lbl_MonthYearBg12
'
' NOTES
'   This routine does not create controls. Controls are created by
'   UF_PickerPanel_Build
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_PickerPanel_EnsureCache"    'Current procedure name

    Dim Index                  As Long                                   'Picker-panel item index
    Dim Fra_PickerPanel        As MSForms.Frame                          'Reusable picker panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RETRIEVE PICKER PANEL
'------------------------------------------------------------------------------
    'Retrieve the picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)

'------------------------------------------------------------------------------
' REBUILD MISSING CACHE REFERENCES
'------------------------------------------------------------------------------
    'Loop through all picker-panel items
        For Index = 1 To DP_PICKER_ITEM_COUNT

            'Resolve the cached picker-panel text label when missing
                If mPickerTextLabels(Index) Is Nothing Then
                    Set mPickerTextLabels(Index) = _
                        Fra_PickerPanel.Controls("Lbl_MonthYear" & CStr(Index))
                End If

            'Resolve the cached picker-panel background label when missing
                If mPickerBackLabels(Index) Is Nothing Then
                    Set mPickerBackLabels(Index) = _
                        Fra_PickerPanel.Controls("Lbl_MonthYearBg" & CStr(Index))
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   terminating
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
'   Suppresses cleanup errors because the form is already terminating
'
' DEPENDENCIES
'   None
'
' NOTES
'   This routine does not delete controls
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index                  As Long                                   'Picker-panel item index

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

    'Loop through all cached picker-panel items
        For Index = 1 To DP_PICKER_ITEM_COUNT

            'Clear the cached picker-panel text label
                Set mPickerTextLabels(Index) = Nothing

            'Clear the cached picker-panel background label
                Set mPickerBackLabels(Index) = Nothing

        Next Index

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
'   Populates month items and shows the picker panel above the day grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if the picker panel or month labels
'   cannot be found or populated
'
' DEPENDENCIES
'   M_Caption_GetMonth
'   UF_PickerPanel_HoverReset
'
' NOTES
'   The current displayed month is highlighted
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_PickerPanel_ShowMonths"     'Current procedure name

    Dim Fra_PickerPanel     As MSForms.Frame        'Reusable picker panel
    Dim Lbl_Item            As MSForms.Label        'Picker item text label
    Dim Lbl_ItemBg          As MSForms.Label        'Picker item background label
    Dim Index               As Long                 'Month index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Ensure cached picker-panel references are available
        UF_PickerPanel_EnsureCache

'------------------------------------------------------------------------------
' RESET HOVER STATE
'------------------------------------------------------------------------------
    'Clear active picker hover before repopulating the panel
        UF_PickerPanel_HoverReset
    'Clear active day-label hover before showing the picker panel
        UF_DayCell_HoverReset
    'Clear active header-label hover before showing the picker panel
        UF_Header_HoverReset

'------------------------------------------------------------------------------
' RETRIEVE PANEL
'------------------------------------------------------------------------------
     'Hide the settings panel before showing the month picker
        UF_SettingsPanel_Hide
    'Retrieve the picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)

'------------------------------------------------------------------------------
' SET PANEL MODE
'------------------------------------------------------------------------------
    'Store month-panel mode
        mPickerPanelMode = 1

'------------------------------------------------------------------------------
' POPULATE MONTHS
'------------------------------------------------------------------------------
    'Loop through the month labels
        For Index = 1 To DP_PICKER_ITEM_COUNT

            'Retrieve the cached picker item text label
                Set Lbl_Item = mPickerTextLabels(Index)
            'Retrieve the cached picker item background label
                Set Lbl_ItemBg = mPickerBackLabels(Index)
            
            'Apply the month caption and value
                With Lbl_Item
                    .Caption = UCase$(Left$(M_Caption_GetMonth(Index, gDP_UseLocalNames), 3))
                    .Tag = CStr(Index)
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

            'Fully reset the picker item background
                With Lbl_ItemBg
                    .Caption = vbNullString
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = DP_PICKER_ITEM_BACK_COLOR
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = DP_PICKER_ITEM_BORDER_COLOR
                    .SpecialEffect = fmSpecialEffectFlat
                    .Enabled = True
                    .Visible = True
                End With

            'Highlight the currently displayed month
                If Index = mDisplayMonth Then

                    'Apply selected-month background formatting
                        With Lbl_ItemBg
                            .BackColor = DP_DAY_SELECTED_BACK_COLOR
                            .BorderColor = DP_DAY_SELECTED_BACK_COLOR
                        End With

                    'Apply selected-month text formatting
                        Lbl_Item.ForeColor = DP_DAY_SELECTED_FORE_COLOR

                End If

            'Move the background behind the text label
                Lbl_ItemBg.ZOrder 1

            'Move the text label to the front
                Lbl_Item.ZOrder 0

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Initializes the visible 12-year range, populates the panel, and shows it
'
' ERROR POLICY
'   Raises a descriptive runtime error if the picker panel cannot be found or
'   populated
'
' DEPENDENCIES
'   UF_PickerPanel_EnsureCache
'   UF_PickerPanel_PopulateYears
'
' NOTES
'   Year navigation arrows are handled by header labels, not by controls inside
'   the picker panel
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_PickerPanel_ShowYears"  'Current procedure name

    Dim Fra_PickerPanel     As MSForms.Frame        'Reusable picker panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Ensure cached picker-panel references are available
        UF_PickerPanel_EnsureCache

'------------------------------------------------------------------------------
' RESET HOVER STATE
'------------------------------------------------------------------------------
    'Clear active picker hover before repopulating the panel
        If mHoveredPickerItemIndex <> 0 Then
            UF_PickerPanel_HoverReset
        End If
    'Clear active day-cell hover before showing the picker panel
        If mHoveredDayCellIndex <> 0 Then
            UF_DayCell_HoverReset
        End If
    'Clear active header hover before showing the picker panel
        If Len(mHoveredHeaderLabelName) <> 0 Then
            UF_Header_HoverReset
        End If

'------------------------------------------------------------------------------
' RETRIEVE PANEL
'------------------------------------------------------------------------------
    'Hide the settings panel before showing the year picker
        UF_SettingsPanel_Hide
    'Retrieve the picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Populates the cached picker-panel text labels and background labels with the
'   12-year range starting from mYearPanelStart
'
' ERROR POLICY
'   Raises a descriptive runtime error if the picker-panel cache cannot be
'   resolved, if mYearPanelStart is invalid, or if one or more cached labels are
'   missing
'
' DEPENDENCIES
'   UF_PickerPanel_EnsureCache
'   mPickerTextLabels
'   mPickerBackLabels
'
' NOTES
'   This routine does not show or hide the panel
'
'   DisplayYear must be recalculated inside the loop. Otherwise all labels keep
'   the default Long value zero
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_PickerPanel_PopulateYears"  'Current procedure name

    Dim Lbl_Item            As MSForms.Label                            'Picker item text label
    Dim Lbl_ItemBg          As MSForms.Label                            'Picker item background label
    Dim Index               As Long                                     'Panel item index
    Dim DisplayYear         As Long                                     'Year displayed in item

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Ensure cached picker-panel references are available
        UF_PickerPanel_EnsureCache

'------------------------------------------------------------------------------
' VALIDATE YEAR RANGE
'------------------------------------------------------------------------------
    'Reject invalid starting years
        If mYearPanelStart < DP_MIN_YEAR Or mYearPanelStart > DP_YEAR_PANEL_MAX_START Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "mYearPanelStart must allow a 12-year range between " & _
                CStr(DP_MIN_YEAR) & " and " & CStr(DP_MAX_YEAR) & "."
        End If

'------------------------------------------------------------------------------
' POPULATE YEARS
'------------------------------------------------------------------------------
    'Loop through the year labels
        For Index = 1 To DP_PICKER_ITEM_COUNT

            'Resolve the year for this picker item
                DisplayYear = mYearPanelStart + Index - 1

            'Retrieve the cached picker item text label
                Set Lbl_Item = mPickerTextLabels(Index)

            'Retrieve the cached picker item background label
                Set Lbl_ItemBg = mPickerBackLabels(Index)

            'Reject a missing cached picker item text label
                If Lbl_Item Is Nothing Then
                    Err.Raise vbObjectError + 514, PROC_NAME, _
                        "Cached picker text label is missing for item " & CStr(Index) & "."
                End If

            'Reject a missing cached picker item background label
                If Lbl_ItemBg Is Nothing Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Cached picker background label is missing for item " & CStr(Index) & "."
                End If

            'Apply the year caption and value
                With Lbl_Item
                    .Caption = CStr(DisplayYear)
                    .Tag = CStr(DisplayYear)
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

            'Fully reset the picker item background
                With Lbl_ItemBg
                    .Caption = vbNullString
                    .BackStyle = fmBackStyleOpaque
                    .BackColor = DP_PICKER_ITEM_BACK_COLOR
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = DP_PICKER_ITEM_BORDER_COLOR
                    .SpecialEffect = fmSpecialEffectFlat
                    .Enabled = True
                    .Visible = True
                End With

            'Highlight the currently displayed year
                If DisplayYear = mDisplayYear Then

                    'Apply selected-year background formatting
                        With Lbl_ItemBg
                            .BackColor = DP_DAY_SELECTED_BACK_COLOR
                            .BorderColor = DP_DAY_SELECTED_BACK_COLOR
                        End With

                    'Apply selected-year text formatting
                        Lbl_Item.ForeColor = DP_DAY_SELECTED_FORE_COLOR

                End If

            'Move the background behind the text label
                Lbl_ItemBg.ZOrder 1

            'Move the text label to the front
                Lbl_Item.ZOrder 0

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub
Public Sub UF_PickerPanel_HandleAction(ByVal ActionName As String)

'
'------------------------------------------------------------------------------
'                           HANDLE PICKER PANEL ACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Handles click actions raised by DatePicker header labels, day labels,
'   footer labels, and picker-panel item labels
'
' WHY THIS EXISTS
'   Runtime-created labels route their Click events through cDatePickerLabelHook
'
' INPUTS
'   ActionName
'     Action name routed by cDatePickerLabelHook
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Routes header navigation, panel display, footer shortcuts, day selection,
'   and picker-panel item selection
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
'   M_DatePolicy_CanSelectDate
'   M_Picker_SelectDate
'   DP_Today
'   DP_Now
'   DP_Close
'
' NOTES
'   This routine is the central action router for runtime-created DatePicker
'   labels
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_PickerPanel_HandleAction"   'Current procedure name
    Const PICKER_ITEM_PREFIX    As String = "PICKER_ITEM_"               'Picker item action prefix
    Const DAY_PICKED_PREFIX     As String = "DAY_PICKED_"                'Day label click action prefix

    Dim ItemIndex               As Long                                  'Clicked item index
    Dim SelectedValue           As Long                                  'Selected month or year value
    Dim RawItemIndex            As String                                'Raw item index text
    Dim RawTagValue             As String                                'Raw label Tag value
    Dim Fra_PickerPanel         As MSForms.Frame                         'Reusable picker panel
    Dim SelectedDate            As Date                                  'Selected date from clicked day label
    
    Dim HandlerStep             As String                                'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Normalize the routed action name
        ActionName = Trim$(ActionName)
    'Reject an empty action name
        If Len(ActionName) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ActionName cannot be empty."
        End If

'------------------------------------------------------------------------------
' ROUTE DIRECT ACTIONS
'------------------------------------------------------------------------------
    'Route the requested direct action
        Select Case ActionName

            Case "PREV_MONTH"
                'Move to the previous displayed month
                    UF_Header_MoveMonth -1
                'Exit after routing the action
                    Exit Sub

            Case "NEXT_MONTH"
                'Move to the next displayed month
                    UF_Header_MoveMonth 1
                'Exit after routing the action
                    Exit Sub

            Case "PREV_YEAR"
                'Move or scroll to the previous year range
                    UF_Header_MoveYear -1
                'Exit after routing the action
                    Exit Sub

            Case "NEXT_YEAR"
                'Move or scroll to the next year range
                    UF_Header_MoveYear 1
                'Exit after routing the action
                    Exit Sub

            Case "SHOW_MONTH_PANEL"
                'Show the month picker panel
                    UF_PickerPanel_ShowMonths
                'Exit after routing the action
                    Exit Sub

            Case "SHOW_YEAR_PANEL"
                'Show the year picker panel
                    UF_PickerPanel_ShowYears
                'Exit after routing the action
                    Exit Sub

            Case "WRITE_NOW"
                'Write today with the current system time
                    DP_Now
                'Exit after routing the action
                    Exit Sub

            Case "WRITE_TODAY"
                'Write today without time
                    DP_Today
                'Exit after routing the action
                    Exit Sub
            
            Case "SHOW_SETTINGS"
                'Show the settings panel
                    UF_Settings_Show
                'Exit after routing the action
                    Exit Sub
            
            Case "HIDE_SETTINGS"
                'Hide the settings panel
                    UF_SettingsPanel_Hide
                'Exit after routing the action
                    Exit Sub
            
            Case "SAVE_SETTINGS"
                'Save settings from the settings panel
                    UF_SettingsPanel_Save
                'Exit after routing the action
                    Exit Sub
            
            Case "SHOW_SETTINGS_DISPLAY"
                'Show the Display Settings page
                    UF_SettingsPanel_SelectPage 0
                'Exit after routing the action
                    Exit Sub

            Case "SHOW_SETTINGS_BEHAVIOR"
                'Show the Behavior Settings page
                    UF_SettingsPanel_SelectPage 1
                'Exit after routing the action
                    Exit Sub

            Case "SHOW_SETTINGS_INTEGRATION"
                'Show the Integration Settings page
                    UF_SettingsPanel_SelectPage 2
                'Exit after routing the action
                    Exit Sub
                    

        
        End Select

'------------------------------------------------------------------------------
' HANDLE DAY LABEL CLICK
'------------------------------------------------------------------------------
    'Handle day-label click actions
        If Left$(ActionName, Len(DAY_PICKED_PREFIX)) = DAY_PICKED_PREFIX Then
            'Track the current handler step
                HandlerStep = "Extract day index"
            'Extract the raw day-cell index
                RawItemIndex = Mid$(ActionName, Len(DAY_PICKED_PREFIX) + 1)
            'Reject non-numeric day-cell indexes
                If Not IsNumeric(RawItemIndex) Then
                    Err.Raise vbObjectError + 514, PROC_NAME, _
                        "Day cell index must be numeric."
                End If
            'Track the current handler step
                HandlerStep = "Parse day index"
            'Parse the clicked day-cell index
                ItemIndex = CLng(RawItemIndex)
            'Reject invalid day-cell indexes
                If ItemIndex < 1 Or ItemIndex > DP_DAY_LABEL_COUNT Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Day cell index must be between 1 and " & CStr(DP_DAY_LABEL_COUNT) & "."
                End If
            'Track the current handler step
                HandlerStep = "Validate cached day-cell date"
            'Reject day cells with no cached date
                If Not mDayCellHasDate(ItemIndex) Then
                    Err.Raise vbObjectError + 516, PROC_NAME, _
                        "Clicked day cell does not have a cached date."
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
        If Left$(ActionName, Len(PICKER_ITEM_PREFIX)) <> PICKER_ITEM_PREFIX Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "Unsupported DatePicker action: " & ActionName
        End If

    'Extract the raw picker-item index
        RawItemIndex = Mid$(ActionName, Len(PICKER_ITEM_PREFIX) + 1)

    'Reject non-numeric picker-item indexes
        If Not IsNumeric(RawItemIndex) Then
            Err.Raise vbObjectError + 519, PROC_NAME, _
                "Picker item index must be numeric."
        End If

    'Parse the picker-item index
        ItemIndex = CLng(RawItemIndex)

    'Reject invalid picker item indexes
        If ItemIndex < 1 Or ItemIndex > DP_PICKER_ITEM_COUNT Then
            Err.Raise vbObjectError + 520, PROC_NAME, _
                "Picker item index must be between 1 and " & CStr(DP_PICKER_ITEM_COUNT) & "."
        End If

    'Retrieve the picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)

'------------------------------------------------------------------------------
' APPLY MONTH SELECTION
'------------------------------------------------------------------------------
    'Handle month picker item selection
        If mPickerPanelMode = 1 Then

            'Store the selected month
                mDisplayMonth = ItemIndex

            'Initialize keyboard date to the first day of the selected display month
                mKeyboardDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
                mHasKeyboardDate = True

            'Hide the picker panel
                Fra_PickerPanel.Visible = False
                mPickerPanelMode = 0

            'Refresh the calendar grid
                UF_DayGrid_Populate mDisplayYear, mDisplayMonth

            'Exit after applying month selection
                Exit Sub

        End If

'------------------------------------------------------------------------------
' APPLY YEAR SELECTION
'------------------------------------------------------------------------------
    'Handle year picker item selection
        If mPickerPanelMode = 2 Then

            'Ensure cached picker-panel references are available
                UF_PickerPanel_EnsureCache
            'Read the selected year value from the cached picker item Tag
                RawTagValue = CStr(mPickerTextLabels(ItemIndex).Tag)

            'Reject an empty year picker Tag
                If Len(Trim$(RawTagValue)) = 0 Then
                    Err.Raise vbObjectError + 521, PROC_NAME, _
                        "Selected year picker item does not contain a value."
                End If
            'Reject non-numeric year picker Tag values
                If Not IsNumeric(RawTagValue) Then
                    Err.Raise vbObjectError + 522, PROC_NAME, _
                        "Selected year picker item Tag must contain a numeric year."
                End If
            'Parse the selected year
                SelectedValue = CLng(RawTagValue)
            'Reject unsupported selected years
                If SelectedValue < DP_MIN_YEAR Or SelectedValue > DP_MAX_YEAR Then
                    Err.Raise vbObjectError + 523, PROC_NAME, _
                        "Selected year must be between " & CStr(DP_MIN_YEAR) & " and " & CStr(DP_MAX_YEAR) & "."
                End If

            'Store the selected year
                mDisplayYear = SelectedValue

            'Initialize keyboard date to the first day of the selected display period
                mKeyboardDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
                mHasKeyboardDate = True

            'Hide the picker panel
                Fra_PickerPanel.Visible = False
                mPickerPanelMode = 0

            'Refresh the calendar grid
                UF_DayGrid_Populate mDisplayYear, mDisplayMonth

            'Exit after applying year selection
                Exit Sub

        End If

'------------------------------------------------------------------------------
' REJECT INVALID PANEL MODE
'------------------------------------------------------------------------------
    'Reject unsupported picker panel modes
        Err.Raise vbObjectError + 524, PROC_NAME, _
            "mPickerPanelMode must be 1 for months or 2 for years."

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
            PROC_NAME & " | Action=" & ActionName & " | Step=" & HandlerStep, _
            Err.Description

End Sub

Private Sub UF_Settings_Show()

'
'------------------------------------------------------------------------------
'                           SHOW SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Shows the in-form DatePicker settings panel
'
' WHY THIS EXISTS
'   The DatePicker footer exposes a Settings entry point. Settings should be
'   displayed inside the DatePicker itself instead of opening a separate modal
'   form
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Shows Fra_Settings using the same overlay geometry as Fra_PickerPanel
'
' ERROR POLICY
'   Raises a descriptive runtime error if the settings panel cannot be shown
'
' DEPENDENCIES
'   UF_SettingsPanel_Show
'
' NOTES
'   This replaces the previous Frm_DpSettings.Show vbModal behavior
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_Settings_Show"  'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Creates or reuses Fra_Settings and its child controls:
'     - title label
'     - save label
'     - close label
'     - fake tab labels
'     - runtime MultiPage
'     - Display Settings page
'     - Behavior Settings page
'     - Integration Settings page
'
'   The Display Settings page contains:
'     - one ComboBox for first-day-of-week selection
'     - four CheckBox controls:
'         - Use local names
'         - Live clock
'         - Compact layout
'         - Highlight weekends
'
'   The Behavior Settings page contains:
'     - Allow outside-month selection
'     - Close after selection
'
'   The Integration Settings page contains:
'     - Right-click menu
'     - In-grid icon
'     - Use WinAPI styling
'
' ERROR POLICY
'   Raises a descriptive runtime error if the panel, MultiPage, pages, labels,
'   ComboBox, or CheckBox controls cannot be created or formatted
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
'   CheckBox formatting is delegated to UF_SettingsCheckBox_Format so each
'   CheckBox receives a clean independent font object
'
'   Runtime ComboBox and CheckBox change persistence should be handled by a
'   dedicated runtime input-hook class in the next step
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "UF_SettingsPanel_Build"        'Current procedure name

    Dim Fra_Settings            As MSForms.Frame                           'Reusable settings panel
    Dim Mp_Settings             As MSForms.MultiPage                       'Settings MultiPage control
    Dim Pge_Display             As MSForms.Page                            'Display settings page
    Dim Pge_Behavior            As MSForms.Page                            'Behavior settings page
    Dim Pge_Integration         As MSForms.Page                            'Integration settings page

    Dim PanelLeft               As Single                                  'Settings panel left position
    Dim PanelTop                As Single                                  'Settings panel top position

    Dim Lbl_Title               As MSForms.Label                           'Settings title label
    Dim Lbl_Save                As MSForms.Label                           'Settings save label
    Dim Lbl_Close               As MSForms.Label                           'Settings close label
    Dim Lbl_FirstDay            As MSForms.Label                           'First-day label

    Dim Cbo_FirstDay            As MSForms.ComboBox                        'First-day ComboBox

    Dim Chk_LocalNames          As MSForms.CheckBox                        'Use-local-names checkbox
    Dim Chk_LiveClock           As MSForms.CheckBox                        'Live-clock checkbox
    Dim Chk_Compact             As MSForms.CheckBox                        'Compact-layout checkbox
    Dim Chk_Weekends            As MSForms.CheckBox                        'Highlight-weekends checkbox

    Dim Chk_AllowOutside        As MSForms.CheckBox                        'Allow outside-month selection checkbox
    Dim Chk_CloseAfter          As MSForms.CheckBox                        'Close-after-selection checkbox

    Dim Chk_RightClick          As MSForms.CheckBox                        'Right-click menu checkbox
    Dim Chk_InGridIcon          As MSForms.CheckBox                        'In-grid icon checkbox
    Dim Chk_WinAPIStyle         As MSForms.CheckBox                        'WinAPI styling checkbox

    Dim LabelHook               As cDatePickerLabelHook                    'Runtime label hook
    Dim TitleFont               As Object                                  'Clean title font object
    Dim BodyFont                As Object                                  'Clean body font object
    Dim IconFont                As Object                                  'Clean settings icon font object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Create or reset settings-panel hooks
        Set mSettingsPanelHooks = New Collection

'------------------------------------------------------------------------------
' VALIDATE CONSTANTS
'------------------------------------------------------------------------------
    'Reject invalid settings panel width
        If DP_PICKER_PANEL_WIDTH <= 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DP_PICKER_PANEL_WIDTH must be greater than zero."
        End If

    'Reject invalid settings panel height
        If DP_PICKER_PANEL_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "DP_PICKER_PANEL_HEIGHT must be greater than zero."
        End If

    'Reject invalid MultiPage dimensions
        If DP_SETTINGS_MULTIPAGE_WIDTH <= 0 Or DP_SETTINGS_MULTIPAGE_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Settings MultiPage width and height must be greater than zero."
        End If

    'Reject invalid first-day ComboBox dimensions
        If DP_SETTINGS_FIRSTDAY_COMBO_WIDTH <= 0 Or DP_SETTINGS_FIRSTDAY_COMBO_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "First-day ComboBox width and height must be greater than zero."
        End If

    'Reject invalid checkbox dimensions
        If DP_SETTINGS_CHECKBOX_WIDTH <= 0 Or DP_SETTINGS_CHECKBOX_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Settings checkbox width and height must be greater than zero."
        End If

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
    'Temporarily ignore lookup errors when the frame does not exist yet
        On Error Resume Next

    'Try to retrieve the existing settings panel
        Set Fra_Settings = Me.Controls(DP_SETTINGS_PANEL_NAME)

    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Create the settings panel if it does not already exist
        If Fra_Settings Is Nothing Then
            Set Fra_Settings = Me.Controls.Add("Forms.Frame.1", DP_SETTINGS_PANEL_NAME, True)
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
' HIDE LEGACY SETTINGS LABELS
'------------------------------------------------------------------------------
    'Temporarily ignore missing legacy labels
        On Error Resume Next

    'Hide the legacy first option label if present
        Fra_Settings.Controls("Lbl_SettingsOption1").Visible = False

    'Hide the legacy second option label if present
        Fra_Settings.Controls("Lbl_SettingsOption2").Visible = False

    'Hide the legacy third option label if present
        Fra_Settings.Controls("Lbl_SettingsOption3").Visible = False

    'Hide the legacy fourth option label if present
        Fra_Settings.Controls("Lbl_SettingsOption4").Visible = False

    'Hide the legacy hint label if present
        Fra_Settings.Controls("Lbl_SettingsHint").Visible = False

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
            .ControlTipText = DP_SETTINGS_SAVE_TOOLTIP
        End With

    'Create a clean settings icon font object
        Set IconFont = CreateObject("StdFont")

    'Configure the settings save icon font
        With IconFont
            .Name = DP_SETTINGS_HEADER_ICON_FONT_NAME
            .Size = DP_SETTINGS_SAVE_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean icon font to the settings save label
        Set Lbl_Save.Font = IconFont

    'Apply the Segoe MDL2 save glyph
        Lbl_Save.Caption = ChrW$(DP_SETTINGS_SAVE_CODEPOINT)

    'Create the settings save click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the save label to the save-settings action
        LabelHook.Initialize Me, Lbl_Save, "SAVE_SETTINGS"

    'Store the hook so that the click event remains alive
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
            .ControlTipText = DP_SETTINGS_CLOSE_TOOLTIP
        End With

    'Create a clean settings icon font object
        Set IconFont = CreateObject("StdFont")

    'Configure the settings close icon font
        With IconFont
            .Name = DP_SETTINGS_HEADER_ICON_FONT_NAME
            .Size = DP_SETTINGS_CLOSE_FONT_SIZE
            .Bold = False
            .Italic = False
            .Underline = False
            .Strikethrough = False
        End With

    'Assign the clean icon font to the settings close label
        Set Lbl_Close.Font = IconFont

    'Apply the Segoe MDL2 close glyph
        Lbl_Close.Caption = ChrW$(DP_SETTINGS_CLOSE_CODEPOINT)

    'Create the settings close click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the close label to the hide-settings action
        LabelHook.Initialize Me, Lbl_Close, "HIDE_SETTINGS"

    'Store the hook so that the click event remains alive
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
    'Retrieve the Display Settings page
        Set Pge_Display = Mp_Settings.Pages(0)

    'Retrieve the Behavior Settings page
        Set Pge_Behavior = Mp_Settings.Pages(1)

    'Retrieve the Integration Settings page
        Set Pge_Integration = Mp_Settings.Pages(2)

'------------------------------------------------------------------------------
' FORMAT MULTIPAGE PAGE CAPTIONS
'------------------------------------------------------------------------------
    'Apply the Display Settings page caption
        Pge_Display.Caption = DP_SETTINGS_PAGE_DISPLAY_CAPTION

    'Apply the Behavior Settings page caption
        Pge_Behavior.Caption = DP_SETTINGS_PAGE_BEHAVIOR_CAPTION

    'Apply the Integration Settings page caption
        Pge_Integration.Caption = DP_SETTINGS_PAGE_INTEGRATION_CAPTION

'------------------------------------------------------------------------------
' CREATE / FORMAT FAKE SETTINGS TABS
'------------------------------------------------------------------------------
    'Create the label-based settings tab strip
        UF_SettingsTabs_Build Fra_Settings, BodyFont

'------------------------------------------------------------------------------
' BUILD PAGE BACKGROUNDS
'------------------------------------------------------------------------------
    'Build the Display Settings page background
        UF_SettingsPage_BackgroundBuild Mp_Settings, 0, "Lbl_DisplaySettingsBack"

    'Build the Behavior Settings page background
        UF_SettingsPage_BackgroundBuild Mp_Settings, 1, "Lbl_BehaviorSettingsBack"

    'Build the Integration Settings page background
        UF_SettingsPage_BackgroundBuild Mp_Settings, 2, "Lbl_IntegrationSettingsBack"

'------------------------------------------------------------------------------
' HIDE LEGACY PAGE CONTROLS
'------------------------------------------------------------------------------
    'Temporarily ignore missing legacy controls
        On Error Resume Next

    'Hide the obsolete borderless-window checkbox when present
        Pge_Display.Controls("Chk_SettingsBorderlessWindow").Visible = False

    'Hide the old behavior placeholder when present
        Pge_Behavior.Controls("Lbl_SettingsBehaviorPlaceholder").Visible = False

    'Hide the old integration placeholder when present
        Pge_Integration.Controls("Lbl_SettingsIntegrationPlaceholder").Visible = False

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

    'Refresh the fake settings tab visual state
        UF_SettingsTabs_RefreshVisualState Mp_Settings.Value

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   picker panel
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Hides the picker panel, resets transient hover state, refreshes settings
'   captions, and shows Fra_Settings
'
' ERROR POLICY
'   Raises a descriptive runtime error if the settings panel cannot be shown
'
' DEPENDENCIES
'   UF_SettingsPanel_Build
'   UF_SettingsPanel_RefreshCaptions
'
' NOTES
'   This routine does not open a separate UserForm
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_SettingsPanel_Show"         'Current procedure name

    Dim Fra_Settings           As MSForms.Frame                           'Reusable settings panel
    Dim Fra_PickerPanel        As MSForms.Frame                           'Reusable picker panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' ENSURE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Temporarily ignore lookup errors when the settings panel does not exist yet
        On Error Resume Next

    'Try to retrieve the existing settings panel
        Set Fra_Settings = Me.Controls(DP_SETTINGS_PANEL_NAME)

    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Build the settings panel if it has not been created yet
        If Fra_Settings Is Nothing Then
            UF_SettingsPanel_Build
            Set Fra_Settings = Me.Controls(DP_SETTINGS_PANEL_NAME)
        End If

'------------------------------------------------------------------------------
' RESET TRANSIENT HOVER STATE
'------------------------------------------------------------------------------
    'Clear active picker hover before showing settings
        If mHoveredPickerItemIndex <> 0 Then
            UF_PickerPanel_HoverReset
        End If
    'Clear active day-cell hover before showing settings
        If mHoveredDayCellIndex <> 0 Then
            UF_DayCell_HoverReset
        End If
    'Clear active header hover before showing settings
        If Len(mHoveredHeaderLabelName) <> 0 Then
            UF_Header_HoverReset
        End If
    'Clear active footer hover before showing settings
        If Len(mHoveredFooterActionName) <> 0 Then
            UF_Footer_HoverReset
        End If

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Temporarily ignore lookup errors when the picker panel is not available
        On Error Resume Next
    'Retrieve the picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)
    'Restore controlled error handling
        On Error GoTo ErrorHandler
    'Hide the picker panel when available
        If Not Fra_PickerPanel Is Nothing Then
            Fra_PickerPanel.Visible = False
        End If
    'Clear picker-panel mode
        mPickerPanelMode = 0

'------------------------------------------------------------------------------
' REFRESH SETTINGS CAPTIONS
'------------------------------------------------------------------------------
    'Refresh the displayed settings values
        UF_SettingsPanel_RefreshCaptions
    'Refresh the fake settings tab visual state
        UF_SettingsTabs_RefreshVisualState
        
'------------------------------------------------------------------------------
' SHOW SETTINGS PANEL
'------------------------------------------------------------------------------
    'Show the settings panel
        Fra_Settings.Visible = True

    'Move the settings panel in front of the day grid
        Fra_Settings.ZOrder 0

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

Private Sub UF_SettingsPanel_Save()

'
'------------------------------------------------------------------------------
'                           SAVE SETTINGS PANEL
'------------------------------------------------------------------------------
' PURPOSE
'   Handles the settings-panel Save action
'
' WHY THIS EXISTS
'   The settings panel exposes an explicit Save icon in its header
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Placeholder for settings persistence logic
'
' ERROR POLICY
'   Raises a descriptive runtime error if saving fails
'
' DEPENDENCIES
'   None
'
' NOTES
'   Wire this routine to your registry-backed settings save logic when the input
'   hooks are added
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "UF_SettingsPanel_Save"     'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SAVE SETTINGS
'------------------------------------------------------------------------------
    'Placeholder for settings persistence
        Beep

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
'   DatePicker grid
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Hides Fra_Settings when it exists
'
' ERROR POLICY
'   Best-effort cleanup. Missing panel is a safe no-op
'
' DEPENDENCIES
'   None
'
' NOTES
'   This routine does not clear saved settings
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Fra_Settings           As MSForms.Frame      'Reusable settings panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress hide errors
        On Error Resume Next

'------------------------------------------------------------------------------
' HIDE PANEL
'------------------------------------------------------------------------------
    'Retrieve the settings panel
        Set Fra_Settings = Me.Controls(DP_SETTINGS_PANEL_NAME)

    'Hide the settings panel when available
        If Not Fra_Settings Is Nothing Then
            Fra_Settings.Visible = False
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub

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
'   Creates or reuses three tab labels:
'     - Display
'     - Behavior
'     - Integration
'
'   Each label routes a click action that changes the hidden MultiPage page
'
' ERROR POLICY
'   Raises a descriptive runtime error if the frame, font, labels, or hooks
'   cannot be created
'
' DEPENDENCIES
'   UF_Ensure_FrameLabel
'   UF_SettingsTab_ApplyVisualState
'   cDatePickerLabelHook
'
' NOTES
'   The native MultiPage tab strip must remain hidden through fmTabStyleNone
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_SettingsTabs_Build"         'Current procedure name

    Dim LabelHook              As cDatePickerLabelHook                    'Runtime tab click hook
    Dim Lbl_TabDisplay         As MSForms.Label                           'Display tab label
    Dim Lbl_TabBehavior        As MSForms.Label                           'Behavior tab label
    Dim Lbl_TabIntegration     As MSForms.Label                           'Integration tab label

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
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentFrame cannot be Nothing."
        End If

    'Reject a missing body font
        If BodyFont Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, "BodyFont cannot be Nothing."
        End If

'------------------------------------------------------------------------------
' CREATE / FORMAT DISPLAY TAB
'------------------------------------------------------------------------------
    'Create or retrieve the Display tab label
        Set Lbl_TabDisplay = UF_Ensure_FrameLabel(ParentFrame, DP_SETTINGS_TAB_DISPLAY_NAME)

    'Apply layout and visual properties to the Display tab
        With Lbl_TabDisplay
            .Caption = DP_SETTINGS_TAB_DISPLAY_CAPTION
            .Left = DP_SETTINGS_TAB_LEFT
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
            .ControlTipText = DP_SETTINGS_PAGE_DISPLAY_CAPTION
        End With

    'Assign the clean body font to the Display tab
        Set Lbl_TabDisplay.Font = BodyFont

    'Create the Display tab click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the Display tab to the Display Settings page
        LabelHook.Initialize Me, Lbl_TabDisplay, "SHOW_SETTINGS_DISPLAY"

    'Store the hook so that the click event remains alive
        mSettingsPanelHooks.Add LabelHook, Lbl_TabDisplay.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT BEHAVIOR TAB
'------------------------------------------------------------------------------
    'Create or retrieve the Behavior tab label
        Set Lbl_TabBehavior = UF_Ensure_FrameLabel(ParentFrame, DP_SETTINGS_TAB_BEHAVIOR_NAME)

    'Apply layout and visual properties to the Behavior tab
        With Lbl_TabBehavior
            .Caption = DP_SETTINGS_TAB_BEHAVIOR_CAPTION
            .Left = DP_SETTINGS_TAB_LEFT + DP_SETTINGS_TAB_DISPLAY_WIDTH + DP_SETTINGS_TAB_GAP
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
            .ControlTipText = DP_SETTINGS_PAGE_BEHAVIOR_CAPTION
        End With

    'Assign the clean body font to the Behavior tab
        Set Lbl_TabBehavior.Font = BodyFont

    'Create the Behavior tab click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the Behavior tab to the Behavior Settings page
        LabelHook.Initialize Me, Lbl_TabBehavior, "SHOW_SETTINGS_BEHAVIOR"

    'Store the hook so that the click event remains alive
        mSettingsPanelHooks.Add LabelHook, Lbl_TabBehavior.Name

'------------------------------------------------------------------------------
' CREATE / FORMAT INTEGRATION TAB
'------------------------------------------------------------------------------
    'Create or retrieve the Integration tab label
        Set Lbl_TabIntegration = UF_Ensure_FrameLabel(ParentFrame, DP_SETTINGS_TAB_INTEGRATION_NAME)

    'Apply layout and visual properties to the Integration tab
        With Lbl_TabIntegration
            .Caption = DP_SETTINGS_TAB_INTEGRATION_CAPTION
            .Left = DP_SETTINGS_TAB_LEFT + DP_SETTINGS_TAB_DISPLAY_WIDTH + DP_SETTINGS_TAB_BEHAVIOR_WIDTH + (DP_SETTINGS_TAB_GAP * 2)
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
            .ControlTipText = DP_SETTINGS_PAGE_INTEGRATION_CAPTION
        End With

    'Assign the clean body font to the Integration tab
        Set Lbl_TabIntegration.Font = BodyFont

    'Create the Integration tab click hook
        Set LabelHook = New cDatePickerLabelHook

    'Connect the Integration tab to the Integration Settings page
        LabelHook.Initialize Me, Lbl_TabIntegration, "SHOW_SETTINGS_INTEGRATION"

    'Store the hook so that the click event remains alive
        mSettingsPanelHooks.Add LabelHook, Lbl_TabIntegration.Name

'------------------------------------------------------------------------------
' APPLY INITIAL VISUAL STATE
'------------------------------------------------------------------------------
    'Apply selected state to the Display tab
        UF_SettingsTab_ApplyVisualState Lbl_TabDisplay, True

    'Apply normal state to the Behavior tab
        UF_SettingsTab_ApplyVisualState Lbl_TabBehavior, False

    'Apply normal state to the Integration tab
        UF_SettingsTab_ApplyVisualState Lbl_TabIntegration, False

'------------------------------------------------------------------------------
' RESTORE LAYERING
'------------------------------------------------------------------------------
    'Move the Display tab to the front
        Lbl_TabDisplay.ZOrder 0

    'Move the Behavior tab to the front
        Lbl_TabBehavior.ZOrder 0

    'Move the Integration tab to the front
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Validates the requested page, changes Mp_Settings.Value, and refreshes the
'   visual state of the fake tabs
'
' ERROR POLICY
'   Raises a descriptive runtime error if the settings panel, MultiPage, or page
'   index is invalid
'
' DEPENDENCIES
'   UF_SettingsTabs_RefreshVisualState
'
' NOTES
'   This routine does not rebuild the settings panel
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_SettingsPanel_SelectPage"    'Current procedure name

    Dim Fra_Settings           As MSForms.Frame                            'Reusable settings panel
    Dim Mp_Settings            As MSForms.MultiPage                        'Settings MultiPage

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject unsupported page indexes
        If PageIndex < 0 Or PageIndex > 2 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "PageIndex must be between 0 and 2."
        End If

'------------------------------------------------------------------------------
' RETRIEVE CONTROLS
'------------------------------------------------------------------------------
    'Retrieve the settings panel
        Set Fra_Settings = Me.Controls(DP_SETTINGS_PANEL_NAME)

    'Retrieve the settings MultiPage
        Set Mp_Settings = Fra_Settings.Controls(DP_SETTINGS_MULTIPAGE_NAME)

'------------------------------------------------------------------------------
' VALIDATE MULTIPAGE STATE
'------------------------------------------------------------------------------
    'Reject page indexes outside the actual MultiPage page count
        If PageIndex > Mp_Settings.Pages.Count - 1 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "PageIndex exceeds the current MultiPage page count."
        End If

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Applies selected formatting to the active tab and normal formatting to the
'   remaining tabs
'
' ERROR POLICY
'   Raises a descriptive runtime error if required controls cannot be resolved
'
' DEPENDENCIES
'   UF_SettingsTab_ApplyVisualState
'
' NOTES
'   This routine changes only visual state
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_SettingsTabs_RefreshVisualState" 'Current procedure name

    Dim Fra_Settings           As MSForms.Frame                               'Reusable settings panel
    Dim Mp_Settings            As MSForms.MultiPage                           'Settings MultiPage

    Dim Lbl_TabDisplay         As MSForms.Label                               'Display tab label
    Dim Lbl_TabBehavior        As MSForms.Label                               'Behavior tab label
    Dim Lbl_TabIntegration     As MSForms.Label                               'Integration tab label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RETRIEVE CONTROLS
'------------------------------------------------------------------------------
    'Retrieve the settings panel
        Set Fra_Settings = Me.Controls(DP_SETTINGS_PANEL_NAME)

    'Retrieve the settings MultiPage
        Set Mp_Settings = Fra_Settings.Controls(DP_SETTINGS_MULTIPAGE_NAME)

    'Resolve the active page index when not supplied
        If ActivePageIndex < 0 Then
            ActivePageIndex = Mp_Settings.Value
        End If

    'Reject unsupported active page indexes
        If ActivePageIndex < 0 Or ActivePageIndex > 2 Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "ActivePageIndex must be between 0 and 2."
        End If

    'Retrieve the Display tab label
        Set Lbl_TabDisplay = Fra_Settings.Controls(DP_SETTINGS_TAB_DISPLAY_NAME)

    'Retrieve the Behavior tab label
        Set Lbl_TabBehavior = Fra_Settings.Controls(DP_SETTINGS_TAB_BEHAVIOR_NAME)

    'Retrieve the Integration tab label
        Set Lbl_TabIntegration = Fra_Settings.Controls(DP_SETTINGS_TAB_INTEGRATION_NAME)

'------------------------------------------------------------------------------
' APPLY TAB VISUAL STATE
'------------------------------------------------------------------------------
    'Refresh the Display tab visual state
        UF_SettingsTab_ApplyVisualState Lbl_TabDisplay, (ActivePageIndex = 0)

    'Refresh the Behavior tab visual state
        UF_SettingsTab_ApplyVisualState Lbl_TabBehavior, (ActivePageIndex = 1)

    'Refresh the Integration tab visual state
        UF_SettingsTab_ApplyVisualState Lbl_TabIntegration, (ActivePageIndex = 2)

'------------------------------------------------------------------------------
' RESTORE LAYERING
'------------------------------------------------------------------------------
    'Move the Display tab to the front
        Lbl_TabDisplay.ZOrder 0

    'Move the Behavior tab to the front
        Lbl_TabBehavior.ZOrder 0

    'Move the Integration tab to the front
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Raises a descriptive runtime error if TabLabel is missing
'
' DEPENDENCIES
'   None
'
' NOTES
'   This routine does not change the MultiPage value
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_SettingsTab_ApplyVisualState" 'Current procedure name

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing tab label
        If TabLabel Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "TabLabel cannot be Nothing."
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
    'Apply selected-tab colors
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
'   gDP_AllowOutsideMonthSel
'   gDP_CloseAfterSelection
'   gDP_EnableRightClickMenu
'   gDP_ShowInGridIcon
'   gDP_UseWinAPI
'   M_Settings_IsValidFirstDayOfWeek
'
' NOTES
'   This routine displays current values only
'
'   Runtime input-event persistence should be added through a dedicated
'   ComboBox / CheckBox hook class
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_SettingsPanel_RefreshCaptions" 'Current procedure name

    Dim Fra_Settings           As MSForms.Frame                              'Reusable settings panel
    Dim Mp_Settings            As MSForms.MultiPage                          'Settings MultiPage
    Dim Pge_Display            As MSForms.Page                               'Display settings page
    Dim Pge_Behavior           As MSForms.Page                               'Behavior settings page
    Dim Pge_Integration        As MSForms.Page                               'Integration settings page

    Dim Cbo_FirstDay           As MSForms.ComboBox                           'First-day ComboBox

    Dim Chk_LocalNames         As MSForms.CheckBox                           'Use-local-names checkbox
    Dim Chk_LiveClock          As MSForms.CheckBox                           'Live-clock checkbox
    Dim Chk_Compact            As MSForms.CheckBox                           'Compact-layout checkbox
    Dim Chk_Weekends           As MSForms.CheckBox                           'Highlight-weekends checkbox

    Dim Chk_AllowOutside       As MSForms.CheckBox                           'Allow outside-month selection checkbox
    Dim Chk_CloseAfter         As MSForms.CheckBox                           'Close-after-selection checkbox

    Dim Chk_RightClick         As MSForms.CheckBox                           'Right-click menu checkbox
    Dim Chk_InGridIcon         As MSForms.CheckBox                           'In-grid icon checkbox
    Dim Chk_WinAPIStyle        As MSForms.CheckBox                           'WinAPI styling checkbox

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
                "gDP_FirstDayOfWeek must be vbSunday or vbMonday."
        End If

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS PANEL
'------------------------------------------------------------------------------
    'Retrieve the settings panel
        Set Fra_Settings = Me.Controls(DP_SETTINGS_PANEL_NAME)

    'Retrieve the settings MultiPage
        Set Mp_Settings = Fra_Settings.Controls(DP_SETTINGS_MULTIPAGE_NAME)

    'Reject an incomplete settings MultiPage
        If Mp_Settings.Pages.Count < 3 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Settings MultiPage must contain at least three pages."
        End If

'------------------------------------------------------------------------------
' RETRIEVE SETTINGS PAGES
'------------------------------------------------------------------------------
    'Retrieve the Display Settings page
        Set Pge_Display = Mp_Settings.Pages(0)

    'Retrieve the Behavior Settings page
        Set Pge_Behavior = Mp_Settings.Pages(1)

    'Retrieve the Integration Settings page
        Set Pge_Integration = Mp_Settings.Pages(2)

'------------------------------------------------------------------------------
' RETRIEVE DISPLAY CONTROLS
'------------------------------------------------------------------------------
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
    'Retrieve the allow-outside-month checkbox
        Set Chk_AllowOutside = Pge_Behavior.Controls("Chk_SettingsAllowOutsideMonth")

    'Retrieve the close-after-selection checkbox
        Set Chk_CloseAfter = Pge_Behavior.Controls("Chk_SettingsCloseAfterSelection")

'------------------------------------------------------------------------------
' RETRIEVE INTEGRATION CONTROLS
'------------------------------------------------------------------------------
    'Retrieve the right-click menu checkbox
        Set Chk_RightClick = Pge_Integration.Controls("Chk_SettingsRightClickMenu")

    'Retrieve the in-grid icon checkbox
        Set Chk_InGridIcon = Pge_Integration.Controls("Chk_SettingsInGridIcon")

    'Retrieve the WinAPI styling checkbox
        Set Chk_WinAPIStyle = Pge_Integration.Controls("Chk_SettingsWinAPIStyling")

'------------------------------------------------------------------------------
' REFRESH FIRST-DAY COMBOBOX
'------------------------------------------------------------------------------
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
    'Refresh the allow-outside-month checkbox value
        Chk_AllowOutside.Value = CBool(gDP_AllowOutsideMonthSel)

    'Refresh the close-after-selection checkbox value
        Chk_CloseAfter.Value = CBool(gDP_CloseAfterSelection)

'------------------------------------------------------------------------------
' REFRESH INTEGRATION CHECKBOX VALUES
'------------------------------------------------------------------------------
    'Refresh the right-click menu checkbox value
        Chk_RightClick.Value = CBool(gDP_EnableRightClickMenu)

    'Refresh the in-grid icon checkbox value
        Chk_InGridIcon.Value = CBool(gDP_ShowInGridIcon)

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Clears the caption, assigns layout and visual state, assigns a fresh font,
'   then restores the caption
'
' ERROR POLICY
'   Raises a descriptive runtime error if TargetCheckBox is missing
'
' DEPENDENCIES
'   UF_SettingsCheckBoxFont_Create
'   DP_SETTINGS_CHECKBOX_WIDTH
'   DP_SETTINGS_CHECKBOX_HEIGHT
'
' NOTES
'   The caption is assigned after the fresh font is assigned
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_SettingsCheckBox_Format"    'Current procedure name

    Dim CheckBoxFont           As Object                                  'Fresh CheckBox font object

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
            Err.Raise vbObjectError + 513, PROC_NAME, "TargetCheckBox cannot be Nothing."
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
            .Caption = vbNullString
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Raises a descriptive runtime error if the font object cannot be created
'
' DEPENDENCIES
'   StdFont
'   DP_FORM_FONT_NAME
'   DP_SETTINGS_CHECKBOX_FONT_SIZE
'
' NOTES
'   Do not reuse the returned font object across multiple CheckBoxes
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_SettingsCheckBoxFont_Create" 'Current procedure name

    Dim NewFont                As Object                                   'Fresh CheckBox font object

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
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Function

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
'   white opaque background surface, disables it, and sends it behind the page
'   controls
'
' ERROR POLICY
'   Raises a descriptive runtime error if the MultiPage is missing, the page
'   index is invalid, the resolved page is missing, the name is blank, or the
'   background label cannot be created or formatted
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
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_SettingsPage_BackgroundBuild" 'Current procedure name

    Dim ParentPage             As MSForms.Page                             'Resolved target page
    Dim Lbl_Background         As MSForms.Label                            'Page background label

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
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentMultiPage cannot be Nothing."
        End If

    'Reject an invalid page index
        If PageIndex < 0 Or PageIndex > ParentMultiPage.Pages.Count - 1 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "PageIndex must be between 0 and " & CStr(ParentMultiPage.Pages.Count - 1) & "."
        End If

    'Reject an empty background name
        If Len(Trim$(BackgroundName)) = 0 Then
            Err.Raise vbObjectError + 515, PROC_NAME, "BackgroundName cannot be empty."
        End If

'------------------------------------------------------------------------------
' RETRIEVE TARGET PAGE
'------------------------------------------------------------------------------
    'Retrieve the target MultiPage page
        Set ParentPage = ParentMultiPage.Pages(PageIndex)

    'Reject a missing resolved page
        If ParentPage Is Nothing Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Target page could not be resolved from the supplied MultiPage."
        End If

'------------------------------------------------------------------------------
' CREATE / FORMAT BACKGROUND LABEL
'------------------------------------------------------------------------------
    'Create or retrieve the page background label
        Set Lbl_Background = UF_Ensure_PageLabel(ParentPage, BackgroundName)

    'Apply layout and visual properties to the page background label
        With Lbl_Background
            .Caption = vbNullString
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Resolves the picker item index, resets the previous picker hover state, and
'   applies hover formatting to the current picker item
'
' ERROR POLICY
'   Raises a descriptive runtime error if LabelName is blank, if the label name
'   does not map to a picker-panel item, or if controls cannot be found
'
' DEPENDENCIES
'   UF_PickerPanel_ItemIndexFromLabelName
'   UF_PickerPanel_HoverReset
'
' NOTES
'   This routine is called by cDatePickerLabelHook when the hook category is
'   PICKER
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "UF_PickerPanel_HoverApply"     'Current procedure name

    Dim ItemIndex           As Long                 'Resolved picker item index
    Dim IsSelected          As Boolean              'True when hovered item is selected
    
    Dim Lbl_ItemBg          As MSForms.Label        'Picker item background label
    Dim Lbl_Item            As MSForms.Label        'Picker item text label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty label name
        If Len(Trim$(LabelName)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "LabelName cannot be empty."
        End If

'------------------------------------------------------------------------------
' RESOLVE PICKER ITEM
'------------------------------------------------------------------------------
    'Resolve the picker item index from the hovered label name
        ItemIndex = UF_PickerPanel_ItemIndexFromLabelName(LabelName)

    'Reject invalid picker item indexes
        If ItemIndex < 1 Or ItemIndex > DP_PICKER_ITEM_COUNT Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "LabelName does not resolve to a valid picker-panel item."
        End If

'------------------------------------------------------------------------------
' EXIT IF ALREADY HOVERED
'------------------------------------------------------------------------------
    'Exit if the requested item is already highlighted
        If mHoveredPickerItemIndex = ItemIndex Then
            Exit Sub
        End If
    'Remove footer hover when entering the picker panel
        UF_Footer_HoverReset
        
'------------------------------------------------------------------------------
' RESET CURRENT HOVER STATE
'------------------------------------------------------------------------------
    'Remove hover formatting from the previously highlighted picker item
        UF_PickerPanel_HoverReset

'------------------------------------------------------------------------------
' RETRIEVE PICKER ITEM CONTROLS
'------------------------------------------------------------------------------
    'Ensure cached picker-panel references are available
        UF_PickerPanel_EnsureCache
    'Retrieve the cached picker item background label
        Set Lbl_ItemBg = mPickerBackLabels(ItemIndex)
    'Retrieve the cached picker item text label
        Set Lbl_Item = mPickerTextLabels(ItemIndex)

'------------------------------------------------------------------------------
' RESOLVE SELECTED STATE
'------------------------------------------------------------------------------
    'Resolve whether the hovered month item is currently selected
        If mPickerPanelMode = 1 Then
            IsSelected = (ItemIndex = mDisplayMonth)
        End If
    'Resolve whether the hovered year item is currently selected
        If mPickerPanelMode = 2 Then
            If Len(Trim$(CStr(Lbl_Item.Tag))) > 0 Then
                If IsNumeric(CStr(Lbl_Item.Tag)) Then
                    IsSelected = (CLng(Lbl_Item.Tag) = mDisplayYear)
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
                    .SpecialEffect = fmSpecialEffectBump
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Restores the currently highlighted picker item to normal or selected state
'
' ERROR POLICY
'   Best-effort cleanup
'
'   Silently exits when the picker panel or the cached item labels are not
'   available
'
' DEPENDENCIES
'   UF_PickerPanel_EnsureCache
'   mPickerTextLabels
'   mPickerBackLabels
'
' NOTES
'   This routine resets only the currently hovered picker item, not all 12 items
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IsSelected             As Boolean              'True when hovered item is selected
    Dim HoveredIndex           As Long                 'Cached hovered picker item index
    Dim Lbl_ItemBg             As MSForms.Label        'Picker item background label
    Dim Lbl_Item               As MSForms.Label        'Picker item text label

'------------------------------------------------------------------------------
' EXIT IF NOTHING IS HOVERED
'------------------------------------------------------------------------------
    'Exit if no picker item is currently highlighted
        If mHoveredPickerItemIndex = 0 Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' CAPTURE HOVER STATE
'------------------------------------------------------------------------------
    'Capture the hovered index before any best-effort operation
        HoveredIndex = mHoveredPickerItemIndex

'------------------------------------------------------------------------------
' ENSURE CACHE
'------------------------------------------------------------------------------
    'Attempt to ensure cached picker-panel references are available
        On Error Resume Next

    'Rebuild missing picker-panel cache references when possible
        UF_PickerPanel_EnsureCache

    'Retrieve the cached picker item background label
        Set Lbl_ItemBg = mPickerBackLabels(HoveredIndex)

    'Retrieve the cached picker item text label
        Set Lbl_Item = mPickerTextLabels(HoveredIndex)

    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' EXIT IF LABELS ARE UNAVAILABLE
'------------------------------------------------------------------------------
    'Clear hover state and exit if the cached labels are unavailable
        If Lbl_ItemBg Is Nothing Or Lbl_Item Is Nothing Then
            mHoveredPickerItemIndex = 0
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RESOLVE SELECTED STATE
'------------------------------------------------------------------------------
    'Resolve whether the hovered month item is currently selected
        If mPickerPanelMode = 1 Then
            IsSelected = (HoveredIndex = mDisplayMonth)
        End If

    'Resolve whether the hovered year item is currently selected
        If mPickerPanelMode = 2 Then
            If Len(Trim$(CStr(Lbl_Item.Tag))) > 0 Then
                If IsNumeric(CStr(Lbl_Item.Tag)) Then
                    IsSelected = (CLng(Lbl_Item.Tag) = mDisplayYear)
                End If
            End If
        End If

'------------------------------------------------------------------------------
' RESTORE NORMAL STATE
'------------------------------------------------------------------------------
    'Restore the normal picker-item background state
        With Lbl_ItemBg
            .BackColor = DP_PICKER_ITEM_BACK_COLOR
            .BorderColor = DP_PICKER_ITEM_BORDER_COLOR
            .SpecialEffect = fmSpecialEffectFlat
        End With

    'Restore the normal picker-item text color
        Lbl_Item.ForeColor = DP_PICKER_ITEM_FORE_COLOR

'------------------------------------------------------------------------------
' RESTORE SELECTED STATE
'------------------------------------------------------------------------------
    'Restore selected-item formatting when applicable
        If IsSelected Then

            'Apply selected picker-item background formatting
                With Lbl_ItemBg
                    .BackColor = DP_DAY_SELECTED_BACK_COLOR
                    .BorderColor = DP_DAY_SELECTED_BACK_COLOR
                End With

            'Apply selected picker-item text formatting
                Lbl_Item.ForeColor = DP_DAY_SELECTED_FORE_COLOR

        End If

'------------------------------------------------------------------------------
' CLEAR HOVER STATE
'------------------------------------------------------------------------------
    'Clear the current hovered picker item index
        mHoveredPickerItemIndex = 0

End Sub
Private Function UF_PickerPanel_ItemIndexFromLabelName(ByVal LabelName As String) As Long

'
'------------------------------------------------------------------------------
'                           GET PICKER ITEM INDEX FROM LABEL NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the picker-panel item index from a picker-panel label name
'
' WHY THIS EXISTS
'   Each picker-panel item is composed of two labels. Hovering either label must
'   resolve to the same item index
'
' INPUTS
'   LabelName
'     Picker-panel label name
'
' RETURNS
'   Picker-panel item index from 1 to 12, or zero when unsupported
'
' ERROR POLICY
'   Returns zero when the label name is blank, unsupported, or does not contain
'   a numeric suffix
'
' DEPENDENCIES
'   None
'
' NOTES
'   The background-label prefix must be tested before the text-label prefix
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const BG_PREFIX            As String = "Lbl_MonthYearBg"             'Picker background label prefix
    Const TEXT_PREFIX          As String = "Lbl_MonthYear"               'Picker text label prefix

    Dim RawIndex               As String                                 'Raw numeric suffix

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Return zero for empty label names
        If Len(Trim$(LabelName)) = 0 Then
            UF_PickerPanel_ItemIndexFromLabelName = 0
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE INDEX FROM BACKGROUND LABEL
'------------------------------------------------------------------------------
    'Resolve indexes from picker background labels first
        If Left$(LabelName, Len(BG_PREFIX)) = BG_PREFIX Then

            'Extract the raw numeric suffix
                RawIndex = Mid$(LabelName, Len(BG_PREFIX) + 1)

            'Return zero when the suffix is not numeric
                If Not IsNumeric(RawIndex) Then
                    UF_PickerPanel_ItemIndexFromLabelName = 0
                    Exit Function
                End If

            'Return the parsed picker item index
                UF_PickerPanel_ItemIndexFromLabelName = CLng(RawIndex)
                Exit Function

        End If

'------------------------------------------------------------------------------
' RESOLVE INDEX FROM TEXT LABEL
'------------------------------------------------------------------------------
    'Resolve indexes from picker text labels
        If Left$(LabelName, Len(TEXT_PREFIX)) = TEXT_PREFIX Then

            'Extract the raw numeric suffix
                RawIndex = Mid$(LabelName, Len(TEXT_PREFIX) + 1)

            'Return zero when the suffix is not numeric
                If Not IsNumeric(RawIndex) Then
                    UF_PickerPanel_ItemIndexFromLabelName = 0
                    Exit Function
                End If

            'Return the parsed picker item index
                UF_PickerPanel_ItemIndexFromLabelName = CLng(RawIndex)
                Exit Function

        End If

'------------------------------------------------------------------------------
' RETURN FALLBACK
'------------------------------------------------------------------------------
    'Return zero for unsupported picker label names
        UF_PickerPanel_ItemIndexFromLabelName = 0

End Function

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
'   Resolves the new display month, validates the rendered calendar boundary,
'   hides overlay panels, initializes the keyboard date to the first day of the
'   new display month, and repopulates the day grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if MonthOffset is unsupported, if the
'   current display period is invalid, if the resulting display period cannot be
'   rendered, or if the grid cannot be refreshed
'
' DEPENDENCIES
'   UF_DayGrid_Populate
'   UF_SettingsPanel_Hide
'
' NOTES
'   DateAdd is used so month transitions across year boundaries are handled by
'   VBA date arithmetic
'
'   January 100 and December 9999 are rejected because the fixed 6 x 7 calendar
'   grid cannot render the required adjacent-month dates outside the VBA Date
'   supported range
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Header_MoveMonth"          'Current procedure name

    Const VBA_DATE_MIN_YEAR    As Long = 100                              'Minimum year supported by VBA Date
    Const VBA_DATE_MAX_YEAR    As Long = 9999                             'Maximum year supported by VBA Date

    Dim NewDisplayDate         As Date                                    'New first day of displayed month
    Dim NewDisplayYear         As Long                                    'New displayed year
    Dim NewDisplayMonth        As Long                                    'New displayed month
    Dim Fra_PickerPanel        As MSForms.Frame                           'Reusable picker panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject unsupported month movement
        If MonthOffset <> -1 And MonthOffset <> 1 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "MonthOffset must be -1 or 1."
        End If

    'Reject invalid current display year
        If mDisplayYear < DP_MIN_YEAR Or mDisplayYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "mDisplayYear must be between " & CStr(DP_MIN_YEAR) & _
                " and " & CStr(DP_MAX_YEAR) & "."
        End If

    'Reject invalid current display month
        If mDisplayMonth < 1 Or mDisplayMonth > 12 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "mDisplayMonth must be between 1 and 12."
        End If

    'Reject movement before the lower supported date boundary
        If mDisplayYear = DP_MIN_YEAR And mDisplayMonth = 1 And MonthOffset < 0 Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Displayed month cannot move before January " & CStr(DP_MIN_YEAR) & "."
        End If

    'Reject movement after the upper supported date boundary
        If mDisplayYear = DP_MAX_YEAR And mDisplayMonth = 12 And MonthOffset > 0 Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "Displayed month cannot move after December " & CStr(DP_MAX_YEAR) & "."
        End If

'------------------------------------------------------------------------------
' RESOLVE NEW DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Resolve the new display date
        NewDisplayDate = VBA.DateAdd("m", MonthOffset, _
            VBA.DateSerial(mDisplayYear, mDisplayMonth, 1))

    'Resolve the new display year
        NewDisplayYear = VBA.Year(NewDisplayDate)

    'Resolve the new display month
        NewDisplayMonth = VBA.Month(NewDisplayDate)

'------------------------------------------------------------------------------
' VALIDATE RENDERABLE BOUNDARIES
'------------------------------------------------------------------------------
    'Reject the lower VBA Date rendering boundary
        If NewDisplayYear = VBA_DATE_MIN_YEAR And NewDisplayMonth = 1 Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "January " & CStr(VBA_DATE_MIN_YEAR) & _
                " cannot render a complete DatePicker grid."
        End If

    'Reject the upper VBA Date rendering boundary
        If NewDisplayYear = VBA_DATE_MAX_YEAR And NewDisplayMonth = 12 Then
            Err.Raise vbObjectError + 519, PROC_NAME, _
                "December " & CStr(VBA_DATE_MAX_YEAR) & _
                " cannot render a complete DatePicker grid."
        End If

'------------------------------------------------------------------------------
' STORE DISPLAY STATE
'------------------------------------------------------------------------------
    'Store the new display year
        mDisplayYear = NewDisplayYear

    'Store the new display month
        mDisplayMonth = NewDisplayMonth

    'Initialize keyboard date to the first day of the new display month
        mKeyboardDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)

    'Mark keyboard navigation as initialized
        mHasKeyboardDate = True

'------------------------------------------------------------------------------
' HIDE OVERLAY PANELS
'------------------------------------------------------------------------------
    'Temporarily ignore lookup errors if the picker panel has not been created
        On Error Resume Next

    'Retrieve the picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)

    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Clear active picker hover before hiding the picker panel
        If mHoveredPickerItemIndex <> 0 Then
            UF_PickerPanel_HoverReset
        End If

    'Hide the picker panel if it exists
        If Not Fra_PickerPanel Is Nothing Then
            Fra_PickerPanel.Visible = False
            mPickerPanelMode = 0
        End If

    'Hide the settings panel after month navigation
        UF_SettingsPanel_Hide

'------------------------------------------------------------------------------
' REFRESH CALENDAR
'------------------------------------------------------------------------------
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Scrolls the year panel when visible, otherwise updates the displayed year,
'   initializes keyboard date to the first day of the displayed month, hides
'   overlay panels, and repopulates the day grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if YearOffset is unsupported or the
'   resulting display period is invalid
'
' DEPENDENCIES
'   UF_PickerPanel_PopulateYears
'   UF_DayGrid_Populate
'   UF_SettingsPanel_Hide
'
' NOTES
'   The year panel start is clamped to the valid 12-year display range
'
'   January 100 and December 9999 are rejected because the fixed 6 x 7 calendar
'   grid cannot render the required adjacent-month dates outside the VBA Date
'   supported range
'
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Header_MoveYear"          'Current procedure name

    Const VBA_DATE_MIN_YEAR    As Long = 100                             'Minimum year supported by VBA Date
    Const VBA_DATE_MAX_YEAR    As Long = 9999                            'Maximum year supported by VBA Date

    Dim NewDisplayYear         As Long                                   'New displayed year
    Dim NewYearPanelStart      As Long                                   'New first year shown in panel
    Dim Fra_PickerPanel        As MSForms.Frame                          'Reusable picker panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject unsupported year movement
        If YearOffset <> -1 And YearOffset <> 1 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "YearOffset must be -1 or 1."
        End If

'------------------------------------------------------------------------------
' RETRIEVE PICKER PANEL
'------------------------------------------------------------------------------
    'Temporarily ignore lookup errors if the picker panel has not been created
        On Error Resume Next

    'Retrieve the picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)

    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SCROLL YEAR PANEL IF VISIBLE
'------------------------------------------------------------------------------
    'Scroll the 12-year panel when it is visible in year-selection mode
        If Not Fra_PickerPanel Is Nothing Then
            If Fra_PickerPanel.Visible And mPickerPanelMode = 2 Then

                'Clear active picker hover before repopulating the year panel
                    If mHoveredPickerItemIndex <> 0 Then
                        UF_PickerPanel_HoverReset
                    End If

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
                    If NewYearPanelStart = mYearPanelStart Then
                        Exit Sub
                    End If

                'Store the new first visible year
                    mYearPanelStart = NewYearPanelStart

                'Refresh the year picker panel
                    UF_PickerPanel_PopulateYears

                'Exit after scrolling the panel
                    Exit Sub

            End If
        End If

'------------------------------------------------------------------------------
' VALIDATE DISPLAY STATE
'------------------------------------------------------------------------------
    'Reject invalid current displayed year
        If mDisplayYear < DP_MIN_YEAR Or mDisplayYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "mDisplayYear must be between " & CStr(DP_MIN_YEAR) & _
                " and " & CStr(DP_MAX_YEAR) & "."
        End If

    'Reject invalid current displayed month
        If mDisplayMonth < 1 Or mDisplayMonth > 12 Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "mDisplayMonth must be between 1 and 12."
        End If

'------------------------------------------------------------------------------
' RESOLVE NEW DISPLAYED YEAR
'------------------------------------------------------------------------------
    'Resolve the new displayed year
        NewDisplayYear = mDisplayYear + YearOffset

    'Reject years outside the supported display range
        If NewDisplayYear < DP_MIN_YEAR Or NewDisplayYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "Displayed year must remain between " & CStr(DP_MIN_YEAR) & _
                " and " & CStr(DP_MAX_YEAR) & "."
        End If

    'Reject the lower VBA Date rendering boundary
        If NewDisplayYear = VBA_DATE_MIN_YEAR And mDisplayMonth = 1 Then
            Err.Raise vbObjectError + 517, PROC_NAME, _
                "January " & CStr(VBA_DATE_MIN_YEAR) & _
                " cannot render a complete DatePicker grid."
        End If

    'Reject the upper VBA Date rendering boundary
        If NewDisplayYear = VBA_DATE_MAX_YEAR And mDisplayMonth = 12 Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "December " & CStr(VBA_DATE_MAX_YEAR) & _
                " cannot render a complete DatePicker grid."
        End If

'------------------------------------------------------------------------------
' STORE DISPLAY STATE
'------------------------------------------------------------------------------
    'Store the new displayed year
        mDisplayYear = NewDisplayYear

    'Initialize keyboard date to the first day of the current display month in the new year
        mKeyboardDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)

    'Mark keyboard navigation as initialized
        mHasKeyboardDate = True

'------------------------------------------------------------------------------
' HIDE OVERLAY PANELS
'------------------------------------------------------------------------------
    'Clear active picker hover before hiding the picker panel
        If mHoveredPickerItemIndex <> 0 Then
            UF_PickerPanel_HoverReset
        End If

    'Hide the reusable picker panel if it exists
        If Not Fra_PickerPanel Is Nothing Then
            Fra_PickerPanel.Visible = False
            mPickerPanelMode = 0
        End If

    'Hide the settings panel after year navigation
        UF_SettingsPanel_Hide

'------------------------------------------------------------------------------
' REFRESH CALENDAR
'------------------------------------------------------------------------------
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub
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
'   Header labels and the day grid depend on mDisplayMonth and mDisplayYear
'
' INPUTS
'   InitialDate
'     Initial date used to set the displayed month and year
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Sets mDisplayYear and mDisplayMonth
'
' ERROR POLICY
'   Raises a descriptive runtime error if the resolved display year is outside
'   the supported DatePicker range
'
' DEPENDENCIES
'   VBA.Date
'
' NOTES
'   A zero date is treated as no explicit initial date
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DisplayPeriod_Initialize"    'Current procedure name

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
            InitialDate = VBA.Date
        End If

'------------------------------------------------------------------------------
' VALIDATE INITIAL DATE
'------------------------------------------------------------------------------
    'Reject years outside the supported display range
        If VBA.Year(InitialDate) < DP_MIN_YEAR Or VBA.Year(InitialDate) > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "InitialDate year must be between " & CStr(DP_MIN_YEAR) & " and " & CStr(DP_MAX_YEAR) & "."
        End If

'------------------------------------------------------------------------------
' STORE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Store the displayed year
        mDisplayYear = VBA.Year(InitialDate)

    'Store the displayed month
        mDisplayMonth = VBA.Month(InitialDate)

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
'   Uses the current selected date when available; otherwise uses InitialDate
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
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_KeyboardDate_Initialize"     'Current procedure name

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
            mKeyboardDate = VBA.DateValue(gDP_SelectedDate)
        Else
            mKeyboardDate = VBA.DateValue(InitialDate)
        End If

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Initializes the keyboard date if needed, applies the requested day offset,
'   and refreshes the visible calendar grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if the movement cannot be applied
'
' DEPENDENCIES
'   UF_KeyboardDate_Set
'
' NOTES
'   This routine does not write to Excel
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_KeyboardDate_MoveDays"       'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' ENSURE KEYBOARD DATE
'------------------------------------------------------------------------------
    'Initialize keyboard date to the first displayed day when missing
        If Not mHasKeyboardDate Then
            mKeyboardDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
            mHasKeyboardDate = True
        End If

'------------------------------------------------------------------------------
' MOVE DATE
'------------------------------------------------------------------------------
    'Move the keyboard date by the requested day offset
        UF_KeyboardDate_Set VBA.DateAdd("d", DayOffset, mKeyboardDate)

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
'   Initializes the keyboard date if needed, applies the requested month offset,
'   and refreshes the visible calendar grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if the movement cannot be applied
'
' DEPENDENCIES
'   UF_KeyboardDate_Set
'
' NOTES
'   VBA DateAdd handles month-end adjustment automatically
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_KeyboardDate_MoveMonths"     'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' ENSURE KEYBOARD DATE
'------------------------------------------------------------------------------
    'Initialize keyboard date to the first displayed day when missing
        If Not mHasKeyboardDate Then
            mKeyboardDate = VBA.DateSerial(mDisplayYear, mDisplayMonth, 1)
            mHasKeyboardDate = True
        End If

'------------------------------------------------------------------------------
' MOVE DATE
'------------------------------------------------------------------------------
    'Move the keyboard date by the requested month offset
        UF_KeyboardDate_Set VBA.DateAdd("m", MonthOffset, mKeyboardDate)

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
'   Keyboard navigation can fire repeatedly. Rebuilding all 42 day cells on every
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
'   - Stores the previous keyboard date
'   - Stores the new keyboard date
'   - If the new date is still in the displayed month, refreshes only the old and
'     new affected day cells
'   - If the new date moves to another displayed month, repopulates the full grid
'   - Hides the picker panel when present
'
' ERROR POLICY
'   Raises a descriptive runtime error if NewKeyboardDate cannot be displayed
'
' DEPENDENCIES
'   UF_DayCell_RefreshVisibleDate
'   UF_DayGrid_Populate
'
' NOTES
'   This routine preserves existing behavior when navigation crosses a month
'   boundary
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_KeyboardDate_Set"        'Current procedure name

    Dim PreviousDate           As Date                                  'Previous keyboard-selected date
    Dim HasPreviousDate        As Boolean                               'True when previous keyboard date exists
    Dim NewDateOnly            As Date                                  'New keyboard date without time
    Dim NeedsFullRefresh       As Boolean                               'True when the displayed month must change
    Dim Fra_PickerPanel        As MSForms.Frame                         'Reusable picker panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CAPTURE PREVIOUS STATE
'------------------------------------------------------------------------------
    'Capture the previous keyboard date when available
        If mHasKeyboardDate Then
            PreviousDate = VBA.DateValue(mKeyboardDate)
            HasPreviousDate = True
        End If

    'Normalize the new keyboard date
        NewDateOnly = VBA.DateValue(NewKeyboardDate)

'------------------------------------------------------------------------------
' RESOLVE REFRESH STRATEGY
'------------------------------------------------------------------------------
    'Resolve whether the keyboard date moved outside the displayed month
        NeedsFullRefresh = _
            (VBA.Year(NewDateOnly) <> mDisplayYear Or _
             VBA.Month(NewDateOnly) <> mDisplayMonth)

'------------------------------------------------------------------------------
' STORE NEW KEYBOARD STATE
'------------------------------------------------------------------------------
    'Store the new keyboard date
        mKeyboardDate = NewDateOnly
    'Mark keyboard navigation as initialized
        mHasKeyboardDate = True

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Temporarily ignore lookup errors if the picker panel has not been created
        On Error Resume Next
    'Retrieve the picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)
    'Restore controlled error handling
        On Error GoTo ErrorHandler
    'Hide the picker panel if available
        If Not Fra_PickerPanel Is Nothing Then
            Fra_PickerPanel.Visible = False
        End If
    'Clear picker-panel mode after hiding the panel
        mPickerPanelMode = 0

'------------------------------------------------------------------------------
' APPLY FULL REFRESH WHEN MONTH CHANGES
'------------------------------------------------------------------------------
    'Repopulate the full grid when the new keyboard date changes displayed month
        If NeedsFullRefresh Then
            'Store the displayed year
                mDisplayYear = VBA.Year(NewDateOnly)
            'Store the displayed month
                mDisplayMonth = VBA.Month(NewDateOnly)
            'Refresh the full calendar grid
                UF_DayGrid_Populate mDisplayYear, mDisplayMonth
            'Exit after full refresh
                Exit Sub
        End If

'------------------------------------------------------------------------------
' APPLY MINIMAL REFRESH WHEN MONTH IS UNCHANGED
'------------------------------------------------------------------------------
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Ignores non-header labels, resets the previous header hover state, and
'   applies a subtle hover state to the requested clickable header label
'
' ERROR POLICY
'   Raises a descriptive runtime error if LabelName is empty or if a valid header
'   label cannot be found or formatted
'
' DEPENDENCIES
'   UF_Header_HoverReset
'
' NOTES
'   This routine avoids changing Font properties because header labels have
'   different normal bold states
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Header_HoverApply"  'Current procedure name

    Dim Lbl_Header             As MSForms.Label                          'Hovered header label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty label name
        If Len(Trim$(LabelName)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "LabelName cannot be empty."
        End If

'------------------------------------------------------------------------------
' EXIT IF NOT HEADER LABEL
'------------------------------------------------------------------------------
    'Exit if the routed label is not a clickable header label
        Select Case LabelName

            Case "Lbl_HeaderMonth", _
                 "Lbl_HeaderYear", _
                 "Lbl_PrevMonth", _
                 "Lbl_NextMonth", _
                 "Lbl_PrevYear", _
                 "Lbl_NextYear"

            Case Else
                Exit Sub

        End Select
    
    'Remove footer hover when entering the header
        UF_Footer_HoverReset
        
'------------------------------------------------------------------------------
' EXIT IF ALREADY HOVERED
'------------------------------------------------------------------------------
    'Exit if the requested label is already highlighted
        If StrComp(mHoveredHeaderLabelName, LabelName, vbBinaryCompare) = 0 Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RESET CURRENT HOVER STATE
'------------------------------------------------------------------------------
    'Remove hover formatting from the previously highlighted header label
        UF_Header_HoverReset

'------------------------------------------------------------------------------
' APPLY NEW HOVER STATE
'------------------------------------------------------------------------------
    'Retrieve the requested header label
        Set Lbl_Header = Me.Controls(LabelName)

    'Apply clickable hover visual formatting
        With Lbl_Header
            .BackStyle = fmBackStyleOpaque
            .BackColor = DP_HEADER_HOVER_BACK_COLOR
            .ForeColor = vbWhite
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectBump
        End With

'------------------------------------------------------------------------------
' STORE HOVER STATE
'------------------------------------------------------------------------------
    'Store the current hovered header label name
        mHoveredHeaderLabelName = LabelName

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
'   header labels and moves back over the UserForm background
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Restores the currently highlighted header label to its normal visual state
'
' ERROR POLICY
'   Silently ignores a missing previously-hovered label because runtime controls
'   may have been rebuilt
'
' DEPENDENCIES
'   UserForm.Controls collection
'
' NOTES
'   This routine restores only the visual properties changed by
'   UF_Header_HoverApply
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Header_HoverReset"  'Current procedure name

    Dim Lbl_Header             As MSForms.Label                          'Previously highlighted label

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' EXIT IF NOTHING IS HOVERED
'------------------------------------------------------------------------------
    'Exit if no header label is currently highlighted
        If Len(mHoveredHeaderLabelName) = 0 Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RETRIEVE CURRENTLY HOVERED LABEL
'------------------------------------------------------------------------------
    'Temporarily ignore lookup errors if the previous label no longer exists
        On Error Resume Next

    'Retrieve the previously highlighted header label
        Set Lbl_Header = Me.Controls(mHoveredHeaderLabelName)

    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESET VISUAL STATE
'------------------------------------------------------------------------------
    'Reset the label if it was found
        If Not Lbl_Header Is Nothing Then

            'Restore normal clickable-header formatting
                With Lbl_Header
                    .BackStyle = fmBackStyleTransparent
                    .ForeColor = DP_HEADER_FORE_COLOR
                    .BorderStyle = fmBorderStyleNone
                    .SpecialEffect = fmSpecialEffectFlat
                End With

        End If

'------------------------------------------------------------------------------
' CLEAR HOVER STATE
'------------------------------------------------------------------------------
    'Clear the current hovered header label name
        mHoveredHeaderLabelName = vbNullString

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

'------------------------------------------------------------------------------
' PUBLIC REFRESH AND CLOCK ENTRY POINTS
'------------------------------------------------------------------------------

Public Sub UF_DP_AfterSuccessfulSelection(ByVal SelectedDate As Date)

'
'==============================================================================
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
'   - captures the previous keyboard-selected date
'   - stores the new selected date as the active keyboard date
'   - hides the picker panel if present
'   - clears transient hover state
'   - restores only the previous selected day, when visible
'   - refreshes only the newly selected day, when visible
'
' ERROR POLICY
'   Raises a descriptive runtime error if the form selection state cannot be
'   synchronized
'
' DEPENDENCIES
'   UF_DayCell_RefreshVisibleDate
'
' NOTES
'   This routine intentionally does not call UF_DayGrid_Populate
'
'   Rebuilding the full grid from inside a label click event can raise MSForms
'   run-time error 5 in some environments
'
' UPDATED
'   2026-04-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DP_AfterSuccessfulSelection" 'Current procedure name

    Dim PreviousDate           As Date                                   'Previously selected keyboard date
    Dim HasPreviousDate        As Boolean                                'True when previous keyboard date exists
    Dim SelectedDateOnly       As Date                                   'Selected date without time
    Dim Fra_PickerPanel        As MSForms.Frame                          'Reusable picker panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

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
' STORE NEW STATE
'------------------------------------------------------------------------------
    'Store the selected date as the active keyboard/navigation date
        mKeyboardDate = SelectedDateOnly
    'Mark keyboard navigation date as available
        mHasKeyboardDate = True

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Temporarily ignore missing picker-panel lookup errors
        On Error Resume Next
    'Retrieve the picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)
    'Restore controlled error handling
        On Error GoTo ErrorHandler
    'Hide the picker panel when available
        If Not Fra_PickerPanel Is Nothing Then
            Fra_PickerPanel.Visible = False
        End If
    'Clear picker-panel mode
        mPickerPanelMode = 0

'------------------------------------------------------------------------------
' CLEAR TRANSIENT HOVER STATE
'------------------------------------------------------------------------------
    'Clear the current day hover state
        mHoveredDayLabelName = vbNullString
    'Clear the current day-cell hover index
        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
    'Clear the current header hover state
        If Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset
    'Clear the current picker hover state
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Clear the current footer hover state
        If Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset

'------------------------------------------------------------------------------
' REFRESH ONLY AFFECTED DAY CELLS
'------------------------------------------------------------------------------
    'Restore the previous selected day when it is visible and changed
        If HasPreviousDate Then
            If VBA.DateValue(PreviousDate) <> SelectedDateOnly Then
                UF_DayCell_RefreshVisibleDate PreviousDate
            End If
        End If

    'Apply the selected visual state to the newly selected day when visible
        UF_DayCell_RefreshVisibleDate SelectedDateOnly

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


Public Sub UF_DP_RefreshFromExternalSelection( _
    ByVal SelectedDate As Date, _
    ByVal HasSelectedDate As Boolean)

'
'==============================================================================
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
'     Date resolved from the new ActiveCell, or current system date when no date
'     was available
'
'   HasSelectedDate
'     True when the new ActiveCell contains a valid date
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - if ActiveCell contains a date, displays that date's month and highlights it
'   - if ActiveCell does not contain a date, clears keyboard-selected state and
'     displays the current system month
'   - clears stale hover state
'   - hides the picker panel
'   - repopulates the day grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if the form cannot be refreshed
'
' DEPENDENCIES
'   UF_DayCell_HoverReset
'   UF_Header_HoverReset
'   UF_PickerPanel_HoverReset
'   UF_Footer_HoverReset
'   UF_DayGrid_Populate
'
' NOTES
'   This routine does not write to Excel
'
' UPDATED
'   2026-04-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DP_RefreshFromExternalSelection"    'Current procedure name

    Dim DisplayDate            As Date                                         'Date used to resolve displayed month
    Dim Fra_PickerPanel        As MSForms.Frame                                'Reusable picker panel

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESET TRANSIENT HOVER STATE
'------------------------------------------------------------------------------
    'Remove current day-label hover formatting
        UF_DayCell_HoverReset
    'Remove current header hover formatting
        UF_Header_HoverReset
    'Remove current picker-panel hover formatting
        UF_PickerPanel_HoverReset
    'Remove current footer hover formatting
        UF_Footer_HoverReset

'------------------------------------------------------------------------------
' RESOLVE DISPLAY DATE
'------------------------------------------------------------------------------
    'Use the selected date when ActiveCell contains a date
        If HasSelectedDate Then
            DisplayDate = VBA.DateValue(SelectedDate)
            mKeyboardDate = DisplayDate
            mHasKeyboardDate = True
        Else
            DisplayDate = VBA.Date
            mKeyboardDate = 0
            mHasKeyboardDate = False
        End If

    'Store the displayed year
        mDisplayYear = VBA.Year(DisplayDate)

    'Store the displayed month
        mDisplayMonth = VBA.Month(DisplayDate)

'------------------------------------------------------------------------------
' HIDE PICKER PANEL
'------------------------------------------------------------------------------
    'Temporarily ignore missing picker-panel lookup errors
        On Error Resume Next
    'Retrieve the picker panel
        Set Fra_PickerPanel = Me.Controls(DP_PICKER_PANEL_NAME)
    'Restore controlled error handling
        On Error GoTo ErrorHandler
    'Hide the picker panel when available
        If Not Fra_PickerPanel Is Nothing Then
            Fra_PickerPanel.Visible = False
        End If
    'Clear picker-panel mode
        mPickerPanelMode = 0

'------------------------------------------------------------------------------
' CLEAR STORED HOVER STATE
'------------------------------------------------------------------------------
    'Clear the current day hover state
        mHoveredDayLabelName = vbNullString
    'Clear the current header hover state
        mHoveredHeaderLabelName = vbNullString
    'Clear the current picker hover state
        If mHoveredPickerItemIndex <> 0 Then UF_PickerPanel_HoverReset
    'Clear the current footer hover state
        If Len(mHoveredFooterActionName) <> 0 Then UF_Footer_HoverReset

        If mHoveredDayCellIndex <> 0 Then UF_DayCell_HoverReset
        If Len(mHoveredHeaderLabelName) <> 0 Then UF_Header_HoverReset


'------------------------------------------------------------------------------
' REFRESH CALENDAR GRID
'------------------------------------------------------------------------------
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub


Public Sub UF_DP_RefreshSettings()

'
'------------------------------------------------------------------------------
'                           REFRESH SETTINGS-DEPENDENT CAPTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes DatePicker captions affected by saved display settings
'
' WHY THIS EXISTS
'   Changing UseLocalDayNames or FirstDayOfWeek while the form is already open
'   must refresh the visible captions and calendar-grid positioning
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Refreshes header captions, day-of-week labels, today footer value, and the
'   day-label grid
'
' ERROR POLICY
'   Raises a descriptive runtime error if any caption refresh fails
'
' DEPENDENCIES
'   UF_Header_BuildLabels
'   UF_WeekdayRow_Build
'   M_Caption_GetDate
'   UF_DayGrid_Populate
'
' NOTES
'   This routine deliberately avoids recreating the full footer section
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DP_RefreshSettings"    'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REFRESH HEADER CAPTIONS
'------------------------------------------------------------------------------
    'Refresh the month and year header labels
        UF_Header_BuildLabels

'------------------------------------------------------------------------------
' REFRESH DAY-OF-WEEK CAPTIONS
'------------------------------------------------------------------------------
    'Refresh the weekday header row
        UF_WeekdayRow_Build

'------------------------------------------------------------------------------
' REFRESH TODAY CAPTION
'------------------------------------------------------------------------------
    'Refresh the today value label
        Me.Controls("Lbl_Today").Caption = M_Caption_GetDate(VBA.Date, gDP_UseLocalNames)

'------------------------------------------------------------------------------
' REFRESH DAY GRID
'------------------------------------------------------------------------------
    'Refresh day labels
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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   The live clock timer should update only the time text without rebuilding or
'   repainting the UserForm
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Updates Lbl_Time only when the displayed value has changed
'
' ERROR POLICY
'   Raises a descriptive runtime error if the time label cannot be updated
'
' DEPENDENCIES
'   UserForm.Controls collection
'
' NOTES
'   This routine deliberately does not call Repaint
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_DP_UpdateLiveClock"          'Current procedure name

    Dim NewTimeCaption         As String                                 'Current formatted time caption

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' BUILD TIME CAPTION
'------------------------------------------------------------------------------
    'Build the new time caption
        NewTimeCaption = Format$(VBA.Time, "hh:nn:ss")

'------------------------------------------------------------------------------
' UPDATE TIME LABEL
'------------------------------------------------------------------------------
    'Update the label only if the caption changed
        If Me.Controls("Lbl_Time").Caption <> NewTimeCaption Then
            Me.Controls("Lbl_Time").Caption = NewTimeCaption
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

'------------------------------------------------------------------------------
' SHARED HELPERS
'------------------------------------------------------------------------------

Private Function UF_DayCell_GetIndexFromLabelName(ByVal LabelName As String) As Long

'
'==============================================================================
'                   GET DAY CELL INDEX FROM LABEL NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the day-cell index from a DatePicker day label name
'
' WHY THIS EXISTS
'   Each calendar day cell is composed of two runtime labels:
'     - Lbl_DayBg1 to Lbl_DayBg42
'     - Lbl_Day1 to Lbl_Day42
'
'   Hovering or clicking either label must resolve to the same day-cell index
'
' INPUTS
'   LabelName
'     Runtime day label name
'
' RETURNS
'   Day-cell index from 1 to 42, or zero when the label name is unsupported
'
' BEHAVIOR
'   Parses supported day-label prefixes and returns the numeric suffix
'
' ERROR POLICY
'   Returns zero for blank, unsupported, or malformed label names
'
' DEPENDENCIES
'   None
'
' NOTES
'   The background-label prefix must be tested before the text-label prefix
'   because Lbl_DayBg also starts with Lbl_Day
'
' UPDATED
'   2026-04-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const BG_PREFIX            As String = "Lbl_DayBg"                  'Day background label prefix
    Const TEXT_PREFIX          As String = "Lbl_Day"                    'Day text label prefix

    Dim RawIndex               As String                                'Raw numeric suffix

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Return zero for empty label names
        If Len(Trim$(LabelName)) = 0 Then
            UF_DayCell_GetIndexFromLabelName = 0
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE INDEX FROM BACKGROUND LABEL
'------------------------------------------------------------------------------
    'Resolve indexes from day background labels first
        If Left$(LabelName, Len(BG_PREFIX)) = BG_PREFIX Then

            'Extract the raw numeric suffix
                RawIndex = Mid$(LabelName, Len(BG_PREFIX) + 1)

            'Return zero when the suffix is not numeric
                If Not IsNumeric(RawIndex) Then
                    UF_DayCell_GetIndexFromLabelName = 0
                    Exit Function
                End If

            'Return the parsed day-cell index
                UF_DayCell_GetIndexFromLabelName = CLng(RawIndex)
                Exit Function

        End If

'------------------------------------------------------------------------------
' RESOLVE INDEX FROM TEXT LABEL
'------------------------------------------------------------------------------
    'Resolve indexes from day text labels
        If Left$(LabelName, Len(TEXT_PREFIX)) = TEXT_PREFIX Then

            'Extract the raw numeric suffix
                RawIndex = Mid$(LabelName, Len(TEXT_PREFIX) + 1)

            'Return zero when the suffix is not numeric
                If Not IsNumeric(RawIndex) Then
                    UF_DayCell_GetIndexFromLabelName = 0
                    Exit Function
                End If

            'Return the parsed day-cell index
                UF_DayCell_GetIndexFromLabelName = CLng(RawIndex)
                Exit Function

        End If

'------------------------------------------------------------------------------
' RETURN FALLBACK
'------------------------------------------------------------------------------
    'Return zero for unsupported day label names
        UF_DayCell_GetIndexFromLabelName = 0

End Function

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
'   create or reuse logic avoids repeated boilerplate across routines
'
' INPUTS
'   ControlName
'     Name of the label control to retrieve or create
'
' RETURNS
'   The requested MSForms.Label control
'
' BEHAVIOR
'   Reuses an existing MSForms.Label or creates a new one on the UserForm
'
' ERROR POLICY
'   Raises a descriptive runtime error if ControlName is blank, if an existing
'   control with the same name is not a label, or if the label cannot be created
'
' DEPENDENCIES
'   UserForm.Controls collection
'
' NOTES
'   This routine creates controls on the UserForm itself, not inside a Frame
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Ensure_Label"                'Current procedure name

    Dim ExistingControl        As Object                                 'Existing runtime control
    Dim Lbl                    As MSForms.Label                          'Target label control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty control name
        If Len(Trim$(ControlName)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ControlName cannot be empty."
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Clear the object reference before lookup
        Set ExistingControl = Nothing

    'Temporarily ignore lookup errors when the control does not exist yet
        On Error Resume Next

    'Try to retrieve an existing control with the expected name
        Set ExistingControl = Me.Controls(ControlName)

    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING LABEL
'------------------------------------------------------------------------------
    'Reuse the existing control if it is a label
        If Not ExistingControl Is Nothing Then

            'Reject an existing non-label control with the same name
                If Not TypeOf ExistingControl Is MSForms.Label Then
                    Err.Raise vbObjectError + 514, PROC_NAME, _
                        "Control '" & ControlName & "' already exists but is not an MSForms.Label."
                End If

            'Return the existing label
                Set UF_Ensure_Label = ExistingControl

            'Exit after returning the existing label
                Exit Function

        End If

'------------------------------------------------------------------------------
' CREATE LABEL IF MISSING
'------------------------------------------------------------------------------
    'Create the label if it does not already exist
        Set Lbl = Me.Controls.Add("Forms.Label.1", ControlName, True)

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Ensure_FrameMultiPage"       'Current procedure name

    Dim ExistingControl        As Object                                   'Existing frame control
    Dim Mp                     As MSForms.MultiPage                        'Target MultiPage control

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
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentFrame cannot be Nothing."
        End If
    'Reject an empty control name
        If Len(Trim$(ControlName)) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ControlName cannot be empty."
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Clear the object reference before lookup
        Set ExistingControl = Nothing
    'Temporarily ignore lookup errors when the control does not exist yet
        On Error Resume Next
    'Try to retrieve an existing control with the expected name
        Set ExistingControl = ParentFrame.Controls(ControlName)
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING MULTIPAGE
'------------------------------------------------------------------------------
    'Reuse the existing control if it is a MultiPage
        If Not ExistingControl Is Nothing Then

            'Reject an existing non-MultiPage control with the same name
                If Not TypeOf ExistingControl Is MSForms.MultiPage Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Control '" & ControlName & "' already exists but is not an MSForms.MultiPage."
                End If

            'Return the existing MultiPage
                Set UF_Ensure_FrameMultiPage = ExistingControl

            'Exit after returning the existing MultiPage
                Exit Function

        End If

'------------------------------------------------------------------------------
' CREATE MULTIPAGE IF MISSING
'------------------------------------------------------------------------------
    'Create the MultiPage if it does not already exist
        Set Mp = ParentFrame.Controls.Add("Forms.MultiPage.1", ControlName, True)

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   Runtime settings pages need labels hosted inside MSForms.Page controls
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
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Ensure_PageLabel"           'Current procedure name

    Dim ExistingControl        As Object                                  'Existing page control
    Dim Lbl                    As MSForms.Label                           'Target label control

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
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentPage cannot be Nothing."
        End If

    'Reject an empty control name
        If Len(Trim$(ControlName)) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ControlName cannot be empty."
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Clear the object reference before lookup
        Set ExistingControl = Nothing

    'Temporarily ignore lookup errors when the control does not exist yet
        On Error Resume Next

    'Try to retrieve an existing control with the expected name
        Set ExistingControl = ParentPage.Controls(ControlName)

    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING LABEL
'------------------------------------------------------------------------------
    'Reuse the existing control if it is a label
        If Not ExistingControl Is Nothing Then

            'Reject an existing non-label control with the same name
                If Not TypeOf ExistingControl Is MSForms.Label Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Control '" & ControlName & "' already exists but is not an MSForms.Label."
                End If

            'Return the existing label
                Set UF_Ensure_PageLabel = ExistingControl

            'Exit after returning the existing label
                Exit Function

        End If

'------------------------------------------------------------------------------
' CREATE LABEL IF MISSING
'------------------------------------------------------------------------------
    'Create the label if it does not already exist
        Set Lbl = ParentPage.Controls.Add("Forms.Label.1", ControlName, True)

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   controls
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
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Ensure_PageComboBox"        'Current procedure name

    Dim ExistingControl        As Object                                  'Existing page control
    Dim Cbo                    As MSForms.ComboBox                        'Target ComboBox control

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
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentPage cannot be Nothing."
        End If

    'Reject an empty control name
        If Len(Trim$(ControlName)) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ControlName cannot be empty."
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Clear the object reference before lookup
        Set ExistingControl = Nothing

    'Temporarily ignore lookup errors when the control does not exist yet
        On Error Resume Next

    'Try to retrieve an existing control with the expected name
        Set ExistingControl = ParentPage.Controls(ControlName)

    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING COMBOBOX
'------------------------------------------------------------------------------
    'Reuse the existing control if it is a ComboBox
        If Not ExistingControl Is Nothing Then

            'Reject an existing non-ComboBox control with the same name
                If Not TypeOf ExistingControl Is MSForms.ComboBox Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Control '" & ControlName & "' already exists but is not an MSForms.ComboBox."
                End If

            'Return the existing ComboBox
                Set UF_Ensure_PageComboBox = ExistingControl

            'Exit after returning the existing ComboBox
                Exit Function

        End If

'------------------------------------------------------------------------------
' CREATE COMBOBOX IF MISSING
'------------------------------------------------------------------------------
    'Create the ComboBox if it does not already exist
        Set Cbo = ParentPage.Controls.Add("Forms.ComboBox.1", ControlName, True)

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   controls
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
' UPDATED
'   2026-05-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Ensure_PageCheckBox"        'Current procedure name

    Dim ExistingControl        As Object                                  'Existing page control
    Dim Chk                    As MSForms.CheckBox                        'Target CheckBox control

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
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentPage cannot be Nothing."
        End If

    'Reject an empty control name
        If Len(Trim$(ControlName)) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ControlName cannot be empty."
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Clear the object reference before lookup
        Set ExistingControl = Nothing

    'Temporarily ignore lookup errors when the control does not exist yet
        On Error Resume Next

    'Try to retrieve an existing control with the expected name
        Set ExistingControl = ParentPage.Controls(ControlName)

    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING CHECKBOX
'------------------------------------------------------------------------------
    'Reuse the existing control if it is a CheckBox
        If Not ExistingControl Is Nothing Then

            'Reject an existing non-CheckBox control with the same name
                If Not TypeOf ExistingControl Is MSForms.CheckBox Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Control '" & ControlName & "' already exists but is not an MSForms.CheckBox."
                End If

            'Return the existing CheckBox
                Set UF_Ensure_PageCheckBox = ExistingControl

            'Exit after returning the existing CheckBox
                Exit Function

        End If

'------------------------------------------------------------------------------
' CREATE CHECKBOX IF MISSING
'------------------------------------------------------------------------------
    'Create the CheckBox if it does not already exist
        Set Chk = ParentPage.Controls.Add("Forms.CheckBox.1", ControlName, True)

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

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
'   The month/year picker panel uses labels hosted inside a frame. Those labels
'   must be created in the frame's Controls collection
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
' ERROR POLICY
'   Raises a descriptive runtime error if ParentFrame is missing, ControlName is
'   blank, an existing control with the same name is not a label, or the label
'   cannot be created
'
' DEPENDENCIES
'   MSForms.Frame
'
' NOTES
'   This routine creates controls inside the supplied frame
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_Ensure_FrameLabel"           'Current procedure name

    Dim ExistingControl        As Object                                 'Existing frame control
    Dim Lbl                    As MSForms.Label                          'Target label control

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
            Err.Raise vbObjectError + 513, PROC_NAME, "ParentFrame cannot be Nothing."
        End If

    'Reject an empty control name
        If Len(Trim$(ControlName)) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ControlName cannot be empty."
        End If

'------------------------------------------------------------------------------
' LOOK FOR EXISTING CONTROL
'------------------------------------------------------------------------------
    'Clear the object reference before lookup
        Set ExistingControl = Nothing

    'Temporarily ignore lookup errors when the control does not exist yet
        On Error Resume Next

    'Try to retrieve an existing control with the expected name
        Set ExistingControl = ParentFrame.Controls(ControlName)

    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REUSE EXISTING LABEL
'------------------------------------------------------------------------------
    'Reuse the existing control if it is a label
        If Not ExistingControl Is Nothing Then

            'Reject an existing non-label control with the same name
                If Not TypeOf ExistingControl Is MSForms.Label Then
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Control '" & ControlName & "' already exists but is not an MSForms.Label."
                End If

            'Return the existing label
                Set UF_Ensure_FrameLabel = ExistingControl

            'Exit after returning the existing label
                Exit Function

        End If

'------------------------------------------------------------------------------
' CREATE LABEL IF MISSING
'------------------------------------------------------------------------------
    'Create the label if it does not already exist
        Set Lbl = ParentFrame.Controls.Add("Forms.Label.1", ControlName, True)

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Function

Private Function UF_CalendarGrid_GetWidth() As Single

'
'------------------------------------------------------------------------------
'                           GET CALENDAR GRID WIDTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the rendered width of the 7-column day grid
'
' WHY THIS EXISTS
'   Header, footer, divider, and picker-panel layout routines need a shared
'   calculation for the horizontal span occupied by the calendar day grid
'
' INPUTS
'   None
'
' RETURNS
'   Calendar grid width
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   Calendar layout constants
'
' NOTES
'   Assumes a regular grid with a fixed horizontal step
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN WIDTH
'------------------------------------------------------------------------------
    'Return the full width occupied by the day grid
        UF_CalendarGrid_GetWidth = _
            DP_DAY_LABEL_WIDTH + _
            DP_DAY_LABEL_HORIZONTAL_STEP * (DP_DAY_LABELS_PER_ROW - 1)

End Function

Private Function UF_CalendarGrid_GetBottom() As Single

'
'------------------------------------------------------------------------------
'                           GET CALENDAR GRID BOTTOM
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the bottom coordinate of the 6-row day grid
'
' WHY THIS EXISTS
'   Footer, divider, and picker-panel layout routines need a shared calculation
'   for the vertical bottom edge of the calendar day grid
'
' INPUTS
'   None
'
' RETURNS
'   Bottom coordinate of the day grid
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   Calendar layout constants
'
' NOTES
'   Assumes a regular 6-row calendar grid with equal vertical spacing
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN BOTTOM
'------------------------------------------------------------------------------
    'Return the bottom coordinate of the last day-label row
        UF_CalendarGrid_GetBottom = _
            DP_DAY_GRID_START_TOP + _
            DP_DAY_LABEL_VERTICAL_STEP * (DP_DAY_GRID_ROWS - 1) + _
            DP_DAY_LABEL_HEIGHT

End Function

Private Sub UF_Validate_CalendarLayoutConstants(ByVal CallerName As String)

'
'------------------------------------------------------------------------------
'                           VALIDATE CALENDAR LAYOUT CONSTANTS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the shared DatePicker calendar layout constants
'
' WHY THIS EXISTS
'   The weekday header row and the day-label grid must use consistent dimensions
'   and spacing
'
' INPUTS
'   CallerName
'     Name of the calling routine, used as the error source
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Raises a descriptive runtime error if CallerName is blank or if one or more
'   layout constants are invalid
'
' DEPENDENCIES
'   Calendar layout constants in this module
'
' NOTES
'   This routine validates local layout constants only
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty caller name
        If Len(Trim$(CallerName)) = 0 Then
            Err.Raise vbObjectError + 512, _
                "UF_Validate_CalendarLayoutConstants", _
                "CallerName cannot be empty."
        End If

'------------------------------------------------------------------------------
' VALIDATE DAY LABEL DIMENSIONS
'------------------------------------------------------------------------------
    'Reject invalid day-label width
        If DP_DAY_LABEL_WIDTH <= 0 Then
            Err.Raise vbObjectError + 513, CallerName, _
                "DP_DAY_LABEL_WIDTH must be greater than zero."
        End If

    'Reject invalid day-label height
        If DP_DAY_LABEL_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 514, CallerName, _
                "DP_DAY_LABEL_HEIGHT must be greater than zero."
        End If

'------------------------------------------------------------------------------
' VALIDATE DAY LABEL SPACING
'------------------------------------------------------------------------------
    'Reject invalid day-label horizontal spacing
        If DP_DAY_LABEL_HORIZONTAL_STEP <= 0 Then
            Err.Raise vbObjectError + 515, CallerName, _
                "DP_DAY_LABEL_HORIZONTAL_STEP must be greater than zero."
        End If

    'Reject invalid day-label vertical spacing
        If DP_DAY_LABEL_VERTICAL_STEP <= 0 Then
            Err.Raise vbObjectError + 516, CallerName, _
                "DP_DAY_LABEL_VERTICAL_STEP must be greater than zero."
        End If

'------------------------------------------------------------------------------
' VALIDATE WEEKDAY LABEL DIMENSIONS
'------------------------------------------------------------------------------
    'Reject invalid weekday-label width
        If DP_DOW_LABEL_WIDTH <= 0 Then
            Err.Raise vbObjectError + 517, CallerName, _
                "DP_DOW_LABEL_WIDTH must be greater than zero."
        End If

    'Reject invalid weekday-label height
        If DP_DOW_LABEL_HEIGHT <= 0 Then
            Err.Raise vbObjectError + 518, CallerName, _
                "DP_DOW_LABEL_HEIGHT must be greater than zero."
        End If

'------------------------------------------------------------------------------
' VALIDATE WEEKDAY LABEL SPACING
'------------------------------------------------------------------------------
    'Reject invalid weekday-label horizontal spacing
        If DP_DOW_LABEL_HORIZONTAL_STEP <= 0 Then
            Err.Raise vbObjectError + 519, CallerName, _
                "DP_DOW_LABEL_HORIZONTAL_STEP must be greater than zero."
        End If

End Sub

Private Function UF_WeekdayCaption_GetFixedEnglish(ByVal DayNumber As Long) As String

'
'------------------------------------------------------------------------------
'                           GET FIXED ENGLISH DAY OF WEEK CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the fixed English weekday caption for a VBA weekday number
'
' WHY THIS EXISTS
'   The DatePicker can operate in local-independent mode, where weekday captions
'   should not depend on Windows or Office regional settings
'
' INPUTS
'   DayNumber
'     VBA weekday number
'
' RETURNS
'   Fixed English weekday caption
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
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "UF_WeekdayCaption_GetFixedEnglish"    'Current procedure name

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject invalid day numbers
        If DayNumber < vbSunday Or DayNumber > vbSaturday Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DayNumber must be between vbSunday and vbSaturday."
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

End Function





