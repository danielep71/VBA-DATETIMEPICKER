# VBA-DATETIMEPICKER

<p align="center">
  <b>A modern, reusable, worksheet-friendly Date / Time Picker for Excel VBA</b><br>
  Clean UI • Smart cell detection • In-grid activation • Modeless workflow • Manager-driven architecture
</p>

<p align="center">
  <img alt="Platform" src="https://img.shields.io/badge/Platform-Excel_VBA-217346">
  <img alt="Office" src="https://img.shields.io/badge/Office-32%2F64--bit-blue">
  <img alt="Architecture" src="https://img.shields.io/badge/Architecture-Manager_driven-6f42c1">
  <img alt="UI" src="https://img.shields.io/badge/UI-Modeless_UserForm-00A3E0">
  <img alt="Ribbon" src="https://img.shields.io/badge/Ribbon_Callbacks-Supported-217346">
  <img alt="License" src="https://img.shields.io/badge/License-MIT-green">
  <img alt="Status" src="https://img.shields.io/badge/Status-Final-green">
</p>

<p align="center">
  <b>🔒 No add-in required • No admin rights — import the source and the picker lives inside your workbook, or install the optional <code>.xlam</code>.</b>
</p>

---

<img width="1536" height="1024" alt="datepicker-home2" src="https://github.com/user-attachments/assets/21bda6f0-d738-47b9-87f1-69f83c308ba7" />

---


## ✨ Overview

<p align="left">
  <img alt="Component" src="https://img.shields.io/badge/Component-Date_%2F_Time_Picker-217346">
  <img alt="Excel UX" src="https://img.shields.io/badge/Excel_UX-Worksheet_friendly-blue">
</p>

**VBA-DATETIMEPICKER** is a modern, modular, and worksheet-friendly **Date / Time Picker for Excel VBA**.

The component provides a **modeless UserForm**, contextual worksheet integration, in-form settings, keyboard shortcuts, live-clock support, right-click access, optional in-grid activation, and optional Ribbon callbacks.

This project focuses on:

- 📅 better date and date-time entry UX in Excel
- 🧩 modular, maintainable VBA architecture
- ⚙️ reusable infrastructure for enterprise workbooks and add-ins
- 🧪 demo-ready and regression-friendly implementation patterns
- 🛡️ defensive lifecycle cleanup for real Excel environments

---

## ⭐ Why this exists

<p align="left">
  <img alt="Problem" src="https://img.shields.io/badge/Problem-Native_date_input-lightgrey">
  <img alt="Goal" src="https://img.shields.io/badge/Goal-Modern_Excel_UX-brightgreen">
  <img alt="Audience" src="https://img.shields.io/badge/Audience-VBA_developers_%2F_Excel_users-blueviolet">
</p>

**Built for locked-down environments.** This project exists because of a specific, very common
problem: on company-managed computers you often can't install anything — no admin rights, no
add-ins, no approved IT request in time. Most Excel date pickers are distributed as installable
add-ins, which makes them a non-starter for many corporate users and breaks shared workbooks for
anyone who doesn't have the add-in.

`VBA-DATETIMEPICKER` takes the opposite approach: it's plain VBA source you import directly into a
workbook. If you can run macros, you can use it — no installation, no admin rights, nothing to
deploy, and nothing for colleagues to install when you share the file. The picker travels inside
the workbook itself.

Excel still powers an enormous number of business-critical workflows, but native date-entry UX is often poor.

This repository exists to provide a **clean, modern, and reusable VBA-based solution** for a very common problem:

> making date input in Excel feel much better without sacrificing control, portability, or maintainability.

The project is intentionally more than a visual widget. It includes manager-driven event handling, settings persistence, workbook lifecycle cleanup, runtime control creation, and a clear separation between **Excel integration**, **form rendering**, and **label-event routing**.

---

## 🚀 Key features

<p align="left">
  <img alt="Modeless UI" src="https://img.shields.io/badge/Modeless_UI-Supported-217346">
  <img alt="Draggable Form" src="https://img.shields.io/badge/Borderless_Drag-Supported-217346">
  <img alt="Date and Time" src="https://img.shields.io/badge/Date_%2F_Time_Write--Back-Supported-217346">
  <img alt="Smart Detection" src="https://img.shields.io/badge/Smart_Cell_Detection-Supported-217346">
  <img alt="In-grid Icon" src="https://img.shields.io/badge/In--Grid_Icon-Supported-217346">
  <img alt="Right Click" src="https://img.shields.io/badge/Right--Click_Menu-Supported-217346">
  <img alt="Keyboard Shortcut" src="https://img.shields.io/badge/Keyboard_Shortcut-Supported-217346">
  <img alt="Ribbon Callbacks" src="https://img.shields.io/badge/Ribbon_Callbacks-Supported-217346">
  <img alt="Keyboard Navigation" src="https://img.shields.io/badge/Keyboard_Navigation-Supported-217346">
  <img alt="Live Clock" src="https://img.shields.io/badge/Live_Clock-Supported-217346">
  <img alt="Compact Mode" src="https://img.shields.io/badge/Compact_Mode-Supported-217346">
  <img alt="Lifecycle Cleanup" src="https://img.shields.io/badge/Lifecycle_Cleanup-Built--In-orange">
  <img alt="Settings Panel" src="https://img.shields.io/badge/Settings-In--Form_Panel-orange">
  <img alt="Demo Workbook" src="https://img.shields.io/badge/Demo_Workbook-Available-orange">
</p>

### 📌 Worksheet-friendly Date / Time input

The picker is designed for real Excel workflows, not only for isolated form demos.

- 📅 **Date-only write-back** through the Today shortcut and calendar selection
- 🕒 **Date-time write-back** through the Now shortcut
- 📌 **Modeless UserForm**, so Excel remains usable while the picker is open
- 🔁 **Repeated write-back workflow**, useful when users move across multiple target cells
- 🧾 **Table-column write behavior** for structured data-entry ranges

> ℹ️ **A selected cell inside an Excel Table data column receives the date on its
> own.** Calendar selection, Today and Now all write only the selected cell. Use
> `DP_FillTableColumn` to fill the whole data column; it reports how many cells
> it would affect before writing.

### 🧠 Smart activation and cell detection

The DatePicker can be shown through multiple entry points and can react to worksheet context.

- 🧠 Detects date-like values and date-formatted cells
- 🖱️ Supports optional **right-click menu** integration
- 🎯 Supports optional **in-grid activation icon** next to eligible cells
- ⌨️ Supports optional **keyboard shortcut** fallback
- 🎛️ Supports optional **Ribbon callbacks** for workbook or add-in Ribbon buttons
- 🧯 Prevents dead access configurations by preserving at least one practical entry path

### 🖱️ In-grid activation icon

The optional worksheet icon gives users a visible, contextual entry point directly in the grid.

- 🎯 Appears near eligible cells
- ⚡ Uses hide / move / reuse behavior for high-frequency selection changes
- 🧹 Can be removed when the feature is disabled
- 🧼 Can be purged during workbook cleanup, save, close, or reset paths

### 🗓️ Calendar rendering and navigation

The picker uses a fixed and predictable calendar surface.

- 🧱 Fixed **6 x 7 calendar grid**
- ◀️ Previous / next month navigation
- 🔼 Previous / next year navigation
- 📆 Month picker overlay
- 📅 Year picker overlay
- 🎯 Today, selected-date, keyboard-date, outside-month, weekend, disabled, and hover visual states

### ⌨️ Keyboard and mouse workflow

The component supports both mouse-first and keyboard-first usage.

- ⬅️ ➡️ ⬆️ ⬇️ Arrow-key day navigation
- 📄 PageUp / PageDown month navigation
- 🏁 Home / End navigation inside the displayed month
- ✅ Enter / Space selection
- ⎋ Esc close / overlay dismiss behavior
- 🔤 Letter shortcuts for Month, Year, Today, and Now actions

### 🪟 Borderless draggable window

The Date / Time Picker can run with a clean borderless visual style while still remaining movable by the user.

- 🪟 Supports optional borderless UserForm styling
- 🖱️ Allows the form to be moved from a custom header drag surface
- 🧭 Keeps normal picker controls clickable while only the passive header area acts as the drag surface
- 🛡️ Falls back safely when WinAPI support is unavailable
- 🧩 Keeps drag behavior separated from calendar, settings, and label-click logic

### ⚙️ In-form settings panel

Settings are exposed inside the picker through a compact overlay panel.

- 🖥️ **Display settings**
  - first day of week
  - local or fixed-English captions
  - live clock
  - compact layout
  - weekend highlighting

- 🧭 **Behavior settings**
  - close after selection
  - allow outside-month selection

- 🔌 **Integration settings**
  - right-click menu
  - in-grid icon
  - WinAPI styling

### 🧱 Clean internal architecture

The project separates responsibilities across focused components.

- 🧠 `cDatePickerManager` handles Excel application events and lifecycle orchestration
- 🖼️ `UF_DatePicker` handles the runtime UserForm, rendering, settings panel, and visual state
- 🔗 `cDatePickerLabelHook` routes events from runtime-created labels
- 🧰 `M_DatePicker` exposes the public API, settings, integration helpers, Ribbon callbacks, timers, and grid-icon infrastructure

### 🧹 Lifecycle and cleanup discipline

The DatePicker includes explicit cleanup paths for real workbook sessions.

- ⏱️ Stops live-clock timers
- 🧼 Closes loaded picker forms
- 🧹 removes or purges in-grid icons
- 🖱️ removes right-click menu entries
- ⌨️ removes keyboard shortcut assignments
- 🔁 repairs or restarts the manager when needed
- 🛡️ preserves the caller's `Application.EnableEvents` state

### 🏢 Deployment-friendly VBA behavior

The implementation is designed for controlled Excel environments.

- 📦 No external installer required
- 🧩 Importable `.bas`, `.cls`, and `.frm` source files
- 🧠 Clear separation between UI, manager, hooks, and infrastructure
- 🧪 Demo-friendly and regression-friendly structure
- 🧾 Compatible with source-controlled VBA exports

Read [Known limitations](#-known-limitations) before deploying into a shared or
business-critical workbook.
  
---

## 📸 Screenshots


### Main picker

<img src="assets/datepicker-main.png" alt="Date / Time Picker main" />

### Settings panel

<img width="3042" height="1238" alt="image" src="https://github.com/user-attachments/assets/e5749a62-2695-4dea-bdca-df4cab925ca8" />

### In-grid activation

<img src="assets/datepicker-ingrid.png" alt="Date / Time Picker in-grid activation icon" />

---

## 🧩 Compatibility 

- Microsoft Excel for Windows
- Excel 2016 and later
- Office 32-bit and 64-bit
- VBA UserForms
- No external runtime dependency

---

## 🏗️ Architecture

<p align="left">
  <img alt="Pattern" src="https://img.shields.io/badge/Pattern-Manager_%2B_Form_%2B_Hooks-6f42c1">
  <img alt="Events" src="https://img.shields.io/badge/Events-Application_level-blue">
  <img alt="Runtime UI" src="https://img.shields.io/badge/Runtime_UI-MSForms-orange">
</p>

The project is organized around a clean separation of responsibilities.

```mermaid
flowchart LR
    Excel[Excel Workbook / Worksheet] --> Manager[cDatePickerManager]
    Manager --> Module[M_DatePicker]
    Module --> Form[UF_DatePicker]
    Form --> Hooks[cDatePickerLabelHook]
    Hooks --> Form
    Module --> Settings[Registry-backed Settings]
    Module --> UI[Context Menu / Keyboard / Grid Icon / Ribbon / Timer]
```

### `M_DatePicker.bas`

Provides the shared companion module for the DatePicker project

- `DP_Start`, `DP_Stop`, `DP_Show`, `DP_Close`
- `DP_Today`, `DP_Now`
- manager bootstrap through `M_Picker_EnsureManager`
- settings load/save and registry persistence
- context-menu integration
- keyboard shortcut integration
- in-grid icon creation, move, hide, remove, and purge
- live-clock timer infrastructure
- optional WinAPI styling and mouse positioning helpers
- selected-date write-back policy and form bridge state

### `cDatePickerManager.cls`

Provides the central Excel Application event manager for the DatePicker project

The DatePicker uses several transient Excel / VBA UI surfaces:
  - modeless runtime UserForm
  - in-grid DatePicker icon
  - workbook, worksheet, window, save, print, close, and activation events
  - optional keyboard shortcut fallback

Centralizing the Application event flow in one manager class keeps the
project predictable, avoids duplicate event wiring, and prevents stale UI
artifacts from surviving normal Excel context transitions

The controller is responsible for:

- Excel Application event hooks
- selection-driven DatePicker refresh
- stale UI cleanup
- workbook/worksheet context handling
- reentrancy protection
- high-frequency in-grid icon show / move / hide behavior
- hard-boundary cleanup before save, print, close, or teardown


### `UF_DatePicker.frm`

Dedicated UI layer responsible for:

- UserForm shell formatting
- runtime creation and reuse of labels, frames, pages, MultiPage controls, ComboBoxes, and CheckBoxes
- header month/year captions and navigation
- fixed 6 × 7 calendar-grid rendering
- paired day-cell labels for larger hover and click targets
- month/year picker overlay
- in-form settings overlay
- Today / Now footer actions
- compact layout behavior
- live footer-clock display
- keyboard navigation and shortcut routing

The DatePicker form is built at runtime. This keeps layout, dynamic control creation, 
label-event routing, calendar rendering, month/year navigation, footer shortcuts, 
in-form settings, keyboard  navigation, and live-clock behavior centralized in one 
UserForm code-behind.

### `cDatePickerLabelHook.cls`

Runtime-created MSForms labels do not expose normal UserForm event
procedures. A WithEvents class is required to capture their Click and
MouseMove events.

It is responsible for:

- click routing for runtime-created MSForms labels
- day-label hover routing
- header-label hover routing
- picker-panel hover routing
- footer-label hover routing
- settings-panel hover routing

---

## 🧠 Design notes

<p align="left">
  <img alt="Design" src="https://img.shields.io/badge/Design-Excel_first-217346">
  <img alt="Safety" src="https://img.shields.io/badge/Safety-Defensive_cleanup-blue">
  <img alt="Maintainability" src="https://img.shields.io/badge/Maintainability-Modular-6f42c1">
</p>

This repo follows a few core principles:

- **Excel-first UX**  
  The picker should feel natural inside worksheet workflows.

- **Reusable VBA engineering**  
  The solution should be portable across workbooks and projects.

- **Separation of concerns**  
  UI rendering, orchestration, infrastructure, and event routing remain clearly separated.

- **Safe event handling**  
  Reentrancy guards, explicit cleanup, and Application-level event ownership matter in real Excel environments.

- **Settings are explicit**  
  Runtime settings are changed through the in-form Settings panel and persisted through an explicit Save action.

- **High-frequency UI paths stay cheap**  
  The in-grid icon is hidden / moved / reused during selection refreshes, while explicit disable / teardown paths remove or purge it.

- **Production-minded code quality**  
  This is intended as a robust, reusable component, not only a demo widget.

---

## 📦 Repository structure

<p align="left">
  <img alt="Source" src="https://img.shields.io/badge/Source-src%2F-blue">
  <img alt="Forms" src="https://img.shields.io/badge/Forms-.frm_%2B_.frx-orange">
  <img alt="Demo" src="https://img.shields.io/badge/Demo-Workbook-lightgrey">
</p>

```text
VBA-DATETIMEPICKER/
├─ src/
│  ├─ modules/
│  │  └─ M_DatePicker.bas
│  ├─ classes/
│  │  ├─ cDatePickerManager.cls
│  │  └─ cDatePickerLabelHook.cls
│  ├─ forms/
│  │  ├─ UF_DatePicker.frm
│  │  └─ UF_DatePicker.frx
│  └─ ribbon/
│     └─ customUI14.xml
├─ demo/
│  ├─ M_DEMO_BUILDER.bas
│  └─ M_DP_DEMO.bas
├─ assets/
│  ├─ datepicker-main.png
│  ├─ datepicker-settings.png
│  ├─ datepicker-ingrid.png
│  └─ vba-datetimepicker-social-preview.png
├─ README.md
└─ LICENSE
```
> Important: if the exported UserForm references `UF_DatePicker.frx`, keep the `.frm` and `.frx` together in source control.
> 
> Ribbon support is callback-based. The VBA callbacks live in a standard module, while the Ribbon layout itself must be provided through RibbonX / `customUI14.xml` in the workbook or add-in package.
---

## 🛠️ Installation

<p align="left">
  <img alt="Install" src="https://img.shields.io/badge/Install-Import_Files-217346">
  <img alt="Add-in Available" src="https://img.shields.io/badge/Add--in-.xlam_available-6f42c1">
  <img alt="VBE" src="https://img.shields.io/badge/VBE-Import-orange">
</p>

> **Which path is for you?** On a managed machine where you can't install add-ins — no admin rights, IT-controlled Excel — use the **source-import path (Option 1)**: it needs no installation and no special permissions, so if you can run macros, you can use it. The **add-in (Option 3)** is a convenience for users who *can* install Excel add-ins and want the picker available across every workbook without importing the source each time.

### Option 1 — Import into an existing VBA project

1. Open the target workbook in Excel.
2. Open the VBA Editor with `Alt + F11`.
3. Import the source files:
   - `M_DatePicker.bas`
   - `cDatePickerManager.cls`
   - `cDatePickerLabelHook.cls`
   - `UF_DatePicker.frm`
   - `UF_DatePicker.frx`, if referenced by the exported form
4. Ensure the **Microsoft Forms 2.0 Object Library** is available.
5. If Ribbon callbacks are used, ensure the workbook or add-in includes the matching RibbonX `customUI14.xml`.
6. Add the workbook lifecycle calls shown below.
7. Compile the VBA project.
8. Save as `.xlsm`, `.xlsb`, or package into your preferred add-in/deployment format.

### Option 2 — Use the demo workbook

Download **`DATETIMEPICKER-demo-vx.y.z.xlsm`** from the
[latest release](https://github.com/danielep71/vba-datetimepicker/releases/latest),
where `vx.y.z` is the release version. Right-click the file → **Properties** →
tick **Unblock** → **OK** before opening it.

Use it to:

- explore the picker behavior
- compare table write scopes: pick a date in one **Expiry Date** cell, then use
  **Fill Table Column**
- validate the interaction model
- test settings persistence
- test right-click, keyboard, and in-grid icon entry points
- use the demo as a starting point for integration

The workbook is not stored in this repository. It is generated from
`demo/M_DP_DEMO.bas` at release time, so if you are working from source you can
rebuild it yourself:

```vba
DP_Demo_CreateDemoSheet
```

### Option 3 — Install the prebuilt add-in (`.xlam`)

For end users who just want the picker available across Excel, without touching VBA source:

1. Download **`DATETIMEPICKER-vx.y.z.xlam`** from the [latest release](https://github.com/danielep71/vba-datetimepicker/releases/latest), where `vx.y.z` is the release version.
2. Right-click the file → **Properties** → tick **Unblock** → **OK** (clears the downloaded-from-internet block).
3. In Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…**, select the `.xlam`, and enable it.
4. The Date / Time Picker is then available through its Ribbon button, right-click entry, and `Ctrl + Shift + D`.

> **Requires the ability to install Excel add-ins.** On a locked-down machine (no admin rights, IT-managed Excel) this option won't be available — use the source-import path (Option 1) instead, which needs no installation. The add-in suits users who *can* install it and want the picker in every workbook; to redistribute the picker inside a shared workbook, also use Option 1 so the code travels with the file.

---

## 🔁 Workbook lifecycle

<p align="left">
  <img alt="Lifecycle" src="https://img.shields.io/badge/Lifecycle-Open_%2F_Save_%2F_Print_%2F_Close-blue">
  <img alt="Startup" src="https://img.shields.io/badge/Startup-DP_Start-217346">
  <img alt="Shutdown" src="https://img.shields.io/badge/Shutdown-DP_Stop-orange">
</p>

The workbook lifecycle should stay thin. Let `cDatePickerManager` own Application-level events, and use workbook events only to start and stop the component.

Recommended `ThisWorkbook` pattern:

```vba
Option Explicit

Private Sub Workbook_Open()
    DP_Start
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DP_Stop
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    If Cancel Then Exit Sub

    DP_Close
    M_GridIcon_PurgeAll

End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)

    If Cancel Then Exit Sub

    DP_Close
    M_GridIcon_PurgeAll

End Sub
```

Do **not** duplicate worksheet `SelectionChange` logic in `ThisWorkbook`. The manager already handles selection changes through Application-level events.

---

## ⚙️ Configuration

<p align="left">
  <img alt="Settings" src="https://img.shields.io/badge/Settings-Registry_backed-6f42c1">
  <img alt="Panel" src="https://img.shields.io/badge/Panel-In--form-blue">
  <img alt="Persistence" src="https://img.shields.io/badge/Persistence-Explicit_Save-217346">
</p>

Settings are stored using VBA registry persistence under:

```text
HKCU\Software\VB and VBA Program Settings\VBA_DATETIMEPICKER
```

The legacy settings application name is intentionally retained as:

```text
VBA_DATETIMEPICKER
```

This preserves compatibility with existing saved settings.

### Persistence is scoped to the Windows user, not the workbook

By default every DatePicker deployment under the same Windows account reads and
writes that one location. Two workbooks that never run at the same time can still
change each other's saved preferences, because loading settings normalizes and
re-saves them.

Some of what is stored is a reasonable user-wide preference — first day of week,
local names, weekend highlighting. Some describes one deployment: the right-click
entry, the in-grid icon, the keyboard shortcut, WinAPI styling, and the holiday
callback name, which the picker later executes.

### Isolating one deployment's settings

Configure a namespace **before anything loads settings**, which includes ordinary
DatePicker startup:

```vb
M_Settings_SetNamespace "TreasuryTool"
```

That deployment then persists under:

```text
HKCU\Software\VB and VBA Program Settings\VBA_DATETIMEPICKER__TreasuryTool
```

Points worth knowing:

- **The default is unchanged.** Configure no namespace and the add-in reads and
  writes exactly where it always did.
- **The namespace locks once settings load.** Changing it afterwards is refused,
  because values read from one namespace would then be written into another.
- **A new namespace starts from defaults.** Nothing is copied from the shared
  location — importing it would carry across the very settings the namespace
  exists to separate, the holiday callback included.
- **You choose the namespace.** It is never derived from the workbook name or
  path, which change when a file is renamed, moved or copied and would make
  settings appear to vanish. The release version is not used either, which would
  make every upgrade look like a reset.
- **This is not the same as running two copies at once.** Namespacing settings
  does not make two simultaneous DatePicker providers supported; see
  [#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37).

Typical settings include:

| Area | Setting | Description |
|---|---|---|
| Display | First day of week | Sunday or Monday |
| Display | Local names | Use local month/day names or fixed English captions |
| Display | Live clock | Static or live footer time display |
| Display | Compact layout | Shorter picker layout with header settings access |
| Display | Highlight weekends | Bold weekend day labels |
| Behavior | Allow outside-month selection | Allow adjacent-month dates visible in the grid to be selected |
| Behavior | Close after selection | Close the picker after a successful write-back |
| Integration | Right-click menu | Add Date / Time Picker to the Excel cell context menu |
| Integration | In-grid icon | Show a contextual worksheet activation icon |
| Integration | WinAPI styling | Use optional Windows-specific styling such as title-bar removal |
| Integration | Keyboard shortcut | Fallback entry point when visual/contextual entry points are disabled |

---

## 🧩 Public API

<p align="left">
  <img alt="API" src="https://img.shields.io/badge/API-Public_entry_points-blue">
  <img alt="Manager" src="https://img.shields.io/badge/Manager-Auto_bootstrap-217346">
  <img alt="Write-back" src="https://img.shields.io/badge/Write--back-Date_%2F_Datetime-orange">
</p>

Primary public entry points are exposed through `M_DatePicker.bas`.

| Procedure | Purpose |
|---|---|
| `DP_Start` | Starts DatePicker runtime integrations and ensures the manager is available |
| `DP_Stop` | Stops runtime integrations and clears transient UI artifacts |
| `DP_Show` | Opens or refreshes the Date / Time Picker for the active target context |
| `DP_Close` | Closes the picker and stops form-level runtime activity |
| `DP_Preload` | Loads and hides the picker once so the first open is instant; failures are suppressed through a fail-safe path |
| `DP_Hide` | Hides the picker while keeping it loaded for fast reuse, unlike `DP_Close` which tears it down |
| `DP_Today` | Writes today’s date to the current target |
| `DP_Now` | Writes today’s date with the current system time to the current target |
| `DP_FillTableColumn` | Writes one date to every cell of the Excel Table data column containing the selection, after confirming the scope |
| `DP_RepairRuntime` | Repairs runtime state, including event enablement and manager recreation |
| `M_Picker_EnsureManager` | Ensures settings are loaded and creates or repairs the manager. Reports the caller's `Application.EnableEvents` state without changing it |
| `M_ContextMenu_Update` | Synchronizes right-click menu integration with current settings |
| `M_KeyboardShortcut_Update` | Synchronizes keyboard shortcut integration with current settings |
| `M_GridIcon_PurgeAll` | Removes all DatePicker grid icons from open workbooks |
| `Ribbon_ShowPicker` | Ribbon callback that opens or refreshes the Date / Time Picker |
| `Ribbon_Reset` | Ribbon callback that repairs DatePicker runtime state |
| `Ribbon_Demo` | Ribbon callback that activates the DatePicker demo worksheet |

A normal consumer should usually call only:

```vba
DP_Start
DP_Show
DP_Close
DP_Stop
```

The manager class should normally be treated as internal infrastructure rather than directly instantiated by workbook code.
Ribbon callback procedures must remain `Public` and must be located in a standard VBA module because RibbonX `onAction` callbacks cannot call private procedures, class methods, or UserForm code-behind procedures directly.

---

## 🖱️ Entry points

<p align="left">
  <img alt="Entry" src="https://img.shields.io/badge/Entry_points-Multiple-otange">
  <img alt="Ribbon" src="https://img.shields.io/badge/Ribbon_Callbacks-Supported-217346">
  <img alt="Keyboard Shortcut" src="https://img.shields.io/badge/Ctrl_%2B_Shift_%2B_D-Supported-217346">
  <img alt="Context Menu" src="https://img.shields.io/badge/Context_Menu-Supported-217346">
  <img alt="Grid Icon" src="https://img.shields.io/badge/Grid_icon-Supported-217346">
</p>

The Date / Time Picker can be opened through several user-facing paths:

- Ribbon button through `Ribbon_ShowPicker`
- right-click menu entry
- in-grid icon next to eligible cells
- keyboard shortcut through `Ctrl + Shift + D`
- direct macro call to `DP_Show`

Each of the three built-in interactive paths — right-click, in-grid icon and
keyboard shortcut — is independently configurable, and all three may be disabled
at once. The picker is then reached through `DP_Show`, `Ribbon_ShowPicker`, or a
button or macro you provide.

The project intentionally keeps `VBA_DATETIMEPICKER` as a stable internal command-bar tag / settings name, while user-facing captions should use:

```text
Date / Time Picker
```

---

## 🎛️ Ribbon integration

<p align="left">
  <img alt="Ribbon" src="https://img.shields.io/badge/RibbonX-Callback_ready-217346">
  <img alt="Callbacks" src="https://img.shields.io/badge/Callbacks-Public_Sub-blue">
  <img alt="Standard Module" src="https://img.shields.io/badge/Location-Standard_Module-orange">
</p>

The project can be exposed through Excel Ribbon buttons by mapping RibbonX `onAction` callbacks to public VBA procedures.

Ribbon support is intentionally callback-based:

- 🎛️ Ribbon layout belongs to the workbook / add-in RibbonX XML
- 🧰 callback procedures live in a standard VBA module
- 🧠 the callbacks delegate to the normal DatePicker public API
- 🧹 runtime repair remains available through a dedicated Ribbon action
- 📘 the demo worksheet can be opened from the Ribbon when included in the host workbook

### Recommended callbacks

| Callback | Purpose |
|---|---|
| `Ribbon_ShowPicker` | Opens or refreshes the Date / Time Picker |
| `Ribbon_Reset` | Repairs DatePicker runtime state through `DP_RepairRuntime` |
| `Ribbon_Demo` | Activates the `DATE PICKER DEMO` worksheet |

### Example RibbonX mapping

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab idMso="TabHome">
        <group id="grpDateTimePicker"
               label="DateTime Picker">
          <splitButton id="spbDateTimePicker" size="large">
            <button id="btnShowPicker"
                    label="Show Picker"
                    image="DP_GridIcon_64"
                    onAction="Ribbon_ShowPicker"/>

            <menu id="mnuDateTimePicker"
                  itemSize="large">
              <button id="BtnMenu_Reset"
                      label="Reset"
                      image="reset"
                      onAction="Ribbon_Reset"/>

              <button id="BtnMenu_Demo"
                      label="Demo"
                      image="demo"
                      onAction="Ribbon_Demo"/>

            </menu>
          </splitButton>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

### Callback implementation policy

Ribbon callbacks should stay thin. They should not duplicate DatePicker logic.

Recommended pattern:

```vba
Public Sub Ribbon_ShowPicker(ByVal Control As IRibbonControl)
    DP_Show
End Sub

Public Sub Ribbon_Reset(ByVal Control As IRibbonControl)
    DP_RepairRuntime
End Sub
```

For production use, wrap callbacks with controlled error handling and user-facing diagnostics so Ribbon failures do not fail silently.

---

## 🧭 Keyboard navigation

<p align="left">
  <img alt="Keyboard" src="https://img.shields.io/badge/Keyboard-Supported-217346">
  <img alt="Navigation" src="https://img.shields.io/badge/Navigation-Arrow_Keys-blue">
  <img alt="Shortcuts" src="https://img.shields.io/badge/Shortcuts-T_%2F_N_%2F_M_%2F_Y-orange">
</p>

The picker can also be opened directly from Excel through the configured keyboard shortcut: Ctrl + Shift + D
This shortcut is intended as a fallback entry point when the user does not want to rely on the right-click menu, Ribbon button, or in-grid icon.

Typical keyboard behavior includes:

| Key | Action |
|---|---|
| `Ctrl + Shift + D` | Open or refresh the Date / Time Picker from Excel |
| `←` / `→` | Move one day backward / forward |
| `↑` / `↓` | Move one week backward / forward |
| `PageUp` / `PageDown` | Move one month backward / forward |
| `Ctrl + PageUp` / `Ctrl + PageDown` | Move one year backward / forward |
| `Home` / `End` | Move to first / last day of displayed month |
| `Enter` / `Space` | Select the keyboard-highlighted date |
| `M` | Open month picker |
| `Y` | Open year picker |
| `T` | Write Today |
| `N` | Write Now |
| `Esc` | Hide overlay or close picker |

Letter shortcuts are ignored when Ctrl or Alt is pressed so the picker does not accidentally steal broader Excel or add-in shortcuts.

---

## 🧹 Runtime cleanup policy

<p align="left">
  <img alt="Cleanup" src="https://img.shields.io/badge/Cleanup-Explicit-blue">
  <img alt="Grid Icon Policy" src="https://img.shields.io/badge/Grid_Icon-Hide_%2F_Remove_%2F_Purge-6f42c1">
</p>

The project separates high-frequency UI cleanup from hard lifecycle cleanup:

| Situation | Recommended behavior |
|---|---|
| Selection refresh / non-eligible cell | Hide the in-grid icon |
| Eligible-cell selection change | Show, move, or reuse the in-grid icon |
| Feature disabled in settings | Remove the tracked icon |
| Workbook save / print / close | Purge all DatePicker icons |
| Form close / teardown | Stop timers and release form-level references |
| Workbook close / add-in unload | `DP_Stop` |

This avoids flicker during normal selection changes while still keeping workbook files clean when saving, printing, or closing.

---

## 🪟 WinAPI behavior

<p align="left">
  <img alt="Windows" src="https://img.shields.io/badge/Windows-Optional_support-blue">
  <img alt="WinAPI" src="https://img.shields.io/badge/WinAPI-Optional_styling-6f42c1">
</p>

The picker supports optional Windows-specific behavior for a more polished UI.

The intended policy is:

| Helper | Meaning |
|---|---|
| `M_Platform_CanUseWinAPI` | Platform capability check only |
| `M_Platform_ShouldUseWinAPI` | Capability plus user setting for optional styling |
| `M_Window_RemoveTitleBar` | Uses the styling policy |
| `M_Window_MoveFormToMouse` | Uses capability only, so positioning may still work when styling is disabled |
| `M_Window_BeginUserFormDrag` | Starts native drag movement from a custom borderless form surface |

When the native title bar is removed, the form can still be moved through a custom drag surface. In the default UI, this should be assigned only to a passive header area, not to active controls such as month, year, navigation arrows, or settings icons.
This preserves the clean borderless look while keeping the UserForm practical in daily use.

Disabling **Use WinAPI styling** should disable optional styling such as title-bar removal. It should not necessarily disable all safe Windows capability paths such as mouse-based positioning.

---

## 🧪 Demo quick start

<p align="left">
  <img alt="Demo" src="https://img.shields.io/badge/Demo-DATE_PICKER_DEMO-217346">
  <img alt="Validation" src="https://img.shields.io/badge/Validation-Manual_Scenarios-blue">
  <img alt="Single Cell" src="https://img.shields.io/badge/Single--cell-Tested-green">
  <img alt="Date Formats" src="https://img.shields.io/badge/Date_Formats-Showcase-6f42c1">
  <img alt="Multi Cell" src="https://img.shields.io/badge/Multi--Cell-Write--back-orange">
  <img alt="Tables" src="https://img.shields.io/badge/Excel_Tables-Supported-green">
  <img alt="Protected Sheets" src="https://img.shields.io/badge/Protected_Sheet-Note-lightgrey">
</p>

The demo workbook includes a dedicated worksheet named **`DATE PICKER DEMO`**.


### Demo worksheet

<img src="assets/datepicker-demosheet.png" alt="Date / Time Picker demo worksheet" />

This sheet is designed as a practical validation surface for the Date / Time Picker. It is not only a visual showcase: it provides structured scenarios for testing cell eligibility, date and datetime formats, repeated write-back, multi-cell selection, and Excel Table behavior.

### 🧭 How to use the demo sheet

1. Select a demo input cell.
2. Invoke the Date / Time Picker using one of the available entry points:
   - 🎛️ Ribbon button
   - 🖱️ right-click menu
   - 🎯 in-grid activation icon
   - ⌨️ `Ctrl + Shift + D`
   - 🧰 direct call to `DP_Show`
3. Pick a date, use **Today**, or use **Now**.
4. Validate the result against the expected behavior shown on the demo sheet.

### 📌 Basic single-cell scenarios

The left-hand scenario block validates the most common worksheet inputs:

| Scenario | Expected behavior |
|---|---|
| Empty date-formatted cell | Positive — picker should be available |
| Pre-filled date | Positive — picker should initialize from the existing date |
| Pre-filled datetime | Positive — datetime handling should be preserved |
| Empty general-format cell | Usually negative — general format should not automatically qualify |
| Text cell | Negative — non-date text should not qualify |
| Formula returning date | Positive — date-like formula result should qualify |
| Empty datetime-formatted cell | Positive — formatted for date-time entry |

### 🧾 Format showcase

The format showcase validates that the picker recognizes common date and datetime display patterns, including:

- `dd/mm/yyyy`
- `dd-mmm-yyyy`
- `dddd, dd mmmm yyyy`
- `dd/mm/yyyy hh:mm`
- `hh:mm:ss`
- `mm/dd/yyyy`
- `yyyy-mm-dd`

This section is useful for checking whether the DatePicker detection logic behaves consistently across different date formats and local display conventions.

### 🔁 Multi-cell selection

The multi-cell section validates repeated write-back behavior across selected ranges.

Typical checks include:

- selecting a valid multi-cell input range
- confirming that all eligible target cells receive the selected date
- confirming that invalid or unsupported selections are rejected or ignored according to policy
- testing repeated date entry without reopening the workbook

### 📊 Excel Table demo

The Excel Table section validates table-friendly behavior using structured business-style fields such as:

- Trade Date
- Settlement Date
- Expiry Date
- Timestamp

This is useful for testing real data-entry workflows where the picker is used inside structured tables rather than isolated worksheet cells.

### 🛡️ Protected-sheet note

The in-grid DatePicker icon is implemented as a worksheet `Shape`.

On protected sheets, the icon may not appear even when the target cell is unlocked. For the icon to remain available under sheet protection, the sheet protection settings must allow drawing objects.

For example, protection should be configured consistently with:

```vba
DrawingObjects:=False
UserInterfaceOnly:=True
```
UserInterfaceOnly:=True is useful for allowing VBA operations while keeping normal user protection active.

✅ Suggested validation checklist

Use the demo sheet to validate:

- single-cell date detection
- datetime write-back
- empty date-formatted cells
- pre-filled date and datetime cells
- rejection of non-date text cells
- formula results that return valid dates
- multiple display formats
- Today date-only write-back
- Now date-time write-back
- multi-cell write-back
- Excel Table date-entry behavior
- in-grid icon show / move / hide behavior
- right-click menu behavior
- Ribbon callback behavior
- Ctrl + Shift + D shortcut behavior
- settings persistence
- compact and normal layout behavior
- WinAPI enabled / disabled behavior
- runtime repair through DP_RepairRuntime
- caller event state preserved across DP_Start, DP_Show and DP_Preload


## ✅ Test quick start

<p align="left">
  <img alt="Tests" src="https://img.shields.io/badge/Tests-Compile_%2B_Manual_%2B_Demo_Worksheet-blue">
  <img alt="Shortcut" src="https://img.shields.io/badge/Ctrl_%2B_Shift_%2B_D-Supported-217346">
  <img alt="Ribbon" src="https://img.shields.io/badge/Ribbon_Callbacks-Optional-217346">
  <img alt="Drag" src="https://img.shields.io/badge/Borderless_Drag-Supported-6f42c1">
</p>

Use this checklist before publishing a release.

1. Compile the VBA project.
2. Run `DP_Start`.
3. Confirm `Application.EnableEvents` is unchanged after manager initialization.
4. Build the demo worksheet with `DP_Demo_CreateDemoSheet`, or open it from the Ribbon Demo button if RibbonX is included.
5. Test the **Basic Single-Cell Scenarios** block:
   - empty date-formatted cell
   - pre-filled date
   - pre-filled datetime
   - empty general-format cell
   - text cell
   - formula returning date
   - empty datetime-formatted cell
6. Test the **Format Showcase** block across the listed date and datetime formats.
7. Test the **Multi-Cell Selection** block and confirm repeated write-back behavior.
8. Test the **Excel Table Demo** block: picking a date in one cell writes that cell only, and the **Fill Table Column** button fills the column after confirming the scope.
9. Select eligible and non-eligible cells.
10. Confirm the in-grid icon shows, moves, hides, removes, and purges according to policy.
11. Open the picker with `DP_Show`.
12. Test day selection, Today, and Now.
13. Test month and year navigation.
14. Test settings save and persistence.
15. Test compact layout and normal layout.
16. Test WinAPI styling enabled and disabled.
17. Test that `Ctrl + Shift + D` opens or refreshes the picker.
18. Test that the borderless form can be moved from the custom header drag surface.
19. Test Ribbon callbacks if RibbonX is included:
   - Show Picker
   - Repair Runtime
   - Demo sheet toggle
20. Test `DP_RepairRuntime` and confirm the runtime remains usable afterward.
21. Test print-preview and workbook close to verify cleanup.
22. Reopen the workbook and confirm startup is clean.
23. Confirm the demo sheet is very hidden on workbook open and after workbook close.
24. Confirm AutoSave / Save does not unexpectedly hide the demo sheet while the user is actively using it.

Recommended Immediate Window checks:

```vba
? Application.EnableEvents

DP_Start
DP_Show
DP_Close
DP_RepairRuntime
DP_Stop

M_Picker_EnsureManager
M_GridIcon_PurgeAll
DP_Demo_CreateDemoSheet
DP_DemoSheet_Show
DP_DemoSheet_HideVeryHidden
```

Ribbon callbacks require RibbonX and an IRibbonControl callback context. Test them by clicking the actual Ribbon buttons rather than calling the callback procedures directly from the Immediate Window.

---

## 🧯 Troubleshooting


| Symptom | Check |
|---|---|
| Selection changes do not trigger picker behavior | Confirm `Application.EnableEvents = True`, then run `DP_Start`. If a macro disabled events, run `DP_RepairRuntime` |
| Form import fails or looks incomplete | Ensure `UF_DatePicker.frm` and `UF_DatePicker.frx` are both present |
| Right-click menu entry does not appear | Check settings, then run `M_ContextMenu_Update` |
| In-grid icon remains after disabling the feature | Run `M_GridIcon_PurgeAll`, then save settings again |
| Picker does not open near the mouse | Confirm platform capability and WinAPI declarations |
| Borderless styling does not apply | Confirm `gDP_UseWinAPI = True` and platform support |
| Settings appear stale | Open settings panel and use Save, or reload with `M_Settings_EnsureLoaded` |
| Events were disabled by another macro | Run `DP_RepairRuntime`. It is the only routine that deliberately re-enables events |
| Ribbon button does nothing | Confirm the RibbonX `onAction` name exactly matches the public VBA callback name |
| Ribbon callback raises a compile error on `IRibbonControl` | Ensure the callback is in a standard module and the Office object library is available |
| Demo Ribbon button fails | The demo sheet is built on first use. If it still fails, confirm `M_DP_DEMO.bas` and `M_DEMO_BUILDER.bas` are imported |

---

## ⚠️ Known limitations

<p align="left">
  <img alt="Status" src="https://img.shields.io/badge/Read_before_deploying-important-orange">
</p>

One known defect is open against the next release. It is safe to work with once
you know it exists.

### Single-owner runtime

The component registers process-wide Excel surfaces — the keyboard shortcut, the
cell context-menu entry, `Application` events, the live-clock timer, worksheet
icons, and registry settings — without an ownership model. Two copies in one
Excel session will interfere: either can remove the other's registrations and
delete the other's worksheet icons.

Do not run the embedded source and the `.xlam` in the same Excel session, and do
not embed the picker into multiple workbooks that will be open at once.
Tracked in [#14](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/14).

### The keyboard shortcut is application-wide

`Ctrl + Shift + D` is registered through `Application.OnKey`, which holds one
assignment per key for the whole Excel session. Enabling the DatePicker shortcut
can therefore replace a binding another workbook or add-in already made.

Excel provides no way to read the existing assignment, so the DatePicker cannot
detect the conflict, cannot record what it replaced, and cannot put it back.
Disabling the shortcut restores Excel's **default** handling for the key, not the
displaced assignment.

The shortcut is only ever registered when you enable it. Disabling the right-click
entry and the in-grid icon does not turn it on — taking a session-wide key binding
is not a decision the component makes on your behalf.

If `Ctrl + Shift + D` matters to something else in your session, leave the
DatePicker shortcut disabled and use `DP_Show`, the Ribbon button, or your own
button instead.

### Operating notes

- Do not call DatePicker entry points inside a macro that depends on
  `Application.EnableEvents = False`. `DP_RepairRuntime` is the only routine
  that deliberately re-enables events.
- Formula cells are preserved. A cell holding a formula is skipped and reported,
  including one that evaluates to a date: replacing a displayed date does not
  imply deleting the formula that produced it. Existing literal values are still
  overwritten. Cells belonging to an array formula cannot be written at all and
  are reported as failures.
- The written and skipped counts are returned, not inferred. `DP_FillTableColumn`
  compares what it predicted against what it targeted, so a fill that preserves
  three formulas reports 247 targeted and 244 written rather than implying every
  cell changed.
- Test the exact Ribbon package, not only the source callbacks.
- Test both 32-bit and 64-bit Office where both are supported.
- Keep an explicit teardown path available — `DP_Stop` or `DP_RepairRuntime`.
- Accessibility, high-DPI, and high-contrast behaviour are not yet tested or
  documented.

---

## 📚 Wiki

<p align="left">
  <img alt="Docs" src="https://img.shields.io/badge/Guidance-Extended_notes-blue">
</p>

The project wiki was written for `v1.1.0` and is broadly accurate for it, but it
has not yet been verified against `v1.1.1`. Two areas are known to be out of
date — the `Application.EnableEvents` behaviour of `M_Picker_EnsureManager`, and
some file paths. Each affected page carries a notice.

[VBA-DATETIMEPICKER Wiki](https://github.com/danielep71/VBA-DATETIMEPICKER/wiki)

Treat `README.md` and the tagged source as authoritative until the wiki rewrite
is complete.

---

## 🔧 Coding style

<p align="left">
  <img alt="Style" src="https://img.shields.io/badge/Style-Structured_VBA-217346">
</p>

This repository follows a structured VBA style.

Typical routine headers include:

- PURPOSE
- WHY THIS EXISTS
- INPUTS
- RETURNS
- BEHAVIOR
- ERROR POLICY
- DEPENDENCIES
- NOTES
- UPDATED

General conventions:

- `Option Explicit`
- comments above executable lines
- inline comments only for declarations
- no comments inside `With ... End With`
- explicit cleanup / fail-safe logic where appropriate
- stable external behavior unless a change is deliberate and documented
- production-quality code intended for real Excel environments

---

## 🙏 Acknowledgements & prior art

<p align="left">
  <img alt="Inspired by" src="https://img.shields.io/badge/Inspired_by-Sam_Rad's_DatePicker-blueviolet">
  <img alt="Implementation" src="https://img.shields.io/badge/Implementation-Independent_%2F_from_scratch-brightgreen">
</p>

`VBA-DATETIMEPICKER` was inspired by [Sam Radakovitz's Excel Date Picker](http://samradapps.com/datepicker),
whose modeless, in-grid approach I admired and used as a reference for the experience I wanted to build.

This is an **independent, open-source reimplementation** written from scratch, with different goals:

- **Open source (MIT)** — full VBA source, published and importable, rather than a compiled add-in
- **Embeddable** — import the source so the picker travels inside your workbook; no per-machine install
- **Developer-oriented** — a documented `DP_*` public API, a regression test harness, and a full wiki

Credit also to the wider VBA community's date-picker tradition for showing what's achievable in pure VBA.

---

## 📄 License

<p align="left">
  <img alt="License" src="https://img.shields.io/badge/License-MIT-green">
</p>

---

## 👤 Author

<p align="left">
  <img alt="Maintainer" src="https://img.shields.io/badge/Maintainer-Daniele_Penza-orange">
</p>

---

