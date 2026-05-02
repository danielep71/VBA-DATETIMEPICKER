# VBA-DATETIMEPICKER

A modern, reusable, worksheet-friendly Date & DateTime Picker for Excel VBA.

Built as part of a broader **Excel VBA Runtime Framework**, this component provides a clean, extensible, and user-friendly way to handle date input directly within Excel workbooks.

---

## ✨ Features

- 📅 Modeless DatePicker UserForm (non-blocking UI)
- 🎯 In-grid activation with contextual icon
- 🧠 Smart detection of date-like cells
- 🔁 Multi-cell write-back support
- ⚙️ Configurable behavior (first day of week, display options)
- 🖱️ Mouse + keyboard navigation
- 🧩 Fully modular architecture (Manager + Form separation)
- 🧪 Demo workbook + regression-ready structure

---

## 🏗️ Architecture

The component is structured around a **central manager class**:

- `cDatePickerManager`
  - Handles all Excel event hooks
  - Controls UI lifecycle (show / hide / rebuild)
  - Ensures reentrancy safety and cleanup

- `UF_DatePicker`
  - Pure UI layer
  - Dynamic calendar grid (6x7)
  - Header navigation (month/year)
  - Settings overlay panel

---

## 📸 Preview

![DatePicker Preview](assets/preview.png)

---

## 🚀 Installation

1. Open your Excel VBA project
2. Import:
   - `.bas` modules
   - `.cls` classes
   - `.frm` UserForm
3. Initialize the manager:

```vba
Public DP As cDatePickerManager

Private Sub Workbook_Open()
    Set DP = New cDatePickerManager
    DP.Initialize Application
End Sub
