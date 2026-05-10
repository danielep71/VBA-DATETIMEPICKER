# VBA-DATETIMEPICKER

<p align="center">
  <b>A modern, reusable, worksheet-friendly Date & DateTime Picker for Excel VBA</b><br>
  Clean UI • Smart cell detection • In-grid activation • Enterprise-friendly architecture
</p>

<p align="center">
  <img alt="Platform" src="https://img.shields.io/badge/platform-Excel_VBA-217346">
  <img alt="Office" src="https://img.shields.io/badge/office-32%2F64--bit-blue">
  <img alt="License" src="https://img.shields.io/badge/license-MIT-green">
  <img alt="Status" src="https://img.shields.io/badge/status-active_development-orange">
</p>

---

<img width="1280" height="640" alt="vba-datetimepicker-social-preview" src="https://github.com/user-attachments/assets/b7e70418-8a08-4b60-9e6c-602d322425be" />


---

## ⭐ Why this repo exists

Excel still powers an enormous number of business-critical workflows, but native date-entry UX is often poor.

This repository exists to provide a **clean, modern, and reusable VBA-based solution** for a very common problem:  
**making date input in Excel feel much better without sacrificing control, portability, or maintainability**.

---

## ✨ Overview

**VBA-DATETIMEPICKER** is a modern, modular, and worksheet-friendly **Date / DateTime Picker for Excel VBA**.

It is designed for developers who want a cleaner and more intuitive way to handle date input in Excel, without forcing users to type dates manually or interact with clunky modal forms.

Built as part of a broader **Excel VBA Runtime Framework**, this project focuses on:

- 📅 better date-entry UX in Excel
- 🧩 modular, maintainable VBA architecture
- ⚙️ reusable infrastructure for enterprise workbooks and add-ins
- 🧪 demo-ready and regression-friendly implementation patterns

---

## 🚀 Key Features

- 📌 **Modeless DatePicker UserForm**  
  Keeps Excel usable while the picker is open

- 🖱️ **In-grid activation**  
  Shows contextual UI directly where the user is working

- 🧠 **Smart date-cell detection**  
  Can react to date-like values and date-formatted cells

- 🔁 **Multi-cell write-back support**  
  Suitable for user workflows involving repeated date entry

- ⚙️ **Settings support**  
  Configurable behavior such as first day of week and display preferences

- ⌨️ **Mouse + keyboard navigation**  
  Designed for practical day-to-day usage

- 🧱 **Manager / UI separation**  
  Core orchestration is separated from form rendering logic

- 🧪 **Demo workbook structure**  
  Easy to showcase, test, and evolve

---

## 📸 Screenshots

### Main picker
<img  src="assets/datepicker-main.png" />
```

### Settings panel

<img  src="assets/datepicker-settings.png" />
```

### In-grid activation
```markdown
![In-Grid Activation](assets/datepicker-ingrid.png)
```

If you already have images ready, you can simply replace the placeholder paths above with your actual filenames.

---

## 🏗️ Architecture

The project is organized around a clean separation of responsibilities.

### `cDatePickerManager`
Central controller responsible for:

- Excel event hooks
- UI lifecycle orchestration
- stale UI cleanup
- selection-driven show / hide behavior
- reentrancy protection
- workbook / worksheet context handling

### `UF_DatePicker`
Dedicated UI layer responsible for:

- calendar rendering
- month navigation
- day-grid interaction
- settings overlay
- user-facing controls and visual state

### Supporting modules / helpers
Depending on your implementation version, the project may also include:

- shared constants and configuration
- WinAPI helpers
- formatting helpers
- demo builder routines
- regression / test helpers

---

## 🧠 Design Principles

This repo follows a few core principles:

- **Excel-first UX**  
  The picker should feel natural inside worksheet workflows

- **Reusable VBA engineering**  
  The solution should be portable across workbooks and projects

- **Separation of concerns**  
  UI, orchestration, and infrastructure should remain clearly separated

- **Safe event handling**  
  Reentrancy guards and cleanup logic matter in real Excel environments

- **Production-minded code quality**  
  This is not just a demo widget; it is intended as a robust reusable component

---

## 📦 Repository Structure

```text
VBA-DATETIMEPICKER/
├─ src/
│  ├─ classes/
│  │  └─ cDatePickerManager.cls
│  ├─ forms/
│  │  └─ UF_DatePicker.frm
│  ├─ modules/
│  │  └─ supporting modules (.bas)
│  └─ demo/
│     └─ VBA_DateTimePicker_Demo.xlsm
├─ assets/
│  └─ screenshots / social preview / branding
├─ README.md
├─ LICENSE
├─ CHANGELOG.md
└─ CONTRIBUTING.md
```

> Adjust the structure above to match your actual repository layout

---

## 🛠️ Installation

### Option 1 — Import into an existing VBA project

1. Open the target workbook in Excel
2. Open the VBA Editor
3. Import the project files:
   - `.bas` modules
   - `.cls` class modules
   - `.frm` UserForm and companion files
4. Ensure all dependencies are included
5. Initialize the manager at workbook startup

Example:

```vba
Public DP As cDatePickerManager

Private Sub Workbook_Open()
    Set DP = New cDatePickerManager
    DP.Initialize Application
End Sub
```

### Option 2 — Use the demo workbook

Open the included demo workbook to:

- explore the component behavior
- validate the interaction model
- test settings
- use it as a starting point for integration

---

## ⚙️ Configuration

Settings are stored using VBA registry persistence under:

```text
HKCU\Software\VB and VBA Program Settings\VBA_DATETIMEPICKER
```

Typical examples include:

- first day of week
- display preferences
- behavior toggles
- integration-related options

---

## 🧩 Typical Use Cases

This component is useful when you want to improve date input in:

- financial models
- treasury tools
- risk-management workbooks
- internal operational tools
- user-facing Excel templates
- demo / showcase workbooks
- reusable VBA frameworks

---

## 🧪 Demo & Testing

A good demo workbook should typically include:

- sample input cells
- date-formatted fields
- multi-cell scenarios
- a settings section
- regression-style usage checks

Suggested validation areas:

- selection change behavior
- workbook / worksheet activation handling
- cleanup after UI close
- month navigation
- multi-cell write-back
- settings persistence

---

## 🔧 Coding Style

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

---

## 🧭 Roadmap

Planned or possible future enhancements include:

- 🌍 localization support
- 📆 business-day / holiday-calendar integration
- 🪶 compact display modes
- ⚡ optional caching
- 🧠 richer detection heuristics
- 🧪 lightweight regression harness
- 🔌 stronger framework integration with related repos

---

## 🌐 Part of a Broader Framework

**VBA-DATETIMEPICKER** is intended to sit naturally alongside other reusable Excel VBA components, such as:

- performance / execution managers
- Excel UI controllers
- demo builders
- progress-bar components
- calendar and date infrastructure

If you are building a broader VBA engineering ecosystem, this repo can serve as the **date-input / calendar UX component** of that toolkit.

---

## 🤝 Contributing

Contributions are welcome.

Please aim to:

- preserve architectural clarity
- keep external behavior stable unless change is justified
- follow the project coding style
- document meaningful changes
- use clear commit messages

Conventional Commit examples:

- `feat(ui): add compact settings layout`
- `fix(manager): clean up stale in-grid icon on sheet change`
- `docs(readme): expand installation instructions`

---

## 📄 License

This project is released under the **MIT License**.

---


