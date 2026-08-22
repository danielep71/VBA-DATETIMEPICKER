<!--
  Sections that do not apply can be deleted outright.
  A template filled with "Not applicable" fifteen times hides the parts that
  matter, so deleting is preferred over padding.

  The collapsed sections near the bottom are relevant only when the change
  touches that subsystem. Expand the ones that apply; delete the rest.
-->

## 📌 Summary

<!-- What changed, and why it was needed. One paragraph is usually enough. -->

## 🔗 Related issue

```text
Closes #
```

---

## 🏷️ Type of change

- [ ] 🐛 Functional or compatibility fix
- [ ] 🎯 Write-scope or data-safety fix
- [ ] 🔒 Application-state or ownership fix
- [ ] ✨ Backward-compatible feature
- [ ] ♻️ Internal refactor with no intended public behavior change
- [ ] 🧪 Regression-test change
- [ ] 🖼️ Demo change
- [ ] 📖 Documentation-only change
- [ ] 🧹 Repository or release maintenance
- [ ] 🔐 Security-related change

## 🎚️ Affected surface

**Public API**

- [ ] 🚀 Startup / repair — `DP_Start`, `DP_Stop`, `DP_RepairRuntime`
- [ ] 🖱️ Open / close — `DP_Show`, `DP_Close`, `DP_Preload`, `DP_Hide`
- [ ] 📅 Write-back — `DP_Today`, `DP_Now`, `M_Picker_SelectDate`
- [ ] ⚙️ Settings getters and setters

**Excel integration**

- [ ] ⌨️ Keyboard shortcut — `Application.OnKey`
- [ ] 🧷 Cell context menu
- [ ] 📌 In-grid icon
- [ ] ⏱️ Live-clock timer — `Application.OnTime`
- [ ] 🧠 Application event manager
- [ ] 🎛️ Ribbon callbacks

**Component**

- [ ] 🖼️ UserForm and runtime controls
- [ ] 🪟 WinAPI styling, positioning or drag
- [ ] 🧾 Registry settings
- [ ] 🧪 Test harness
- [ ] 📖 Documentation only

---

## 🔒 Public API and Semantic Versioning

```text
Public behavior changed:
Backward compatible:
Suggested release:        patch / minor / major
Migration required:
```

Confirm changes to names, signatures, parameter order and defaults, enum values,
settings semantics, and the recovery path. Write `No public behavior change`
where applicable.

> [!IMPORTANT]
> The `VBA_DATETIMEPICKER` registry application name and command-bar tag are
> stable legacy identifiers. Renaming either silently discards every user's
> saved settings, or orphans every context-menu entry a previous version
> installed.

---

## ✅ Testing performed

```text
Debug → Compile VBAProject          →
TST_DP_RunAll                       →
TST_DP_RunAll_WithUISmoke           →
Manual DP_Show / DP_Close           →
Manual DP_RepairRuntime             →
```

**Harness summary line**

```text
INFO | Harness | Summary |
```

> [!CAUTION]
> The summary line only prints when the harness completed. Its absence is
> itself the verdict — an aborted run is not a pass, whatever the assertions
> that did execute reported.

## 🖥️ Validation environment

```text
Excel product/version/build:
Office bitness:                     32-bit / 64-bit
Windows version:
Workbook type:                      .xlsm / .xlam
Display scaling:                    100% / 125% / 150% / 200%
Monitors:
Other add-ins loaded:
```

List only environments actually tested. A defect that only appears on the other
bitness is invisible to the person who wrote it.

---

## 📋 Always required

### 🧹 Source

- [ ] Current branch was confirmed before committing.
- [ ] The project compiled cleanly with all required components present.
- [ ] Changed components were exported to the correct repository paths.
- [ ] `UF_DatePicker.frx` accompanies any `.frm` change.
- [ ] `UPDATED` dates were bumped on every changed routine.
- [ ] No conflict markers or duplicate procedures remain.
- [ ] The textual diff contains only intended changes.
- [ ] No lock, backup, credential, client or production-data file is included.

### 🔒 Application state

- [ ] `Application.EnableEvents` is preserved for the caller. `DP_RepairRuntime`
      remains the only routine that deliberately enables it.
- [ ] `Application.OnKey` registrations are removable and do not assume
      exclusive ownership of the session.
- [ ] `Application.OnTime` is cancelled with the exact scheduled time and
      qualified procedure name it was created with.
- [ ] Context-menu controls are removed by tag, never by index.
- [ ] Worksheet shapes are purged only at documented boundaries.
- [ ] Registry writes happen on explicit save only.
- [ ] Teardown removes everything the change registers.

### 🎯 Write scope

- [ ] No write extends beyond the visible selection without explicit user
      intent.
- [ ] Effect on the table write-scope default is stated, or none. A selected
      table cell receives the date on its own; `DP_FillTableColumn` is the only
      route to the whole column.
- [ ] The three `WriteBack` scope paths still pass — omitted argument, explicit
      `NoTableGrow:=False`, and `DP_FillTableColumn`.
- [ ] Formulas and existing values are not overwritten as collateral.
- [ ] Partial writes are reportable rather than silent.
- [ ] `M_WriteBack_Apply` still restores caller event state on both the success
      and failure paths.

### 🧾 Diagnostics

- [ ] `Err.Number` and `Err.Description` are captured before any `On Error`
      statement.
- [ ] Every routine's declared `ERROR POLICY` matches what the code does.
- [ ] Anything reachable from an error handler cannot itself raise.
- [ ] No unsolicited production `MsgBox` was introduced.

### 📖 Documentation

- [ ] README · CHANGELOG · CONTRIBUTING · wiki · procedure banners, as affected
- [ ] `CHANGELOG.md` entry added
- [ ] No documentation change required

> [!TIP]
> Documentation belongs in **this** pull request, not a follow-up. A follow-up
> documentation commit is a commit that does not get written.

---

## 🖼️ Demo sheet

<details>
<summary>Expand when <code>demo/M_DP_DEMO.bas</code> or <code>demo/M_DEMO_BUILDER.bas</code> changed</summary>

<br>

The demo is built from source. No workbook is tracked, so the change is
reviewable as a diff — but the rendered result is not, and the release asset is
built from this code.

```text
Sections changed:
Rebuilt and inspected with:         DP_Demo_CreateDemoSheet
Tested from:                        embedded / .xlam / both
```

- [ ] The sheet was rebuilt with `DP_Demo_CreateDemoSheet` and visually checked.
- [ ] Rebuilding twice produces the same sheet — the builder is idempotent.
- [ ] Any new worksheet control has a working `OnAction` target.
- [ ] The demo sheet still starts `xlSheetVeryHidden`.
- [ ] No personal data, local paths or scratch content is written into the sheet.
- [ ] The demo still demonstrates current behaviour, not superseded behaviour.

</details>

## 🪟 WinAPI

<details>
<summary>Expand when the change touches window style, position or drag</summary>

<br>

```text
API used:
Owned style bits:
32-bit path:
64-bit path:
Err.LastDllError treatment:
Frame refresh:
Multi-window behavior:
```

- [ ] Declarations exist in both the `#If VBA7` and pre-`VBA7` branches.
- [ ] `SetLastError 0` precedes any call whose zero return is ambiguous.
- [ ] `Err.LastDllError` is read on the statement immediately after the call.
- [ ] `M_Window_GetUserFormHwnd` remains the single handle resolver — no new
      `FindWindow` call was added.
- [ ] `WS_CAPTION` is manipulated in `M_Window_RemoveTitleBar` and nowhere else.
- [ ] No code assigns `UserForm.Caption` as part of window resolution.
- [ ] Tested with more than one workbook window open.

</details>

## 🧪 Test harness

<details>
<summary>Expand when <code>test/M_cDP_Test.bas</code> changed</summary>

<br>

```text
Suites added or changed:
Registered in TST_DP_RunAllInternal:
Registered in TST_DP_RunSuiteSafe dispatcher:
Assertions added:
```

- [ ] A new suite is wired into **both** the runner and the dispatcher.
- [ ] `SuiteFail` captures `Err` before any `On Error` statement.
- [ ] The suite restores harness application state on every exit path.
- [ ] Scratch worksheet content is cleaned up.
- [ ] Existing suites still pass.

</details>

## 🎛️ Ribbon

<details>
<summary>Expand when Ribbon callbacks or <code>customUI14.xml</code> changed</summary>

<br>

```text
Callbacks changed:
customUI14.xml changed:
Image resources:
Tested in:                          .xlsm / .xlam / both
```

- [ ] Every `onAction` name matches a `Public Sub` in a standard module exactly.
- [ ] Callbacks stay thin and delegate to the public API.
- [ ] Any custom image referenced by the XML exists in the package.
- [ ] Tested in the actual package, not only through the source callbacks.

</details>

---

## 💬 Reviewer notes

<!--
  Trade-offs, known limitations, environments not tested, and follow-up work.
  A limitation stated here is a decision. The same limitation discovered later
  is a defect.
-->
