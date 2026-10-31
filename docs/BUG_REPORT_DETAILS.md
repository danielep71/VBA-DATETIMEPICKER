# Bug report: subsystem details

Optional detail checklists for the **Bug report** issue form. Copy the sections
that apply to your report into its **Subsystem details** field and fill them in.
Leave out sections that do not apply. Record only states and tests you actually
observed.

Do not use a public issue for a suspected vulnerability, credential exposure,
malicious release artifact or exploitable trust-boundary issue: follow
[SECURITY.md](../SECURITY.md) and report privately.

## 🎯 Write-back / target scope

<details>
<summary>Expand for wrong cell/range/Table scope, partial writes, Today/Now, or DP_FillTableColumn</summary>

<br>

```text
Selected address:
Resolved/expected target:
Worksheet:
Inside Excel Table:        yes / no
Table name:
Table column:
Selection areas:
ConfirmFill used:
Expected attempted cells:
Observed attempted cells:
Observed written cells:
```

### Cell state before the write

- [ ] Empty
- [ ] Literal date/time
- [ ] Other literal
- [ ] Formula returning date/time
- [ ] Other formula
- [ ] Array formula / legacy array formula
- [ ] Locked cell
- [ ] Protected sheet
- [ ] Merged cell
- [ ] Other

```text
Formula/value before:
Number format:
Formula/value after:
```

Expected normal behavior includes:

```text
selected Table cell
    → selected cell only

explicit DP_FillTableColumn
    → whole resolved Table data column
```

If the reported defect contradicts one of those, include the exact call that
selected the path.

</details>

---

## 🧮 Formula preservation / `DP_WriteResult`

<details>
<summary>Expand when a formula changed unexpectedly, a formula should have changed, or result counts are wrong</summary>

<br>

```text
Formula before:
Displayed value before:
OverwriteFormulas passed:  True / False / omitted
Formula after:
Displayed value after:
```

Paste the result:

```text
AttemptedCount:
WrittenCount:
LockedSkippedCount:
LockedSkippedAddresses:
FormulaSkippedCount:
FormulaSkippedAddresses:
FailedCount:
FailedAddresses:
ResolvedTargetAddress:
ExpandedToTableColumn:
TableName:
ColumnName:
AreasCount:
EventsDisabledByCaller:
TechnicalFailureOccurred:
TechnicalFailureStep:
TechnicalFailureNumber:
TechnicalFailureDescription:
```

Completed result should conceptually balance as:

```text
AttemptedCount =
    WrittenCount +
    LockedSkippedCount +
    FormulaSkippedCount +
    FailedCount
```

A result with `TechnicalFailureOccurred = True` does not obey that identity; the
`TechnicalFailure` fields carry the cause. If any result does not balance, paste
the exact values rather than correcting them manually.

</details>

---

## ⚙️ Settings / namespace

<details>
<summary>Expand for settings not persisting, shared settings, wrong defaults, or namespace errors</summary>

<br>

```text
Namespace configured:
M_Settings_GetNamespace result:
When M_Settings_SetNamespace was called:
DP_Start already called first:     yes / no
Settings already loaded:           yes / no / unknown
Expected persisted value:
Observed persisted value:
```

Which scope was expected?

- [ ] Legacy/default `VBA_DATETIMEPICKER`
- [ ] Explicit namespace
- [ ] Unsure

Did the problem survive a complete Excel restart?

```text
Yes / No / Not tested
```

A settings namespace isolates **persisted preferences**. It does not enable two
DatePicker providers to run concurrently.

</details>

---

## 🪪 Provider lease / duplicate provider

<details>
<summary>Expand for startup refusal, teardown refusal, stale lease, or provider interference</summary>

<br>

```text
Provider A version:
Provider A deployment:        embedded / .xlam
Provider A started first:      yes / no / unknown

Provider B version:
Provider B deployment:        embedded / .xlam
Provider B started second:     yes / no / unknown

Exact DP_Start message/error:
DP_Stop result:
DP_RepairRuntime result:
VBA project reset occurred:    yes / no
Excel fully restarted after:   yes / no
```

Check what you observed:

- [ ] Second lease-participating provider was correctly refused
- [ ] Second lease-participating provider was not refused
- [ ] Refused provider altered owner state
- [ ] Refused provider removed owner state during teardown
- [ ] Stale lease remained after VBA reset
- [ ] Lease remained after **full Excel process restart**
- [ ] Lease-participating provider + pre-`v1.2.0` provider interference
- [ ] Other

The temporary lease should disappear when the Excel process actually exits.

If Excel still had a background process alive, note that.

</details>

---

## 🧠 `Application.EnableEvents` / manager behavior

<details>
<summary>Expand if events changed unexpectedly, manager startup failed, or selection handling stopped</summary>

<br>

```text
Application.EnableEvents before:
Call made:
Application.EnableEvents after:
DP_RepairRuntime involved:       yes / no
Workbook event involved:
Selection event involved:
```

Normal DatePicker entry points are expected to preserve caller event state.

`DP_RepairRuntime` is intentionally different: it is a repair operation that may
re-enable events.

If your business macro deliberately had events disabled, include the smallest
transaction that demonstrates the change.

</details>

---

## ⌨️ Keyboard shortcut / right-click / access paths

<details>
<summary>Expand for Ctrl+Shift+D, right-click registration, access-path coupling, or another macro binding</summary>

<br>

```text
ShowRightClick:
ShowGridIcon:
EnableKeyboardShortcut:
Shortcut worked before:
Shortcut works after:
Other known Ctrl+Shift+D owner:
Right-click entry count/behavior:
```

This is a valid configuration:

```text
ShowRightClick = False
ShowGridIcon = False
EnableKeyboardShortcut = False
```

The DatePicker should not force-enable the keyboard shortcut merely because the
other built-in access paths are disabled.

For `Application.OnKey` conflicts, note that Excel exposes no supported getter
for the previous assignment, so the exact predecessor may not be observable.

</details>

---

## 📌 In-grid icon / worksheet Shape

<details>
<summary>Expand for missing, duplicated, stale, misplaced, or protected-sheet grid icons</summary>

<br>

```text
Worksheet:
Selected cell:
Sheet protected:
DrawingObjects protected:
UserInterfaceOnly:
Icon visible before:
Icon visible after:
Sheet deleted/recreated:
Shape deleted manually:
```

Check any that apply:

- [ ] Icon missing on eligible cell
- [ ] Duplicate icons
- [ ] Icon remains after leaving eligible cell
- [ ] Icon remains after workbook/sheet lifecycle boundary
- [ ] Error after sheet deletion
- [ ] Error after shape deletion
- [ ] Wrong worksheet owns the icon
- [ ] Protection-specific behavior

If the issue followed deletion of a worksheet or Shape, say whether the same
DatePicker session had previously tracked that icon.

</details>

---

## ⏱️ Live clock / `Application.OnTime`

<details>
<summary>Expand for clock not updating, timer surviving close, repeated callbacks, or teardown problems</summary>

<br>

```text
Clock mode:
Picker open/closed:
DP_Close called:
DP_Stop called:
Excel restart required:
Repeated callbacks observed:
Exact error/output:
```

If the problem is intermittent, note whether it occurs:

- [ ] after opening/closing repeatedly
- [ ] after VBA reset
- [ ] after switching workbooks
- [ ] after stopping provider
- [ ] after Excel resumes from sleep
- [ ] other

</details>

---

## 🪟 WinAPI / borderless window

<details>
<summary>Expand for title-bar removal, positioning, rollback, RecoveryRequired, monitor, or native API behavior</summary>

<br>

```text
UseWinAPI:
Office bitness:
Display scaling:
Monitor topology:
Picker position expected:
Picker position observed:
Borderless expected:
Borderless observed:
```

If available, paste `DP_WindowStyleResult` values:

```text
Attempted:
Applied:
Committed:
RolledBack:
RecoveryRequired:
FailedStep:
LastApiError:
```

Check any that apply:

- [ ] Safe non-attempt unexpectedly reported as failure
- [ ] Style committed but UI did not finish applying
- [ ] Rollback succeeded
- [ ] Rollback failed
- [ ] `RecoveryRequired=True`
- [ ] Native result says success but visible/native state disagrees
- [ ] Wrong monitor / position
- [ ] Multi-window-specific
- [ ] 32/64-bit-specific

</details>

---

## 🖼️ UserForm / keyboard navigation

<details>
<summary>Expand for visual layout, focus, key handling, overlays, month/year navigation, or scaling</summary>

<br>

```text
Form mode:               normal / compact
Clock:                   static / live
First day of week:
Use local names:
Display scaling:
Keyboard key/sequence:
Focused control before:
Observed focus after:
```

Include a screenshot or short recording if the bug is visual.

Please crop or sanitize workbook/client information.

</details>

---

## 🎛️ RibbonX / packaging

<details>
<summary>Expand if a Ribbon control is missing/dead or behavior differs between source and .xlam/.xlsm package</summary>

<br>

```text
Package:                 .xlsm / .xlam
Ribbon XML source:
Callback name:
Button/control:
Source callback works when called directly: yes / no
Packaged Ribbon works:                  yes / no
```

If the callback works directly but the control does not, the defect may be in
Office package/RibbonX wiring rather than DatePicker runtime logic.

State whether the package is an official GitHub Release asset.

</details>

---

## 🧪 Regression harness

<details>
<summary>Expand for FAIL_DIRTY_START, cleanup failures, missing suites, TST_DP_SCRATCH, or setup errors</summary>

<br>

```text
Entry point:
State:
Run:
Passed:
Failed:
CleanupFailures:
Suites dispatched:
Suites completed:
TST_DP_SCRATCH existed before run:
TST_DP_SCRATCH exists after run:
TST_DP_RESULTS existed before run:
```

If setup failed:

```text
Exact setup step:
Error number:
Error description:
TST_DP_ReportEnvironment output:
```

For a `Worksheets.Add` failure, record:

```text
Worksheet count before:
Worksheet count after:
New worksheet observed despite error: yes / no
More than one candidate appeared:     yes / no
```

Do not repeatedly rerun a setup failure without inspecting the workbook; that can
destroy useful evidence about the first failure.

</details>

---

## 📦 Release asset / source provenance

<details>
<summary>Expand if the bug exists only in a downloaded .xlam/.xlsm or differs from imported source</summary>

<br>

```text
Release tag:
Asset filename:
Official GitHub Release:     yes / no
Source commit used for comparison:
Behavior reproduces from imported source: yes / no / not tested
Behavior reproduces only in binary asset: yes / no / unknown
```

If you calculated a checksum:

```text
SHA-256:
```

Do not assume that a source regression result proves a later-saved binary is the
same artifact. Package behavior and source behavior are useful separate facts.

</details>
