---
name: 🐞 Bug report
about: Report reproducible DatePicker behavior, write-back, runtime, UI, compatibility, or recovery defects
title: "[Bug]: "
labels: bug
---

<!--
  VBA-DATETIMEPICKER bug report

  Thank you for taking the time to report a defect.

  Fill the core sections first. The subsystem sections near the bottom are
  optional: expand only the ones relevant to the problem and delete the rest.

  Precision helps much more than volume. Record only states and tests you
  actually observed.

  SECURITY:
  Do NOT use this public template for a suspected vulnerability, credential
  exposure, malicious release artifact, or exploitable trust-boundary issue.
  Follow SECURITY.md and report privately.
-->

## 🐞 Summary

<!--
Describe the defect in one or two paragraphs.

Good:
"Selecting a formula cell that displays a date and choosing another date replaces
the formula even though OverwriteFormulas was not requested."

Less useful:
"DatePicker broken."
-->

## 💥 Practical impact

<!--
What does this prevent, corrupt, misreport, or make difficult?

Examples:
- wrong cell(s) changed
- formula lost
- picker cannot start
- another provider is disrupted
- Excel events changed unexpectedly
- UI cannot be recovered
- intermittent test evidence
-->

```text
Impact:
Severity from your perspective: low / medium / high / blocking
```

> [!IMPORTANT]
> If the issue may expose confidential data, execute unintended code, disclose a
> credential, tamper with an official release asset, or create a concrete
> security/integrity exploit, do **not** continue in public.
>
> Follow the private process in `SECURITY.md`.

---

## 🔖 Exact version and source state

Do not write only:

```text
latest
current
main
```

Identify the exact source you tested.

```text
Release tag:         <e.g. v1.2.0 / N/A>
Commit SHA:          <full 40-character SHA if known>
Branch:              <main / release/v1.2.0 / other / N/A>
Source obtained from:<official repository / GitHub Release / other>
```

### Deployment

- [ ] Embedded source in workbook
- [ ] `.xlam` add-in
- [ ] Official demo `.xlsm`
- [ ] Development / test workbook
- [ ] Other

```text
Workbook/add-in filename:
```

---

## 🎚️ Affected area

Check all that apply.

**Picker and write-back**

- [ ] 📅 Calendar selection
- [ ] 🗓️ `DP_Today`
- [ ] 🕒 `DP_Now`
- [ ] 📊 `DP_FillTableColumn`
- [ ] 🎯 Selected-cell / range scope
- [ ] 🧮 Formula preservation / overwrite
- [ ] 🔒 Locked / protected cells
- [ ] 🧾 `DP_WriteResult`

**Runtime / Excel integration**

- [ ] 🚀 `DP_Start`
- [ ] 🛑 `DP_Stop`
- [ ] 🛠️ `DP_RepairRuntime`
- [ ] 🪪 Provider lease / duplicate provider
- [ ] ⚙️ Settings / namespace
- [ ] 🧠 Application event manager
- [ ] ⌨️ `Ctrl + Shift + D` / `Application.OnKey`
- [ ] 🧷 Right-click menu
- [ ] 📌 In-grid icon / worksheet Shape
- [ ] ⏱️ Live clock / `Application.OnTime`
- [ ] 🎛️ RibbonX

**User interface / native layer**

- [ ] 🖼️ UserForm rendering / controls
- [ ] ⌨️ Picker keyboard navigation
- [ ] 🪟 Borderless styling / WinAPI
- [ ] 📍 Form positioning / monitor behavior
- [ ] 🧾 `DP_WindowStyleResult`

**Project / validation**

- [ ] 🧪 Regression harness
- [ ] 🖼️ Demo workbook/source
- [ ] 📦 `.xlam` / `.xlsm` packaging
- [ ] 📖 Documentation
- [ ] Other

---

## 🔁 Minimal steps to reproduce

Please reduce the problem to the smallest sequence that still fails.

1.
2.
3.

```text
Reproducibility: always / often / intermittent / once only
Approximate frequency if intermittent:
First observed after:
```

If the problem appears only after a particular lifecycle event, include it:

- [ ] Workbook open
- [ ] Workbook close/reopen
- [ ] Excel restart
- [ ] VBA project reset
- [ ] Switching workbooks
- [ ] Deleting/recreating a worksheet
- [ ] Protecting/unprotecting a sheet
- [ ] Starting/stopping another DatePicker provider
- [ ] Changing DatePicker settings
- [ ] Other

---

## ✅ Expected behavior

<!-- What should have happened? -->

## ❌ Observed behavior

<!-- What actually happened? Include exact wording of errors/messages where possible. -->

```text
Runtime error number:
Runtime error description:
Message shown:
Immediate Window output:
```

---

## 🔢 Smallest exact call

If the issue can be reproduced from VBA, paste the smallest call.

```vb
Option Explicit

Public Sub ReproduceDatePickerIssue()

    ' Replace with the minimal call that demonstrates the defect.
    DP_Show

End Sub
```

For write-back defects, include the actual call and result where possible:

```vb
Dim Result As DP_WriteResult

Result = DP_FillTableColumn( _
            ValueToWrite:=VBA.Date, _
            ConfirmFill:=False)

Debug.Print "Attempted: "; Result.AttemptedCount
Debug.Print "Written:   "; Result.WrittenCount
Debug.Print "Locked:    "; Result.LockedSkippedCount
Debug.Print "Formula:   "; Result.FormulaSkippedCount
Debug.Print "Failed:    "; Result.FailedCount
```

Do not rewrite the call into what you think the project intended. Paste the form
that actually reproduced the issue.

---

# 🖥️ Environment

```text
Excel product:          <Microsoft 365 / Excel 2024 / Excel 2021 / other>
Excel version:
Excel build:
Office bitness:         32-bit / 64-bit
Windows version:
Workbook type:          .xlsm / .xlsb / .xlam / other
Display scaling:        100% / 125% / 150% / 200% / other
Monitor count:
```

### Excel session

```text
Open workbooks:
Other add-ins active:
DatePicker deployment:  embedded / .xlam / demo
Other DatePicker copy:  none / yes
Other provider version:
Other provider type:    embedded / .xlam / unknown
```

> [!IMPORTANT]
> For provider/registration problems, the **version of every DatePicker copy in
> the same Excel process matters**.
>
> `v1.2.0 + v1.2.0` participates in the provider-lease protocol.
>
> `v1.2.0 + pre-v1.2.0` is a mixed-version session and is not protected in the
> same way.

---

# 🧪 Regression / compile evidence

Run what is practical in a controlled workbook.

```text
Debug → Compile VBAProject          →
TST_DP_RunAll                       →
TST_DP_RunAll_WithUISmoke           →
TST_DP_ReportEnvironment            →
```

### Harness summary

Paste the actual line if available:

```text
State=; Run=; Passed=; Failed=; CleanupFailures=
```

Harness states are:

```text
PASS
FAIL
FAIL_CLEANUP
FAIL_DIRTY_START
INCOMPLETE_SKIPPED
```

> [!CAUTION]
> Only `PASS` is a passing run.
>
> `FAIL_DIRTY_START` means the harness did not trust the starting environment.
> It should not be interpreted as ordinary assertion evidence for the current
> code.
>
> The assertion count itself is not fixed; report the count you actually got.

If you imported the harness manually into a clean workbook, note that it also
depends on:

```text
demo/M_DEMO_BUILDER.bas
```

---

# 🆘 Recovery result

Please state what restored the session, if anything.

```text
DP_Close:                         restored / did not restore / not tried / N/A
DP_Stop:                          restored / did not restore / refused / not tried / N/A
DP_RepairRuntime:                 restored / did not restore / refused / not tried / N/A
Closing/reopening workbook:       restored / did not restore / not tried / N/A
Restarting all of Excel:          restored / did not restore / not tried / N/A
Other action:
```

### Provider lease recovery

Did you run:

```vb
DP_ForceReleaseProviderLease
```

```text
Yes / No
```

If yes, explain why you knew no other live DatePicker provider owned the session:

```text

```

> [!WARNING]
> `DP_ForceReleaseProviderLease` is an operator-only recovery command.
>
> Please do not run it merely to make a startup refusal disappear. If provider
> ownership is uncertain, restarting Excel is the safer diagnostic/recovery step.

A defect that leaves Excel or the DatePicker difficult to recover is usually more
important than a cosmetic defect.

---

# 🔬 Optional subsystem details

<!--
Expand only the sections relevant to the bug. Delete the rest.
-->

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

Expected v1.2.0 normal behavior includes:

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
```

Completed result should conceptually balance as:

```text
AttemptedCount =
    WrittenCount +
    LockedSkippedCount +
    FormulaSkippedCount +
    FailedCount
```

If it does not, paste the exact values rather than correcting them manually.

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

- [ ] Second `v1.2.0` provider was correctly refused
- [ ] Second `v1.2.0` provider was not refused
- [ ] Refused provider altered owner state
- [ ] Refused provider removed owner state during teardown
- [ ] Stale lease remained after VBA reset
- [ ] Lease remained after **full Excel process restart**
- [ ] Mixed-version (`v1.2.0` + older) interference
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

---

## 🔀 Scope and interaction

Check anything that appears necessary for reproduction:

- [ ] One workbook only
- [ ] Multiple workbooks in one Excel process
- [ ] More than one Excel process
- [ ] Embedded source only
- [ ] `.xlam` only
- [ ] Embedded + `.xlam`
- [ ] Two `v1.2.0` providers
- [ ] `v1.2.0` + pre-v1.2.0 provider
- [ ] Another third-party add-in
- [ ] Protected worksheet
- [ ] Excel Table
- [ ] Formula cells
- [ ] Multi-area selection
- [ ] VBA project reset
- [ ] Excel process restart
- [ ] Multiple monitors
- [ ] High DPI/scaling
- [ ] RibbonX
- [ ] Other

---

## 📎 Screenshots, logs, or minimal reproduction

Attach anything that materially helps:

- sanitized screenshots;
- short screen recording/GIF;
- Immediate Window output;
- minimal `.bas` reproduction;
- regression summary;
- environment output;
- sanitized minimal workbook if genuinely necessary.

> [!WARNING]
> Do not attach workbooks or screenshots containing:
>
> - client/production data;
> - credentials or tokens;
> - personal data;
> - proprietary VBA you cannot publish;
> - production connection strings;
> - sensitive file-system paths;
> - confidential external links.
>
> Excel workbooks can also contain hidden sheets, names, connections, cached
> values, VBA, document properties, and other metadata that are not obvious from
> the visible worksheet.

---

## 📝 Anything else

<!--
Add hypotheses, timing observations, known workarounds, or related issues.

Clearly label a suspected cause as a hypothesis rather than a confirmed cause.
-->

```text
Known workaround:
Suspected cause (if any):
Related issue/PR:
Other context:
```

---

<!--
Maintainer triage reminder:

Separate:
- reproducible source behavior
- deployment/package behavior
- mixed-version/provider interaction
- Excel/platform limitation
- regression-harness evidence quality
- security-sensitive reports that should be moved to private handling
-->
