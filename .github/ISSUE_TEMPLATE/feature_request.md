---
name: ✨ Feature request
about: Propose a backward-compatible DatePicker capability, integration, safety, recovery, diagnostic, testing, or deployment enhancement
title: "[Feature]: "
labels: enhancement
---

<!--
  VBA-DATETIMEPICKER feature request

  Thank you for proposing an improvement.

  Start with the problem and desired behavior. Do not feel required to design
  the implementation before the use case is clear.

  The core sections are intentionally short. The subsystem sections near the
  bottom are optional and collapsed; expand only the ones that matter.

  If the proposal is really a reproducible defect in current documented
  behavior, use the Bug report template instead.

  If the proposal concerns a vulnerability, credential, malicious release
  artifact, or security-sensitive disclosure, follow SECURITY.md privately.
-->

## ✨ Problem / real use case

Describe the workflow, limitation, or recurring friction.

What are you trying to achieve that the current DatePicker does not support
cleanly?

<!--
Good:
"We distribute three independent workbooks that use different DatePicker
preferences. I need each workbook to retain its own settings without changing
runtime ownership."

Less useful:
"Add more settings."
-->

## 👤 Who benefits?

<!--
Examples:
- workbook author
- end user
- add-in deployment
- locked-down corporate user
- maintainer / tester
- release engineer
-->

```text
Primary user:
Typical deployment:
Frequency of need:
```

## 💡 Desired behavior

Describe the outcome **before** prescribing the implementation.

```text
Requested behavior:
Expected scope:
Expected default:
Expected failure/refusal behavior:
Expected result/diagnostics:
```

Where useful, show an illustrative call:

```vb
' Illustrative only — the final API may differ.
```

---

## ✅ Acceptance criteria

What would make the feature complete from the user's perspective?

1.
2.
3.

<!--
Prefer observable outcomes.

Good:
"With ShowRightClick=False, ShowGridIcon=False and
EnableKeyboardShortcut=False, the new Ribbon action still opens the picker."

Less useful:
"Add helper M_XYZ."
-->

## 🚫 Non-goals

What should this feature **not** try to solve?

```text

```

This helps keep a useful feature from silently becoming an architectural rewrite.

---

## 🔀 Current workaround / alternatives considered

Describe what you do today.

Check any that apply:

- [ ] Call existing `DP_...` procedures differently
- [ ] Use a workbook-specific wrapper macro
- [ ] Use RibbonX
- [ ] Use the right-click menu
- [ ] Use the in-grid icon
- [ ] Use `Ctrl + Shift + D`
- [ ] Use a settings namespace
- [ ] Use `DP_RepairRuntime`
- [ ] Restart Excel
- [ ] Use a separate `.xlam`
- [ ] Duplicate / modify DatePicker source
- [ ] Leave the behavior outside this repository
- [ ] No practical workaround

```text
Current workaround:
Why it is insufficient:
```

---

## 🎯 Affected area

Check all that appear relevant.

**Supported consumer surface**

- [ ] 🚀 Startup / lifecycle — `DP_Start`, `DP_Stop`
- [ ] 🖱️ Picker UI — `DP_Show`, `DP_Close`, `DP_Preload`, `DP_Hide`
- [ ] 📅 Direct write-back — `DP_Today`, `DP_Now`
- [ ] 📊 Explicit Table fill — `DP_FillTableColumn`
- [ ] 🛠️ Runtime repair — `DP_RepairRuntime`
- [ ] 🪪 Lease recovery — `DP_ForceReleaseProviderLease`
- [ ] ⚙️ Settings namespace — `M_Settings_SetNamespace`, `M_Settings_GetNamespace`
- [ ] 🧾 `DP_WriteResult`
- [ ] 🪟 `DP_WindowStyleResult`
- [ ] 🔢 Public enums / constants
- [ ] ➕ New supported API

**Excel integration**

- [ ] 🧠 Application event manager
- [ ] ⌨️ Keyboard shortcut / `Application.OnKey`
- [ ] 🧷 Right-click menu / CommandBars
- [ ] 📌 In-grid worksheet icon
- [ ] ⏱️ Live clock / `Application.OnTime`
- [ ] 🎛️ RibbonX
- [ ] 🗃️ Registry-backed settings
- [ ] 🪪 Runtime-provider lease

**Component / project**

- [ ] 🖼️ UserForm / runtime controls
- [ ] 🪟 WinAPI styling / positioning
- [ ] 🧪 Regression harness
- [ ] 🖼️ Demo
- [ ] 📦 `.xlam` / `.xlsm` packaging
- [ ] 🚦 Release process / provenance
- [ ] ⚙️ GitHub workflow / repository tooling
- [ ] 📖 Documentation only
- [ ] Other

---

# 🧩 Compatibility and API design

## 🔒 Public API / Semantic Versioning

```text
New supported member proposed:
Existing supported member changed:
Existing calls affected:
Backward compatible:             Yes / No / Unsure
Migration required:
Suggested release:               patch / minor / major / unsure
```

If proposing a new public member, explain why the behavior cannot be expressed
cleanly through the current supported surface.

Current principal supported surface includes:

```text
DP_Start
DP_Stop
DP_Show
DP_Close
DP_Preload
DP_Hide
DP_Today
DP_Now
DP_FillTableColumn
DP_RepairRuntime
DP_ForceReleaseProviderLease

M_Settings_SetNamespace
M_Settings_GetNamespace

DP_WriteResult
DP_WindowStyleResult

DP_WriteAction
DP_ClockMode
DP_SizeMode
```

> [!NOTE]
> A VBA procedure may need to be technically `Public` for RibbonX,
> `Application.Run`, Office callbacks, or cross-module testing without becoming
> supported consumer API.
>
> If the proposal adds a technically public internal seam, state that explicitly.

### Compatibility-sensitive behavior

If affected, explain changes to:

- procedure / function names;
- parameter order;
- optional defaults;
- enum numeric values;
- structured-result fields;
- safe Table write scope;
- formula-preservation default;
- application-state ownership;
- settings persistence;
- runtime-provider ownership;
- recovery semantics;
- stable `VBA_DATETIMEPICKER` identifiers.

---

## 🧭 Default behavior

Should the feature be:

- [ ] Enabled by default
- [ ] Disabled / opt-in by default
- [ ] Controlled per call
- [ ] Persisted as a setting
- [ ] Derived from existing state
- [ ] Automatic only in a narrowly defined condition
- [ ] Unsure

Explain why that default is safe and backward compatible.

```text

```

> [!IMPORTANT]
> In this project, a convenient automatic behavior can still be the wrong
> default if it expands write scope, changes caller-owned Excel state, consumes
> an application-wide resource, or destroys formulas.

---

# 🔬 Optional design sections

<!--
Expand the sections that apply. Delete the rest.
-->

## 🎯 Write-back / target scope

<details>
<summary>Expand when the feature writes worksheet data, changes target resolution, or adds Table behavior</summary>

<br>

Describe the intended target:

- [ ] Active cell
- [ ] Current selection
- [ ] Multi-area selection
- [ ] Resolved Table data cell
- [ ] Entire Table data column
- [ ] Specified `Range`
- [ ] Other

```text
How target is selected:
Can target expand beyond visible selection:
What constitutes explicit user/caller intent:
Preview/confirmation required:
Behavior if target changes before write:
```

Current v1.2.0 safety baseline:

```text
normal calendar / Today / Now
    → selected/resolved cell scope

DP_FillTableColumn
    → explicit whole Table data-column scope
```

Explain why the proposed behavior should preserve or change that distinction.

### Partial write

Should the operation:

- [ ] Be all-or-nothing
- [ ] Permit partial success and report it
- [ ] Skip protected/locked cells
- [ ] Skip formulas
- [ ] Fail on any non-writable cell
- [ ] Other

If partial success is allowed, describe the result contract.

</details>

---

## 🧮 Formula behavior

<details>
<summary>Expand when formula cells can be targeted, displayed dates are used, or overwrite policy changes</summary>

<br>

Current default:

```text
literal value        → writable
ordinary formula     → preserved
date-returning formula → preserved
array formula        → non-overridable failure
```

Proposed behavior:

```text
Ordinary formula:
Date-returning formula:
Array formula:
Explicit overwrite:
Persisted preference:
```

Questions to address:

- should a formula remain eligible to initialize the picker?
- should it remain protected from write-back?
- should destructive overwrite require a per-call opt-in?
- should the feature add a new `DP_WriteResult` classification?
- how should formula skips appear in diagnostics?

> [!CAUTION]
> A cell displaying a date is not necessarily a literal date cell.

</details>

---

## 🧾 Result / diagnostic contract

<details>
<summary>Expand when the feature can partially succeed, refuse safely, or needs new diagnostics</summary>

<br>

Should it:

- [ ] Remain fire-and-forget
- [ ] Return a Boolean
- [ ] Return / extend `DP_WriteResult`
- [ ] Return / extend `DP_WindowStyleResult`
- [ ] Introduce a new structured result type
- [ ] Log to the Immediate Window
- [ ] Produce a user-facing message
- [ ] Introduce machine-readable status/failure categories
- [ ] Preserve the current contract with no new diagnostics

Describe the desired contract:

```text
Attempted work:
Successful work:
Skipped work:
Failed work:
Refusal:
Recovery required:
Address/object identity included:
```

### If proposing a new result type

Explain why an existing result type does not fit the domain.

A worksheet write result and a native-window transaction are deliberately
different result domains today.

</details>

---

## ⚙️ Settings / persistence / namespace

<details>
<summary>Expand when the feature introduces a setting, changes persistence, or depends on deployment identity</summary>

<br>

Should the feature be:

- [ ] Session-only
- [ ] Persisted
- [ ] Per Windows user
- [ ] Per DatePicker namespace
- [ ] Per workbook
- [ ] Per add-in deployment
- [ ] Other

```text
Proposed setting:
Default:
Persistence lifetime:
Migration behavior:
```

Current persistence baseline:

```text
legacy/default base: VBA_DATETIMEPICKER
optional namespace:  VBA_DATETIMEPICKER__<stable namespace>
```

Current namespace rules:

- namespace is configured before first settings load;
- namespace locks after load;
- no automatic migration from legacy scope;
- namespace should represent stable deployment identity;
- namespace is not runtime-provider identity.

If proposing workbook-specific persistence, explain whether it should be:

```text
explicit namespace chosen by host
```

or:

```text
automatic workbook-derived identity
```

and how rename / Save As / file move should behave.

### Sensitive data

The DatePicker settings store must not become a credential or secret store.

If the feature requires secrets, explain why this repository is the right place
for them at all.

</details>

---

## 🪪 Runtime provider ownership / multi-provider behavior

<details>
<summary>Expand when the feature affects startup, teardown, shared Excel resources, provider coexistence, or the lease</summary>

<br>

Current v1.2.0 model:

```text
one participating current-version provider per Excel process
```

Lease:

```text
__VBA_DATETIMEPICKER_RUNTIME_PROVIDER_LEASE__
```

What should change?

- [ ] No ownership change
- [ ] Better duplicate-provider diagnostics
- [ ] Stale-owner recovery improvement
- [ ] Multiple providers should coexist
- [ ] Ownership should move between providers
- [ ] Resource ownership should become per workbook
- [ ] Other

Describe:

```text
Provider identity:
Resource identity:
Acquisition:
Conflict behavior:
Teardown ownership:
Dead/stale owner detection:
Recovery:
```

### Application-wide resources affected

- [ ] `Application.OnKey`
- [ ] Application events
- [ ] `Application.OnTime`
- [ ] CommandBars/context menu
- [ ] Worksheet Shape naming
- [ ] Other

> [!IMPORTANT]
> A settings namespace does not solve runtime ownership.
>
> Persistence identity and process-level ownership have different lifetimes.

### Mixed-version behavior

How should the feature behave with:

```text
new version + current v1.2.0
new version + pre-v1.2.0
```

A released old provider cannot be retrofitted with a protocol it never had.

</details>

---

## 🧠 Application state / event ownership

<details>
<summary>Expand when Application.EnableEvents, manager creation, selection events, or shared Excel state changes</summary>

<br>

Current baseline:

```text
normal DatePicker entry points
    → preserve caller Application.EnableEvents

DP_RepairRuntime
    → explicit recovery path that may re-enable events
```

Proposed behavior:

```text
Caller state read:
Caller state changed:
State restored:
Failure path:
Repair behavior:
```

Questions:

- does the feature require events to be enabled?
- can it operate while the caller deliberately has events disabled?
- should inability to operate be a refusal or an automatic repair?
- which component owns the repair decision?
- can the feature create event recursion/re-entrancy?

Avoid hidden repair behavior in ordinary entry points unless the use case truly
requires a contract change.

</details>

---

## ⌨️ Keyboard / right-click / grid-icon / Ribbon access

<details>
<summary>Expand when the feature adds or changes a way to open or invoke the DatePicker</summary>

<br>

Affected access paths:

- [ ] Right-click
- [ ] In-grid icon
- [ ] `Ctrl + Shift + D`
- [ ] RibbonX
- [ ] Caller-provided button/macro
- [ ] New access path

Current valid configuration:

```text
ShowRightClick = False
ShowGridIcon = False
EnableKeyboardShortcut = False
```

Zero built-in interactive access paths are permitted.

Describe the new access path:

```text
Registration lifetime:
Owner:
Persistence:
Conflict behavior:
Teardown:
Callback identity:
```

### If keyboard-based

Excel's `Application.OnKey` exposes no getter for the predecessor binding.

Explain:

- key combination;
- collision behavior;
- removal behavior;
- why the chosen key is appropriate;
- whether Excel default or a third-party binding is expected after teardown.

Do not design around restoring an unknown predecessor unless new platform
capability makes it observable.

</details>

---

## 📌 In-grid icon / worksheet Shape

<details>
<summary>Expand when the feature changes grid-icon rendering, placement, interaction, or ownership</summary>

<br>

Describe:

```text
When icon appears:
When icon hides:
When icon is created:
When icon is deleted:
Shape name/identity:
Protected-sheet behavior:
```

Consider:

- worksheet deletion;
- Shape deletion;
- stale retained object references;
- protected drawing objects;
- `UserInterfaceOnly`;
- high-frequency selection performance;
- accidental deletion of unrelated Shapes;
- multiple workbook/provider interaction.

Prefer reuse/move/hide over recreate-on-every-selection unless measurement shows
a reason to change.

</details>

---

## ⏱️ Timer / live clock

<details>
<summary>Expand when the feature uses Application.OnTime, periodic work, or background-like scheduling</summary>

<br>

```text
Interval:
Scheduled procedure:
Qualification:
Stored scheduled time:
Cancellation:
Excel shutdown behavior:
VBA reset behavior:
```

Questions:

- who owns the scheduled callback?
- how is exact cancellation identity retained?
- what prevents duplicate scheduling?
- what happens if the callback fires after the form closes?
- can a failure create an uncontrolled rescheduling loop?

The DatePicker has no background service; periodic behavior is still executed by
Excel/VBA in the user session.

</details>

---

## 🖼️ UserForm / interaction design

<details>
<summary>Expand when the feature changes visual layout, controls, keyboard behavior, overlays, or accessibility</summary>

<br>

Describe the proposed interaction:

```text
Normal mode:
Compact mode:
Keyboard:
Mouse:
Focus:
Overlay behavior:
```

Consider:

- modeless workflow;
- keyboard navigation;
- focus retention;
- screen scaling / DPI;
- high contrast;
- localization;
- runtime-created controls;
- `.frm` / `.frx` impact;
- visual recovery after errors.

Attach a mockup if it materially clarifies the request.

</details>

---

## 🪟 WinAPI / native-window behavior

<details>
<summary>Expand when the feature changes positioning, borderless styling, drag behavior, native APIs, or monitor handling</summary>

<br>

Proposed native behavior:

```text
API(s):
Window identity:
32-bit path:
64-bit path:
Commit point:
Rollback:
Recovery if rollback fails:
```

Current styling result domain:

```vb
DP_WindowStyleResult
```

with:

```text
Attempted
Applied
Committed
RolledBack
RecoveryRequired
FailedStep
LastApiError
```

If the proposal adds another native mutation, explain:

- whether it belongs in the same transaction;
- how partial commit is detected;
- how rollback works;
- how the post-operation state is independently verified;
- what happens when the window handle cannot be resolved;
- multi-monitor behavior.

</details>

---

## 🧪 Regression harness / testability

<details>
<summary>Expand when the feature needs new tests, changes harness behavior, or is difficult to reproduce naturally</summary>

<br>

How should the feature be tested?

- [ ] Standard regression case
- [ ] UserForm UI smoke
- [ ] Failure-path test
- [ ] Dirty-start test
- [ ] Cleanup test
- [ ] Multi-workbook test
- [ ] Multi-provider test
- [ ] Settings persistence test
- [ ] Formula/write-back test
- [ ] Protected-sheet test
- [ ] Native-state test
- [ ] 32-bit Office
- [ ] 64-bit Office
- [ ] Multi-monitor
- [ ] Package-level `.xlam` test
- [ ] Demo update
- [ ] Manual characterization only

```text
Proposed test cases:
```

Current harness states:

```text
PASS
FAIL
FAIL_CLEANUP
FAIL_DIRTY_START
INCOMPLETE_SKIPPED
```

A feature should not weaken the meaning of `PASS`.

### Hard-to-reach failures

If the feature needs a test-only seam:

```text
Why natural input cannot reproduce the failure:
How the seam is armed:
How it is consumed/reset:
How production behavior is protected:
How side effects are verified independently:
```

</details>

---

## 📦 Deployment / packaging

<details>
<summary>Expand when the feature behaves differently in embedded source, .xlam, demo workbook, or Ribbon package</summary>

<br>

Should the feature support:

- [ ] Embedded source
- [ ] `.xlam`
- [ ] Demo `.xlsm`
- [ ] Ribbon-enabled package
- [ ] All supported deployment modes

```text
Deployment-specific behavior:
Additional package content:
Upgrade impact:
Removal/uninstall impact:
```

Consider whether the feature requires:

- new Ribbon XML;
- new binary form resources;
- new release asset;
- workbook event wiring;
- host-specific configuration.

Source and package behavior should not silently diverge.

</details>

---

## 🚦 Release / provenance / automation

<details>
<summary>Expand when the feature concerns CI, release evidence, manifests, checksums, workflow health, or artifact provenance</summary>

<br>

What is the requested guarantee?

- [ ] Static repository checks
- [ ] Workflow health / failed-run visibility
- [ ] VBA source certification
- [ ] Package-level `.xlam` / `.xlsm` smoke
- [ ] Exact source ↔ release asset binding
- [ ] Checksums
- [ ] Machine-readable manifest
- [ ] Release gate
- [ ] Other

```text
Requested evidence:
Failure condition:
Release should be blocked when:
```

Keep these concepts distinct:

```text
source regression
package validation
workflow success
artifact provenance
```

A passing source regression alone does not cryptographically prove a later-saved
binary was built from that exact source.

### Repository credentials

If workflow changes are proposed, explain:

```text
Secrets required:
Environment:
Permissions:
Third-party actions:
Contributor-controlled code executed:
```

The analytics `TRAFFIC_TOKEN` should remain isolated from jobs that execute
repository build/test code.

</details>

---

## 🔐 Security / privacy impact

<details>
<summary>Expand when the feature touches trust boundaries, external data, logging, credentials, or destructive behavior</summary>

<br>

Does the proposal:

- [ ] Write workbook data
- [ ] Read workbook values
- [ ] Persist data in the registry
- [ ] Log worksheet addresses / metadata
- [ ] Execute callbacks by name
- [ ] Use WinAPI
- [ ] Add a GitHub secret
- [ ] Add an external dependency
- [ ] Access the network
- [ ] Add a release binary/resource
- [ ] None of the above

```text
Security/privacy impact:
Sensitive data involved:
Mitigation:
```

Do not propose storing credentials in DatePicker settings.

If the feature fundamentally requires secret management, explain why that belongs
inside this project rather than the host application.

</details>

---

# 🧪 Validation proposal

Even before implementation, define what evidence would prove the feature works.

```text
Compile:
Regression:
UI smoke:
Manual scenario:
Failure path:
Cleanup/recovery:
Package test:
Environments:
```

### Suggested acceptance matrix

| Scenario | Expected behavior |
|---|---|
| Normal path | |
| Invalid / unsupported input | |
| Partial failure | |
| Recovery | |
| Excel restart | |
| VBA reset | |
| Protected sheet | |
| Multiple providers | |
| Mixed-version provider | |
| 32-bit Office | |
| 64-bit Office | |

Delete rows that genuinely do not apply.

---

## 🖥️ Platform / environment considerations

Check known concerns:

- [ ] Office 32-bit
- [ ] Office 64-bit
- [ ] Different Microsoft 365 channels/builds
- [ ] Windows versions
- [ ] Multiple monitors
- [ ] DPI/display scaling
- [ ] High contrast/accessibility
- [ ] Protected worksheets
- [ ] Excel Tables
- [ ] Formula cells
- [ ] Multiple workbooks
- [ ] Other add-ins
- [ ] Embedded + `.xlam`
- [ ] VBA project reset
- [ ] Ribbon package
- [ ] Macro-security / Trusted Location policy
- [ ] Other

Explain any environment-specific assumption:

```text

```

---

## 📖 Documentation impact

Which documentation would need to change if accepted?

- [ ] README
- [ ] INSTALLATION
- [ ] CONTRIBUTING
- [ ] SECURITY
- [ ] CHANGELOG
- [ ] Wiki
- [ ] Demo
- [ ] Code comments / procedure banners
- [ ] No user-facing documentation impact
- [ ] Unsure

```text
Documentation concept that must be explained:
```

---

## 📎 Mockups / pseudocode / references

Attach sanitized material that clarifies the request.

Useful:

- UI mockup;
- pseudocode;
- before/after workflow;
- small VBA example;
- platform documentation;
- related issue/PR;
- comparable implementation.

Do not attach confidential/client workbooks or credentials.

---

## 📝 Additional context

```text
Related issue/PR:
Known limitation:
Open design question:
Other context:
```

---

<!--
Maintainer triage prompts:

1. Is this a feature, or a bug against current documented behavior?
2. Can current supported API already express it?
3. Does it expand destructive write scope?
4. Does it consume caller-owned/application-wide Excel state?
5. Does it change provider ownership or persistence identity?
6. Can partial success occur, and is the result observable?
7. Is recovery explicit?
8. Is it backward compatible?
9. Can it be regression-tested deterministically?
10. Does it belong in core DatePicker, host integration, or release tooling?
-->

<div align="center">

**Feature principle: start from the workflow, keep destructive intent explicit, preserve caller ownership, and define the evidence before the implementation.**

</div>
