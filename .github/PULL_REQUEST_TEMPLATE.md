<!--
  VBA-DATETIMEPICKER pull request template

  Keep the PR focused. Delete sections that do not apply rather than filling
  the template with "N/A".

  The core sections near the top are intended for every substantive PR.
  The subsystem sections below are collapsed on purpose: expand the ones your
  change touches and delete the rest.

  Evidence matters more than checkbox volume. Record only tests and environments
  actually used.
-->

<div align="center">

# 🔀 VBA-DATETIMEPICKER Pull Request

**Focused change · Explicit contract · Reproducible evidence · Known boundaries**

</div>

---

## 📌 Summary

<!--
What changed and why?

Prefer:
  "Preserve formulas during ordinary write-back and report formula skips through
   DP_WriteResult."

Avoid:
  "Various fixes and cleanup."
-->

## 🎯 Problem / motivation

<!--
What concrete defect, limitation, maintenance problem, or user need does this PR
address? State the pre-change behavior where useful.
-->

## 🔗 Related issue

```text
Closes #
```

<!-- Use "Refs #" instead when the PR does not fully close the issue. -->

---

## 🏷️ Type of change

- [ ] 🐛 Functional / compatibility fix
- [ ] 🎯 Write-scope / data-integrity fix
- [ ] 🧮 Formula-safety change
- [ ] 🔒 Application-state / runtime-ownership fix
- [ ] 🪟 WinAPI / native-window hardening
- [ ] ✨ Backward-compatible feature
- [ ] ♻️ Internal refactor with no intended supported-behavior change
- [ ] 🧪 Regression-harness / test change
- [ ] 🖼️ Demo change
- [ ] 🎛️ Ribbon / packaging change
- [ ] 📖 Documentation-only change
- [ ] 🧹 Repository / release-engineering maintenance
- [ ] 🔐 Security-related change

## 🎚️ Affected surface

**Supported consumer surface**

- [ ] 🚀 Startup / lifecycle — `DP_Start`, `DP_Stop`
- [ ] 🖱️ Picker UI — `DP_Show`, `DP_Close`, `DP_Preload`, `DP_Hide`
- [ ] 📅 Direct write-back — `DP_Today`, `DP_Now`
- [ ] 📊 Explicit Table fill — `DP_FillTableColumn`
- [ ] 🛠️ Runtime recovery — `DP_RepairRuntime`
- [ ] 🪪 Lease recovery — `DP_ForceReleaseProviderLease`
- [ ] ⚙️ Settings namespace — `M_Settings_SetNamespace`, `M_Settings_GetNamespace`
- [ ] 🧾 `DP_WriteResult`
- [ ] 🪟 `DP_WindowStyleResult`
- [ ] 🔢 Public enums / constants

**Excel integration**

- [ ] 🧠 Application event manager
- [ ] ⌨️ `Application.OnKey`
- [ ] 🧷 Cell context menu / CommandBars
- [ ] 📌 In-grid worksheet icon
- [ ] ⏱️ `Application.OnTime`
- [ ] 🎛️ RibbonX callbacks / `customUI14.xml`
- [ ] 🗃️ Registry-backed settings
- [ ] 🪪 Process-visible provider lease

**Component / repository**

- [ ] 🖼️ `UF_DatePicker.frm/.frx`
- [ ] 🧩 `cDatePickerManager`
- [ ] 🏷️ `cDatePickerLabelHook`
- [ ] 🪟 WinAPI styling / positioning / drag
- [ ] 🧪 `M_cDP_Test`
- [ ] 🖼️ Demo source
- [ ] 📦 Release asset / provenance
- [ ] ⚙️ Repository configuration / workflow
- [ ] 📖 Documentation only

---

## 🔒 Public API and Semantic Versioning

```text
Supported behavior changed:
Backward compatible:
Suggested release:          patch / minor / major
Migration required:
New Public VBA members:
Removed / renamed members:
```

Review the change against the **supported consumer contract**, not only against
the VBA keyword `Public`.

Confirm impact on:

- supported procedure/function names;
- parameter order;
- required vs optional arguments;
- existing defaults;
- enum numeric values;
- structured-result fields and meaning;
- safe selected-cell write-back default;
- formula-preservation default;
- settings namespace semantics;
- runtime-provider ownership;
- recovery behavior;
- stable legacy identifiers.

Write:

```text
No supported behavior change
```

when that is genuinely the case.

> [!IMPORTANT]
> The following are compatibility-sensitive:
>
> ```text
> VBA_DATETIMEPICKER
> ```
>
> as the legacy settings base name and context-menu identifier, plus the
> established `DP_...` supported API.
>
> A technically `Public` callback or test seam is not automatically a supported
> consumer API. If this PR changes the distinction, explain it explicitly.

---

## 🧭 Behavioral contract

Complete the lines that matter to this PR.

```text
Intended target/scope:
Caller-owned state preserved:
Partial-success behavior:
Failure / refusal behavior:
Recovery path:
Persistence impact:
Runtime-ownership impact:
Known limitation introduced or retained:
```

<!--
This section is deliberately short. It should let a reviewer understand the
contract before reading the diff.
-->

---

# ✅ Validation

## Testing performed

```text
Debug → Compile VBAProject          →

TST_DP_RunAll                       →
TST_DP_RunAll_WithUISmoke           →

Manual DP_Show / DP_Close           →
Manual DP_Stop / DP_Start           →
Manual DP_RepairRuntime              →

Other focused/manual validation     →
```

Delete tests that do not apply rather than marking an unperformed test as
successful.

### Harness verdict

Paste the **actual** summary line:

```text
State=; Run=; Passed=; Failed=; CleanupFailures=
```

The latest recorded v1.2.0 standard pack has reached:

```text
State=PASS; Run=302; Passed=302; Failed=0; CleanupFailures=0
```

That is a historical baseline, **not a required fixed assertion count**.

> [!CAUTION]
> Only `PASS` is a passing run.
>
> ```text
> FAIL
> FAIL_CLEANUP
> FAIL_DIRTY_START
> INCOMPLETE_SKIPPED
> ```
>
> are not release-quality evidence, even when some or all executed assertions
> passed.

### Evidence classification

- [ ] VBA compile
- [ ] Automated regression
- [ ] UI smoke
- [ ] Manual scenario validation
- [ ] Source inspection
- [ ] Package-level validation
- [ ] Platform/documented contract
- [ ] Unresolved hypothesis remains

<!--
Do not call a source inspection "tested" or a one-host non-reproduction
"verified everywhere".
-->

---

## 🖥️ Validation environment

Record only environments actually tested.

```text
Source tag / full commit SHA:
Excel product/version/build:
Office bitness:                   32-bit / 64-bit
Windows version:
Workbook host:                    .xlsm / .xlsb / .xlam
Deployment:                       embedded / add-in
Display scaling:
Monitors:
Worksheet protection state:
WinAPI setting:
Other add-ins loaded:
Other DatePicker provider:        none / version / deployment
Settings namespace:               legacy default / value
Application.EnableEvents at start:
```

For multi-provider behavior, record **both provider versions**.

---

# 📋 Always required

## 🧹 Source and repository hygiene

- [ ] Current branch/base was confirmed before committing.
- [ ] The project compiled with the complete source set required by the change.
- [ ] Changed components were exported to the correct repository paths.
- [ ] `UF_DatePicker.frm` and matching `UF_DatePicker.frx` remain synchronized.
- [ ] Procedure/module `UPDATED` metadata was reviewed where affected.
- [ ] No conflict markers, duplicate procedures, or accidental VBE export noise remain.
- [ ] The diff contains only intended changes.
- [ ] No generated workbook/add-in binary was committed unintentionally.
- [ ] No lock, backup, log, credential, signing key, client data, or production-data file is included.
- [ ] `.gitattributes` / `.gitignore` behavior remains intentional where affected.
- [ ] No unrelated formatting churn is mixed into the functional change.

## 🔒 Caller-owned state and cleanup

- [ ] Normal DatePicker paths preserve the caller's `Application.EnableEvents`.
- [ ] `DP_RepairRuntime` remains the explicit recovery exception that may re-enable events.
- [ ] Every new registration has a deliberate teardown path.
- [ ] Teardown removes only DatePicker-owned resources.
- [ ] `Application.OnTime` cancellation uses the exact schedule/procedure identity recorded at registration.
- [ ] Context-menu cleanup uses the DatePicker identifier rather than collection position.
- [ ] Retained worksheet `Shape` references are not assumed live without validation.
- [ ] No unsolicited low-level production `MsgBox` was introduced.

## 🎯 Data integrity

- [ ] Normal interactive Table behavior remains selected-cell / resolved-target by default.
- [ ] Whole-column Table write requires explicit `DP_FillTableColumn` intent.
- [ ] Formula cells remain preserved by default.
- [ ] `OverwriteFormulas:=True` remains explicit/destructive where supported.
- [ ] Array-formula / non-writable cases remain controlled.
- [ ] Partial writes remain observable rather than silently promoted to full success.
- [ ] Write-back restores caller event state on success and failure paths.

## 🧾 Error and diagnostic integrity

- [ ] `Err` values are captured before calls / `On Error` statements can destroy them.
- [ ] Anything reachable from an error handler cannot replace the original failure.
- [ ] `On Error Resume Next` scopes remain narrow and deliberate.
- [ ] The implemented error policy matches the procedure banner/documentation.
- [ ] Structured results describe the real post-operation state.
- [ ] Refusal / fail-closed outcomes are not reported as success.

## 📖 Documentation

- [ ] README updated where user-facing behavior/API changed.
- [ ] INSTALLATION updated where deployment/startup/upgrade behavior changed.
- [ ] CONTRIBUTING updated where engineering rules changed.
- [ ] SECURITY updated where trust/security boundaries changed.
- [ ] Wiki updated where the reviewed v1.2.0 contract changed.
- [ ] `CHANGELOG.md` updated for user-visible behavior.
- [ ] No documentation change is required.

> [!TIP]
> Documentation belongs in this PR when it describes this PR.
>
> A knowingly stale README/Wiki after a behavior change is part of the defect.

---

# 🧩 Subsystem review sections

<!--
Expand the sections that apply and delete the others before review.
-->

## 🎯 Write-back / Excel Table scope

<details>
<summary>Expand when target resolution, Today/Now, calendar selection, Table behavior, or DP_FillTableColumn changed</summary>

<br>

```text
Entry point(s):
Selection / target before:
Selection / target after:
Table expansion possible:
Confirmation behavior:
Predicted scope:
AttemptedCount:
WrittenCount:
Skipped / failed classifications:
```

- [ ] Selected-cell behavior inside an Excel Table remains safe by default.
- [ ] Broad Table scope is reached only through explicit caller/user intent.
- [ ] Previewed scope is checked against `AttemptedCount`, not `WrittenCount`.
- [ ] Locked/formula/failure outcomes do not make a correct target prediction look wrong.
- [ ] Target metadata is attached once from the resolver rather than re-derived inconsistently.
- [ ] Multi-area behavior remains deliberate.
- [ ] No write path returns without populating its structured result semantics.
- [ ] User-facing shortfall messaging occurs once per operation, not once per cell/area.

### Scope regression cases

- [ ] Omitted/default safe scope
- [ ] Explicit internal broad-scope path where still required by architecture/tests
- [ ] `DP_FillTableColumn`
- [ ] Locked cells
- [ ] Formula cells
- [ ] Array-formula / non-writable cell
- [ ] Multi-area target where applicable

</details>

---

## 🧮 Formula policy / `DP_WriteResult`

<details>
<summary>Expand when formulas, write classifications, counts, addresses, or result fields changed</summary>

<br>

```text
New / changed result fields:
Formula default:
Override behavior:
Address reporting:
Bulk-path behavior:
Balance assertion updated:
```

Completed-result invariant:

```text
AttemptedCount =
    WrittenCount +
    LockedSkippedCount +
    FormulaSkippedCount +
    FailedCount
```

- [ ] Ordinary formulas are preserved by default.
- [ ] Date-returning formulas are not treated as literals merely because they display a date.
- [ ] Explicit overwrite does not become an implicit/persisted global preference.
- [ ] Array formulas remain non-overridable where the engine cannot safely write them.
- [ ] Counts remain exact even when diagnostic address strings are capped.
- [ ] Bulk writes contribute their own correct counts.
- [ ] `WrittenCount` is not inferred by subtraction when Excel can silently refuse an assignment.
- [ ] `M_WriteBack_DescribeShortfall` / reporting reflects any new classification.
- [ ] `TST_DP_AssertWriteResultBalances` or equivalent balance coverage was updated.

</details>

---

## ⚙️ Settings persistence / namespace

<details>
<summary>Expand when registry settings, defaults, namespace resolution, or settings lifecycle changed</summary>

<br>

```text
Legacy default affected:
Namespace behavior affected:
Effective application name:
Migration behavior:
Load / save timing:
```

- [ ] `VBA_DATETIMEPICKER` remains the stable legacy base name unless a breaking change is explicit.
- [ ] Effective settings name is resolved centrally.
- [ ] `M_Settings_SetNamespace` still must run before settings are loaded.
- [ ] Namespace remains locked after load.
- [ ] No automatic migration silently copies legacy settings into a new namespace.
- [ ] Namespace is a stable deployment identity, not a version/path/workbook-name identity.
- [ ] Settings namespace is not used as runtime-provider identity.
- [ ] No secret/credential storage was introduced into DatePicker settings.
- [ ] Default and namespaced persistence were both tested where relevant.

</details>

---

## 🪪 Runtime provider ownership

<details>
<summary>Expand when DP_Start, DP_Stop, DP_RepairRuntime, the CommandBar lease, or multi-provider behavior changed</summary>

<br>

```text
Provider A version/deployment:
Provider B version/deployment:
Lease acquired before registration:
Refused-provider behavior:
Owner teardown behavior:
Stale-lease behavior:
Force-release behavior:
Mixed-version boundary:
```

Lease name:

```text
__VBA_DATETIMEPICKER_RUNTIME_PROVIDER_LEASE__
```

- [ ] Ownership is acquired before the first shared Excel registration.
- [ ] A second participating current-version provider is refused before it can displace owner state.
- [ ] `DP_Stop` is ownership-guarded.
- [ ] `DP_RepairRuntime` is ownership-guarded.
- [ ] A refused provider cannot dismantle the owner.
- [ ] Release requires a local token and exact matching lease token.
- [ ] Unreadable/ambiguous lease state fails closed.
- [ ] Stale lease after VBA reset is not automatically reclaimed.
- [ ] `DP_ForceReleaseProviderLease` remains explicit operator-only recovery.
- [ ] Force release is never called automatically.
- [ ] `v1.2.0 + pre-v1.2.0` remains documented as unprotected unless the architecture genuinely changes.
- [ ] Multi-provider manual scenarios were run when ownership logic changed.

</details>

---

## ⌨️ Keyboard / access paths

<details>
<summary>Expand when OnKey, right-click, grid icon, Ribbon access, or access-path settings changed</summary>

<br>

```text
Shortcut:
Right-click:
Grid icon:
Ribbon:
Zero-built-in-access configuration:
Third-party OnKey predecessor behavior:
```

This remains a valid configuration:

```text
ShowRightClick = False
ShowGridIcon = False
EnableKeyboardShortcut = False
```

- [ ] Keyboard registration reflects explicit `EnableKeyboardShortcut` configuration only.
- [ ] No code silently enables the keyboard shortcut because other paths are disabled.
- [ ] Removal restores Excel default handling rather than claiming to restore an unobservable predecessor.
- [ ] Documentation does not claim DatePicker can read/restore a third-party `Application.OnKey` assignment.
- [ ] Right-click and grid-icon setters do not create hidden coupling to keyboard state.
- [ ] New access paths document registration lifetime, teardown, ownership, conflicts, persistence and testability.

</details>

---

## 📌 In-grid icon / worksheet lifecycle

<details>
<summary>Expand when Shape creation, tracking, movement, purge, sheet deletion, or protection behavior changed</summary>

<br>

```text
Tracked shape lifecycle:
Sheet deletion case:
Shape deletion case:
Protected-sheet behavior:
High-frequency selection behavior:
```

- [ ] A retained `Shape` reference is validated before reuse.
- [ ] Stale tracked state is cleared rather than treated as live.
- [ ] Purge removes only DatePicker-owned shapes.
- [ ] High-frequency selection handling favors show/move/hide/reuse over needless recreate cycles.
- [ ] Sheet/workbook lifecycle cases were tested where changed.
- [ ] Protected-sheet behavior is documented and does not silently weaken host protection.

</details>

---

## 🧠 Application events / manager lifecycle

<details>
<summary>Expand when cDatePickerManager, selection handling, workbook lifecycle, or EnableEvents behavior changed</summary>

<br>

```text
Event(s) affected:
Manager creation:
Caller EnableEvents at entry:
Caller EnableEvents at exit:
Re-entrancy / recursion treatment:
Workbook lifecycle treatment:
```

- [ ] `cDatePickerManager` remains the central Application-event orchestrator.
- [ ] No duplicate host `SelectionChange` orchestration was introduced without an explicit new contract.
- [ ] Normal paths do not force `Application.EnableEvents = True`.
- [ ] Caller event state is preserved through failure paths.
- [ ] `DP_RepairRuntime` remains explicitly different from ordinary bootstrap/show paths.
- [ ] High-frequency event handlers remain best-effort where intended.
- [ ] Teardown removes/neutralizes manager state cleanly.

</details>

---

## ⏱️ Timer / `Application.OnTime`

<details>
<summary>Expand when live clock scheduling or teardown changed</summary>

<br>

```text
Scheduled procedure:
Scheduled time stored:
Qualification:
Cancellation path:
Failure/restart behavior:
```

- [ ] Cancellation uses the exact scheduled timestamp.
- [ ] Cancellation uses the exact qualified procedure identity.
- [ ] The component does not claim to enumerate/own unrelated Excel `OnTime` jobs.
- [ ] A failed/stale callback cannot create an uncontrolled rescheduling loop.
- [ ] `DP_Close` / `DP_Stop` behavior remains deliberate for pending timer state.
- [ ] No timer survives lifecycle boundaries unintentionally.

</details>

---

## 🪟 WinAPI / `DP_WindowStyleResult`

<details>
<summary>Expand when window style, positioning, drag, native declarations, rollback, or fault injection changed</summary>

<br>

```text
Native API(s):
Handle resolver:
32-bit path:
64-bit path:
Commit point:
Post-commit failure point(s):
Rollback:
RecoveryRequired condition:
Independent native verification:
```

- [ ] VBA7 and legacy declaration branches remain correct where applicable.
- [ ] 32-bit and 64-bit paths were considered/tested as applicable.
- [ ] Window handle resolution remains centralized unless a deliberate test isolation requires duplication.
- [ ] Ambiguous native zero returns are handled according to the API contract.
- [ ] Native commit is distinguishable from full application success.
- [ ] Post-commit failures attempt rollback.
- [ ] Rollback failure produces `RecoveryRequired=True`.
- [ ] `DP_WindowStyleResult` fields describe the real state.
- [ ] A test seam remains one-shot and cannot poison a later real call.
- [ ] Tests verify native state independently rather than trusting only the returned result.
- [ ] `UserForm.Caption` is not used in a way that unintentionally restores `WS_CAPTION`.
- [ ] Multi-window / real-host behavior was checked where relevant.

</details>

---

## 🧪 Regression harness

<details>
<summary>Expand when test/M_cDP_Test.bas or harness lifecycle changed</summary>

<br>

```text
Suites added/changed:
Assertions added/changed:
Run-state logic changed:
Preflight changed:
Scratch-sheet setup changed:
Cleanup changed:
Result-sheet behavior changed:
```

Run states:

```text
PASS
FAIL
FAIL_CLEANUP
FAIL_DIRTY_START
INCOMPLETE_SKIPPED
```

- [ ] New suite is wired into the runner and safe dispatcher.
- [ ] Mandatory-suite completion remains observable.
- [ ] Dirty-start preflight runs before the current run mutates the workbook.
- [ ] Module-state and persistent worksheet dirty-start evidence remain deliberate.
- [ ] A dirty start does not execute normal suites and manufacture evidence.
- [ ] Cleanup failure cannot be reported as `PASS`.
- [ ] Final observable cleanup state is checked before `PASS`.
- [ ] `SuiteFail` captures `Err` before anything can reset it.
- [ ] Calls that clear the caller's handler are followed by deliberate re-arming where required.
- [ ] Harness test/fault injections are reset/consumed safely.
- [ ] `TST_DP_SCRATCH` does not remain after a clean pass.

### `Worksheets.Add` partial-success recovery

- [ ] Zero-new-sheet failure remains fail-loud.
- [ ] Exactly-one-new-sheet failure is resolved by object identity and validated before adoption.
- [ ] More-than-one-new-sheet failure remains ambiguous / delete-none / fail.
- [ ] The harness does not blindly retry `Worksheets.Add`.
- [ ] Scratch-sheet identity semantics continue to protect `FAIL_DIRTY_START`.

</details>

---

## 🖼️ UserForm

<details>
<summary>Expand when UF_DatePicker.frm/.frx, runtime controls, keyboard interaction, overlays, or visual behavior changed</summary>

<br>

```text
Visual area changed:
.frm changed:
.frx changed:
Runtime-created controls changed:
Keyboard behavior changed:
Scaling / DPI checked:
```

- [ ] Matching `.frm` / `.frx` pair was exported and reviewed together where required.
- [ ] `TST_DP_RunAll_WithUISmoke` was run.
- [ ] `DP_Show` / `DP_Close` was manually checked.
- [ ] Modeless behavior remains intact.
- [ ] No unexpected focus/key capture was introduced.
- [ ] Runtime label hooks are created/released correctly.
- [ ] Settings overlays open/close correctly where affected.
- [ ] Borderless/title-bar behavior remains recoverable where affected.
- [ ] Any DPI/high-contrast/accessibility limitation is stated rather than assumed solved.

</details>

---

## 🎛️ RibbonX / Office package

<details>
<summary>Expand when Ribbon callbacks, customUI14.xml, or release packaging changed</summary>

<br>

```text
Callbacks changed:
customUI14.xml changed:
Package tested:
Embedded deployment tested:
.xlam deployment tested:
External resources introduced:
```

- [ ] Every `onAction` name resolves to the intended public callback exactly.
- [ ] Ribbon callbacks remain thin and delegate to normal DatePicker API behavior.
- [ ] XML/package changes were tested in an actual Office package, not only by calling the VBA callback directly.
- [ ] No unexpected external resource or executable reference was introduced.
- [ ] Ribbon behavior remains optional; non-Ribbon entry points still work as documented.
- [ ] Package/source synchronization is stated accurately.

</details>

---

## 🖼️ Demo source / release demo

<details>
<summary>Expand when demo/M_DP_DEMO.bas, demo/M_DEMO_BUILDER.bas, or demo release assets changed</summary>

<br>

```text
Demo sections changed:
Builder changed:
Rebuilt using:
Visual inspection:
Release demo binary rebuilt:
```

- [ ] Demo was rebuilt from committed source.
- [ ] Rebuilding does not accumulate duplicate controls/content.
- [ ] New controls have valid `OnAction` targets.
- [ ] Demo illustrates current v1.2.x behavior, not superseded Table/formula behavior.
- [ ] Formula examples remain formulas after normal picker interaction.
- [ ] Table examples distinguish selected-cell write from explicit column fill.
- [ ] No personal data, machine-local path, secret, or scratch content entered the demo source/artifact.
- [ ] Any release binary was built from the intended source and package-level smoke tested.

</details>

---

## 📦 Release engineering / provenance

<details>
<summary>Expand when release assets, manifests, release notes, workflow gates, checksums, or certification evidence changed</summary>

<br>

```text
Release candidate:
Source commit:
Assets:
Source regression:
Package smoke:
Manual scenarios:
Checksums / manifest:
Exact source↔asset binding:
```

- [ ] Source regression evidence and package-level evidence are distinguished.
- [ ] A passing `TST_DP_RunAll` is not presented as proof of an unverified later-saved binary.
- [ ] Asset names/version match the intended release.
- [ ] Release notes / CHANGELOG match the actual source.
- [ ] No unsupported exact-SHA provenance claim is made.
- [ ] Checksums/manifests, if added, bind the correct files.
- [ ] Generated `.xlam` / `.xlsm` remains a release artifact rather than ordinary repository source.
- [ ] Workflow failures, missed runs, or credential failures cannot silently masquerade as certification where this PR touches automation.

</details>

---

## 🔐 Security / repository automation

<details>
<summary>Expand when SECURITY.md, GitHub Actions, credentials, permissions, release trust, or sensitive logging changed</summary>

<br>

```text
Workflow(s):
Secrets:
Environment:
Permissions before:
Permissions after:
Third-party actions:
Sensitive logging:
```

Current analytics workflow boundary:

```text
.github/workflows/daily-traffic.yml
TRAFFIC_TOKEN
analytics environment
contents: write
issues: write
```

- [ ] Analytics credentials remain isolated from jobs that execute contributor-controlled build/test code.
- [ ] Workflow permissions are least-privilege for the actual behavior.
- [ ] New third-party actions are pinned to immutable full commit SHAs.
- [ ] No secret/token is printed or committed.
- [ ] No credential was moved from an environment-scoped secret to a wider scope without explicit review.
- [ ] Traffic/analytics automation is not represented as software-quality certification.
- [ ] Security-relevant release/provenance claims are accurate.
- [ ] Confidential workbook/user data is not added to logs or artifacts.

</details>

---

## 📚 Documentation consistency

<details>
<summary>Expand for documentation-heavy or cross-document contract changes</summary>

<br>

```text
README:
INSTALLATION:
CONTRIBUTING:
SECURITY:
CODE_OF_CONDUCT:
CHANGELOG:
Wiki pages:
Reviewed Wiki commit metadata:
```

- [ ] Same API name/signature is documented consistently everywhere.
- [ ] `DP_FillTableColumn` is described as returning `DP_WriteResult`.
- [ ] Formula-preservation policy is consistent.
- [ ] Normal `Application.EnableEvents` preservation vs `DP_RepairRuntime` is consistent.
- [ ] Zero-built-in-access-path configuration remains documented as valid.
- [ ] Provider lease / mixed-version boundary is consistent.
- [ ] Settings namespace / before-load rule is consistent.
- [ ] Source regression is not conflated with packaged release certification.
- [ ] Wiki page review/version metadata was updated if the reviewed contract changed.
- [ ] No rewritten-away / superseded commit SHA was introduced into issue/document history.

</details>

---

# 💬 Reviewer notes

<!--
Trade-offs, known limitations, tests not run, environments not covered,
follow-up work, or evidence boundaries.

A limitation stated here is reviewable.
A limitation discovered after release is a defect.
-->

## Known boundary / not proved by this PR

```text

```

## Follow-up

```text

```

---

<div align="center">

**PR principle: state the contract, show the evidence, preserve caller ownership, and make uncertainty visible.**

</div>
