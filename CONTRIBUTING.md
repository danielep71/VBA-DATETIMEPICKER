<div align="center">

# 🤝 Contributing to VBA-DATETIMEPICKER

### Engineering guidance for safe, reviewable Excel/VBA contributions

[![Conduct](https://img.shields.io/badge/read_first-code_of_conduct-6f42c1?style=flat-square)](CODE_OF_CONDUCT.md)
[![Security](https://img.shields.io/badge/read_first-security_policy-d73a49?style=flat-square)](SECURITY.md)
[![Source](https://img.shields.io/badge/model-source--first-217346?style=flat-square)](#-source-first-development-model)
[![Harness](https://img.shields.io/badge/regression-TST__DP__RunAll-0969da?style=flat-square)](#-required-validation)
[![Release](https://img.shields.io/badge/release-evidence_required-d97706?style=flat-square)](#-release-certification-boundary)

<br>

**Predictable write scope · Caller-owned Excel state · Recoverable runtime · Explicit ownership · Structured outcomes · Regression evidence**

<br>

[Before you start](#-before-you-start)
&nbsp;·&nbsp;
[Workflow](#-source-first-development-model)
&nbsp;·&nbsp;
[Validation](#-required-validation)
&nbsp;·&nbsp;
[Public API](#-public-api-and-supported-surface)
&nbsp;·&nbsp;
[Runtime ownership](#-runtime-provider-ownership)
&nbsp;·&nbsp;
[Write-back](#-write-back-contract)
&nbsp;·&nbsp;
[Source style](#-source-style)
&nbsp;·&nbsp;
[Pull requests](#-pull-requests)

</div>

---

## 🧭 Before you start

Read:

- [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)
- [SECURITY.md](SECURITY.md)
- [README.md](README.md)
- [CHANGELOG.md](CHANGELOG.md)

For non-trivial work, **open or reference an issue before implementation**.

Scope agreed before coding is cheaper than scope discovered in review.

Small documentation corrections, obvious comment fixes, and genuinely trivial
one-line changes can normally go directly to a focused pull request.

> [!IMPORTANT]
> This repository is not a generic VBA playground. It is a reusable component
> that can write into user data, register application-wide Excel behavior, store
> persistent settings, and modify a native UserForm window. Contributions are
> reviewed against those operational consequences, not only against whether the
> code compiles.

### Engineering priorities

A change that trades away one of these properties must say so explicitly.

| Priority | Why it is non-negotiable |
|---|---|
| 🎯 **Predictable write scope** | The component writes into user data. Writing more cells than the user intended is worse than declining the operation. |
| 🧮 **Formula preservation by default** | A displayed date does not imply permission to destroy the formula that produced it. |
| 🔒 **Caller-owned Excel state** | `Application.EnableEvents`, `OnKey`, `OnTime`, context menus and worksheet objects may be shared with other code. |
| 🪪 **Explicit runtime ownership** | Two DatePicker providers must not silently dismantle each other's application-wide registrations. |
| 🧹 **Recoverable runtime** | Every transient registration must have a deliberate teardown or recovery path. |
| 🧾 **Structured outcomes** | Partial success must be observable; “no exception” is not sufficient evidence that all work completed. |
| 🪟 **Transactional native UI changes** | A half-applied borderless window style must be rolled back or reported as requiring recovery. |
| ⚙️ **32-bit and 64-bit parity** | A defect that exists only on the untested bitness is still a defect. |
| 🧪 **Permanent regression coverage** | A fix without a regression test is an invitation for the same defect to return. |
| 📖 **Readable exported source** | The repository files are the review artifact. The VBE is only the editing host. |

---

## ⚡ Quick reference

```text
Debug → Compile VBAProject       compile the imported project

TST_DP_RunAll                    standard source regression pack
TST_DP_RunAll_WithUISmoke        standard pack + UserForm lifecycle smoke
TST_DP_ReportEnvironment         print host/setup diagnostics

DP_Show                          manual picker smoke
DP_Close                         manual close/teardown check
DP_RepairRuntime                 explicit runtime repair
DP_Stop                          normal owning-provider teardown
```

The latest recorded standard regression result, on the `v1.2.1` source cycle, is:

```text
State=PASS; Run=431; Passed=431; Failed=0; CleanupFailures=0
```

That number is a useful baseline, **not a hard-coded acceptance criterion**.
Adding a legitimate regression assertion should increase it.

What matters is the run state, failure count, cleanup result, and whether the
expected suites actually completed.

---

## 🧱 Quality model

The project uses several different kinds of evidence. Do not collapse them into
one claim.

| Evidence class | What it can establish |
|---|---|
| **Compile** | VBA syntax, declarations, references and cross-module compatibility compile on the tested host |
| **Automated regression** | Deterministic behaviors covered by `M_cDP_Test.bas` |
| **UI smoke** | UserForm can load, show, interact minimally and close on the tested host |
| **Manual scenario validation** | Behavior that requires multiple workbooks/providers, Excel lifecycle actions, or platform interaction |
| **Source inspection** | Structural properties visible directly in the exported source |
| **Platform contract** | Behavior documented by Excel/VBA/Windows interfaces where direct observation is limited |
| **Unresolved hypothesis** | A suspected cause not yet proved by a controlled reproduction |

Use the strongest term the evidence supports.

Do not write `fully verified` when the evidence is only `source inspected` or
`could not reproduce on one host`.

That distinction matters particularly for Excel session behavior.

---

## 📁 Project layout

```text
VBA-DATETIMEPICKER/
├─ .github/
├─ assets/
├─ demo/
│  ├─ M_DEMO_BUILDER.bas
│  └─ M_DP_DEMO.bas
├─ dist/
│  └─ README.md
├─ images/
├─ src/
│  ├─ classes/
│  │  ├─ cDatePickerManager.cls
│  │  └─ cDatePickerLabelHook.cls
│  ├─ forms/
│  │  ├─ UF_DatePicker.frm
│  │  └─ UF_DatePicker.frx
│  ├─ modules/
│  │  └─ M_DatePicker.bas
│  └─ ribbon/
│     └─ customUI14.xml
├─ test/
│  └─ M_cDP_Test.bas
├─ CHANGELOG.md
├─ CODE_OF_CONDUCT.md
├─ CONTRIBUTING.md
├─ LICENSE
├─ README.md
└─ SECURITY.md
```

### Component responsibilities

| Component | Responsibility |
|---|---|
| `M_DatePicker.bas` | Public API, settings, write-back, context menu, keyboard registration, grid icon, timer, provider lease, Ribbon callbacks and WinAPI helpers |
| `cDatePickerManager.cls` | Application-level Excel event orchestration and selection/workbook lifecycle |
| `cDatePickerLabelHook.cls` | `WithEvents` routing for runtime-created UserForm labels |
| `UF_DatePicker.frm/.frx` | Modeless DatePicker UI, calendar rendering, overlays, settings panel and keyboard interaction |
| `customUI14.xml` | Optional RibbonX layout |
| `M_cDP_Test.bas` | Regression harness |
| `M_DEMO_BUILDER.bas` / `M_DP_DEMO.bas` | Source-defined demo construction |

> [!IMPORTANT]
> `UF_DatePicker.frm` and `UF_DatePicker.frx` are one logical source component.
> The `.frm` is text; the `.frx` is its binary resource companion. Import,
> export, review and commit them together whenever the form resource changes.

---

## 📦 Source and binary policy

This is a **source-first repository**.

The repository deliberately tracks the code and resources required to reproduce
the component rather than opaque workbook/add-in binaries.

| Artifact | Tracked? | Policy |
|---|:---:|---|
| `*.bas` | ✅ | Authoritative VBA source |
| `*.cls` | ✅ | Authoritative class source |
| `*.frm` | ✅ | UserForm text source |
| `src/forms/UF_DatePicker.frx` | ✅ | Required binary form-resource companion |
| `customUI14.xml` | ✅ | RibbonX source |
| Demo builder/source modules | ✅ | Source of the release demo |
| Regression module | ✅ | Executable test source |
| `DATETIMEPICKER-demo-v<version>.xlsm` | ❌ | Generated release asset |
| The packaged `.xlam` | ❌ | Generated release asset |

The `.frx` is the intentional exception to the “reviewable text” preference:
it is source, not a distributable build artifact.

> [!IMPORTANT]
> Change the demo by editing `demo/M_DP_DEMO.bas` and
> `demo/M_DEMO_BUILDER.bas`, not by hand-editing a workbook that later becomes a
> release asset. A hand-edited binary cannot be reconstructed from the commit
> that supposedly produced it.

Release binaries belong on the
[GitHub Releases](https://github.com/danielep71/VBA-DATETIMEPICKER/releases)
page.

Release provenance and machine-readable release evidence are still being
strengthened under the repository's release-engineering work. Do not claim an
artifact is exact-SHA reproducible unless the release process actually records
that binding.

---

## 🧩 Architecture and responsibility boundaries

### `M_DatePicker`

`M_DatePicker` is intentionally large because several Excel callback mechanisms
require public procedures in a standard module.

It owns:

```text
supported DP_* entry points
M_Settings_* persistence
M_WriteBack_* target resolution and write engine
M_ContextMenu_* integration
M_KeyboardShortcut_* integration
M_GridIcon_* integration
M_Timer_* footer clock
M_Window_* native window helpers
Ribbon_* callbacks
runtime-provider lease
shared DatePicker constants/enums/types
```

Do not add an unrelated responsibility merely because the module already exists.

A future split should preserve callback reachability and supported API names
rather than forcing consumers to understand internal modules.

### `cDatePickerManager`

Owns Application event orchestration.

Do not reproduce its responsibilities with additional `ThisWorkbook` or
worksheet event code unless the change is explicitly about a new host-integration
contract.

One selection change should have one orchestrator.

### `cDatePickerLabelHook`

Owns runtime MSForms label event routing.

Keep it focused on event plumbing. Calendar business logic belongs in the form or
module layer that owns the behavior.

### `UF_DatePicker`

Owns visual state and interaction.

Do not move application-wide ownership decisions into the form merely because a
button triggers them.

The form is a consumer of the runtime infrastructure, not its owner.

---

## 🔒 Public API and supported surface

VBA has two different meanings of `Public` in this project:

```text
technically callable from another module
```

and:

```text
supported as a consumer contract
```

They are not automatically the same.

Excel callbacks, RibbonX, `Application.Run`, test seams, and cross-module VBA
access sometimes require `Public` declarations that are not intended as stable
consumer API.

### Primary supported consumer surface

The README is the current consumer-facing reference. The principal supported
surface includes:

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

Ribbon callbacks are public for Office resolution, not because callers should use
them instead of the normal API.

### Compatibility rules

Preserve unless an explicitly approved breaking release says otherwise:

- procedure/function names;
- parameter order;
- existing optional defaults;
- enum numeric values;
- structured-result field semantics;
- the legacy registry application name `VBA_DATETIMEPICKER`;
- the stable context-menu identifier `VBA_DATETIMEPICKER`;
- selected-cell write-back as the normal default;
- formula preservation as the normal default;
- `DP_RepairRuntime` as the explicit event-repair path;
- namespace-before-load semantics;
- one-provider fail-closed ownership behavior for participating current-version
  providers.

New backward-compatible parameters should normally be **optional and trailing**.

> [!NOTE]
> The repository is separately formalizing the distinction between supported,
> callback and internal public members. Until that classification is fully
> manifested, do not infer support solely from the VBA keyword `Public`.

---

## 🌿 Branch workflow

Do not make routine development changes directly on `main`.

Use a branch that communicates the purpose:

```text
fix/formula-writeback-classification
fix/provider-lease-teardown
feature/new-calendar-option
test/writeback-array-formula
docs/contributing-v1.2-contract
ci/repository-quality-gate
chore/gitignore-hardening
release/v<major>.<minor>.<patch>
```

| Prefix | Use |
|---|---|
| `fix/` | Defect correction |
| `feature/` | Backward-compatible capability |
| `test/` | Regression/test-only work |
| `docs/` | Documentation only |
| `ci/` | Workflow, quality gate or automation |
| `chore/` | Repository/configuration hygiene |
| `release/` | Version integration branch |

Confirm your branch before every commit:

```bash
git status -sb
git branch --show-current
```

### Keep local and remote state explicit

Before pushing a release branch:

```bash
git fetch origin
git rev-list --left-right --count origin/release/<version>...HEAD
```

Interpretation:

```text
0 1    local is one commit ahead
0 0    synchronized
1 0    remote has one commit local does not
```

Do not use force-push as a reflex to resolve an unexpected divergence.

Understand the history first.

---

## 📝 Commit discipline

Keep commits conceptually coherent.

Good:

```text
fix(writeback): preserve formulas by default
test(harness): reject dirty-start runs
docs(readme): document provider lease
chore(gitignore): harden repository ignore policy
```

Avoid commits that combine behavior fixes with unrelated formatting,
documentation rewrites, and repository cleanup unless they are genuinely one
inseparable migration.

For substantive commits, the body should explain:

```text
what changed
why it changed
important invariants
validation performed
related issues
```

Never invent a commit SHA in documentation before the commit exists.

---

## 🔁 Source-first development model

The normal workflow is:

```text
issue / agreed scope
    ↓
branch
    ↓
import source into controlled Excel host
    ↓
edit
    ↓
compile
    ↓
focused validation
    ↓
standard regression
    ↓
manual/UI validation where required
    ↓
export changed source
    ↓
review repository diff
    ↓
update docs/changelog
    ↓
commit
    ↓
push
    ↓
pull request / release integration
```

### Recommended import order

```text
src/modules/M_DatePicker.bas
src/classes/cDatePickerManager.cls
src/classes/cDatePickerLabelHook.cls
src/forms/UF_DatePicker.frm          with UF_DatePicker.frx beside it

test/M_cDP_Test.bas                  when testing

demo/M_DEMO_BUILDER.bas              when working on demo source
demo/M_DP_DEMO.bas                   when working on demo source
```

Then:

```text
VBA Editor → Debug → Compile VBAProject
```

### Export discipline

After editing in the VBE:

1. export every changed component;
2. overwrite the matching repository path;
3. ensure the `.frx` accompanies a changed form;
4. inspect the Git diff;
5. reject VBE-only changes that never reached the repository;
6. reject repository changes that were never re-imported/compiled when they alter
   executable source.

The repository and the tested VBE project must describe the same code.

---

## ✅ Required validation

Validation depends on what changed.

### Baseline for production VBA changes

```text
Debug → Compile VBAProject
TST_DP_RunAll
```

For form/UI changes:

```text
TST_DP_RunAll_WithUISmoke
DP_Show
DP_Close
```

For runtime-registration changes, include the relevant manual lifecycle checks.

For native-window changes, include a real Windows host and the applicable
window-style tests.

### Validation matrix

| Change area | Minimum expected validation |
|---|---|
| Documentation only | Link/anchor/code-sample review |
| Repository config | Git behavior/diff review; no source behavior claim |
| Settings | Compile + `TST_DP_RunAll` + namespace/default-scope checks |
| Write-back | Compile + `TST_DP_RunAll` + affected manual scope case |
| Formula policy | Compile + formula/literal/array-formula regression cases |
| UserForm | Compile + `TST_DP_RunAll_WithUISmoke` + manual show/close |
| Grid icon | Compile + regression + worksheet lifecycle/manual selection check |
| Keyboard/context menu | Compile + regression + registration/removal check |
| Provider lease | Compile + regression + multi-provider manual matrix |
| WinAPI styling | Compile + WindowStyle and WindowRecovery suites + UI smoke on Windows |
| Release artifact | Source validation + regression pack executed inside the packaged artifact + manual certification and provenance checks |

### Regression verdict

A run reports one of:

```text
PASS
FAIL
FAIL_CLEANUP
FAIL_DIRTY_START
INCOMPLETE_SKIPPED
```

Only `PASS` is a passing run.

A run that ends early prints **no summary line at all**. An absent summary is
itself a failure result, not missing output — do not read it as a tooling glitch.

There are 24 standard suites, 25 registered including UI smoke. A run that
reports fewer has skipped something.

A preferred pull-request evidence line is:

```text
State=PASS; Run=431; Passed=431; Failed=0; CleanupFailures=0
```

If the assertion total differs because the suite changed, report the actual
number.

### What `PASS` means

A valid pass requires:

```text
clean preflight
+
mandatory suites completed
+
assertions passed
+
required cleanup succeeded
+
observable final state verified
```

A dirty environment invalidates the evidence.

`FAIL_DIRTY_START` is not “a small test failure”; it means the current run did
not establish the environment whose behavior it would otherwise claim to test.

---

## 🧹 Harness lifecycle and clean-start invariant

The harness uses two independent dirty-start indicators:

```text
module-level run-in-progress state
leftover TST_DP_SCRATCH worksheet
```

Both are necessary.

A VBA project reset clears module variables but does not delete a worksheet.

Preflight must therefore happen **before the current run mutates the workbook**.

Do not move setup work ahead of the dirty-start decision.

A run that detects dirty start should refuse to manufacture test evidence from
an environment it did not create.

---

## 🧪 Harness worksheet-creation recovery

Excel has been observed reporting `1004` from `Worksheets.Add` after a new
worksheet had already appeared.

The accepted recovery state machine is:

```text
Worksheets.Add succeeds
    → normal setup

Worksheets.Add raises
    + zero new sheets
        → fail setup

Worksheets.Add raises
    + exactly one new sheet
        → identify by object identity
        → validate it
        → adopt only if it is the expected usable candidate

Worksheets.Add raises
    + more than one new sheet
        → ownership ambiguous
        → delete none
        → fail setup
```

Do not “fix” this by blindly retrying `Worksheets.Add`.

A retry after a partial commit creates a second sheet and converts a recoverable
state into an ambiguous one.

---

## ⚙️ Application-state ownership

The component touches Excel surfaces that may also be used by business macros or
other add-ins.

| Surface | Contract |
|---|---|
| `Application.EnableEvents` | Normal DatePicker entry points preserve caller state. `DP_RepairRuntime` is the deliberate recovery exception that may re-enable events. |
| `Application.OnKey` | Registration is explicit configuration; removal can restore Excel default but cannot reconstruct an unknown predecessor. |
| `Application.OnTime` | Cancel with the exact scheduled time and qualified procedure identity used to register it. |
| Cell context menu | Manage DatePicker controls by stable tag, not by collection index. |
| Worksheet Shape | Validate liveness before using a retained shape reference; purge only at documented lifecycle boundaries. |
| Registry settings | Resolve through the effective settings namespace; persist only within that resolved scope. |
| Provider lease | Process-visible runtime ownership; never confuse with persistent settings identity. |

> [!CAUTION]
> Never reintroduce a blanket `Application.EnableEvents = True` into normal
> `DP_Start`, `DP_Show`, `DP_Preload`, `M_Picker_EnsureManager`, or other ordinary
> paths. A caller may have disabled events deliberately as part of its own
> transaction.

---

## 🪪 Runtime provider ownership

Several DatePicker surfaces are shared across the Excel process.

Two providers that both believe they own them can replace each other's shortcut,
duplicate/remove context-menu entries, interfere with Application events, cancel
timers, or dismantle resources the other still needs.

`v1.2.1` therefore enforces one process-visible lease, on every entry path, for
participating current-version providers.

Lease name:

```text
__VBA_DATETIMEPICKER_RUNTIME_PROVIDER_LEASE__
```

### Settled rules

**Acquire before shared registration.**

```text
lease first
registrations second
```

Checking after registration is too late.

**Admission is per entry path, not per startup.**

Every path that can reach the manager — `DP_Start`, `DP_Show`, `DP_Preload`,
`DP_Click`, `DP_OpenForActiveCell`, the keyboard handler and `Ribbon_ShowPicker`
— must prove ownership before it acts. Guarding only the startup path leaves the
lease decorative: the ordering rule above was satisfied at `DP_Start` and the
defect shipped anyway ([#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37)).

**Guard teardown as well as startup.**

A refused provider must not be able to call `DP_Stop` or `DP_RepairRuntime` and
dismantle the current owner.

**Release only proven ownership.**

Release requires a local owner token, a readable lease, and a matching stored
token. Unreadable/ambiguous ownership fails closed.

**Stale leases fail closed.**

A project reset can destroy the local token while leaving the process-visible
lease alive. Restarting Excel clears the temporary lease.

**Force release is operator-only.**

`DP_ForceReleaseProviderLease` exists for an informed operator who knows no other
provider is alive. It must never become an automatic recovery step.

### Mixed-version boundary

Current protection is:

```text
v1.2.1 + v1.2.1
    → second participating provider refused on every entry path
```

It is **not**:

```text
v1.2.1 + pre-v1.2.1
    → guaranteed safe
```

Pre-`v1.2.0` code has no lease protocol at all. `v1.2.0` has the protocol but
admits it only at `DP_Start`, so on every other entry path it behaves as a
non-participant while appearing to be one — treat a `v1.2.0` peer as
pre-`v1.2.1`, not as a participating provider.

Any provider-ownership test or documentation must state the versions involved.

---

## 🎯 Write-back contract

Write-back touches user data and therefore has the strictest behavioral
invariants in the component.

### Safe normal scope

Normal interactive selection writes the selected/resolved target.

A single selected cell inside an Excel Table does **not** implicitly mean “fill
the entire table data column”.

Whole-column fill is explicit:

```vb
DP_FillTableColumn
```

Do not reintroduce implicit table growth or whole-column expansion as a
convenience behavior.

### Preview and application must agree

When a broad scope is previewed before confirmation, validate the eventual
attempted target against the predicted scope.

Compare prediction with `AttemptedCount`, not `WrittenCount`.

A predicted 247-cell target that writes 244 because three cells are safely
skipped is still the same target.

A different attempted count means the target itself changed.

---

## 🧮 Formula-preservation contract

Formula cells are preserved by default.

Default behavior:

| Cell | Default |
|---|---|
| Empty/literal | Write |
| Ordinary formula | Preserve |
| Formula returning date/time | Preserve |
| Locked protected cell | Skip |
| Array-formula cell | Fail/non-overridable |
| Other object-model failure | Report failure |

The advanced explicit override:

```vb
OverwriteFormulas:=True
```

can replace ordinary formulas where supported.

Do not turn the override into a persisted global preference without an explicit
API/design decision.

Destructive intent should remain local to the call that requested it.

---

## 🧾 `DP_WriteResult`

All worksheet write-back outcomes use one structured result.

Conceptually it reports attempted cells, written cells, locked skips, formula
skips, other failures, worksheet-qualified addresses, resolved target, table
expansion metadata, area count, and caller event state.

The completed-result balance is:

```text
AttemptedCount =
    WrittenCount +
    LockedSkippedCount +
    FormulaSkippedCount +
    FailedCount
```

If a new skip class is introduced:

1. extend `DP_WriteResult`;
2. extend the balance assertion;
3. extend the shortfall description;
4. add regression coverage;
5. update README/API documentation.

Do not create a parallel second result mechanism.

### Result rules

- every successful exit populates the result;
- every early/refusal path has deliberate result semantics;
- bulk paths account for their own counts;
- cell addresses are diagnostic facts, not a substitute for exact counts;
- a failed write is not inferred from subtraction when Excel can silently decline
  an assignment;
- the write engine collects facts;
- interactive entry points decide whether to show a user-facing summary;
- a technical failure part-way through a multi-area write must still surface a
  populated `DP_WriteResult`. An exception must never be the only outcome once
  user data has been mutated; raise only when no cell produced any outcome,
  because only then is nothing lost by raising;
- multi-area totals must not depend on the order Excel enumerates
  `Target.Areas`.

Do not show one modal dialog per area/cell from the low-level write loop.

### Internal test seams

These are `Public` only so the harness can reach them. They are not supported
API, and are to be classified `internal` under [#25](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/25):

```text
M_Lease_Test_SilenceRefusalReport
M_Lease_Test_RefusalReportCount
M_Settings_ResolveKeyboardShortcutOnSave
M_WriteBack_Test_SetFaultInjection
```

Add a seam only when the path cannot be produced otherwise, and say so in the
PR. `M_WriteBack_Test_SetFaultInjection` exists because
`M_WriteBack_TryWriteCell` classifies every per-cell failure it can observe and
raises nothing, so a genuine technical error inside the area loop is precisely
what a test has no way to arrange.

---

## ⚙️ Settings persistence and namespaces

The stable legacy application name remains:

```text
VBA_DATETIMEPICKER
```

Without a namespace, that scope remains the backward-compatible default.

A deployment can opt into:

```vb
M_Settings_SetNamespace "MyDeployment"
```

before settings are loaded.

Effective persistence then uses a stable derived name such as:

```text
VBA_DATETIMEPICKER__MyDeployment
```

### Settled namespace rules

- **Resolution is centralized.** Do not hand-build effective names at individual
  registry call sites.
- **Namespace locks after load.** Avoid `read from A / write to B`.
- **No automatic migration.** A new namespace starts from defaults.
- **Namespace is not runtime identity.** Persistence survives Excel; the provider
  lease does not.
- **Do not version the namespace.** A deployment identity should survive software
  upgrades.

---

## ⌨️ Keyboard shortcut contract

The DatePicker shortcut uses `Application.OnKey`.

Excel exposes no getter for the previous assignment.

Therefore the DatePicker cannot reliably capture and later restore an unknown
third-party predecessor.

### Registration is explicit

The shortcut is registered because:

```text
EnableKeyboardShortcut = True
```

not because the component decides the user needs at least one built-in path.

This is valid:

```text
ShowRightClick = False
ShowGridIcon = False
EnableKeyboardShortcut = False
```

Zero built-in interactive access paths are permitted.

The host may still use `DP_Show`, RibbonX, or a caller-provided button/macro.

### Removal restores Excel default

Removing the DatePicker shortcut restores Excel's default handling for the key.

It does **not** restore an unknown third-party macro.

Do not add documentation or code that claims otherwise.

---

## 🖼️ Grid-icon lifecycle

The in-grid icon is a worksheet `Shape`.

A retained `Shape` object reference is not proof that the shape is still live.
The sheet or shape may have been deleted independently.

Before using a tracked shape, validate that the retained object is still live.

If it is stale:

```text
clear tracked state
resolve/create from current worksheet context
```

Do not blindly call methods on a stale tracked reference.

High-frequency selection handling should continue to favor:

```text
show / move / hide / reuse
```

over:

```text
delete / recreate on every selection
```

unless measurement proves otherwise.

---

## 🪟 WinAPI and borderless-window changes

Windows styling is optional, but when used it must be defensive.

### Declaration rules

- maintain VBA7 and legacy declaration branches;
- preserve 32-bit / 64-bit correctness;
- keep window-handle resolution centralized;
- do not duplicate native declarations without a specific test/isolation reason;
- treat zero native returns according to the API contract rather than assuming
  zero always means failure.

### Styling transaction

Conceptually:

```text
read original style
    ↓
write target style
    ↓
refresh non-client frame
    ↓
redraw
```

A failure after style commit must attempt rollback.

`DP_WindowStyleResult` distinguishes:

```text
Attempted
Applied
Committed
RolledBack
RecoveryRequired
FailedStep
LastApiError
```

Do not collapse these states back into a single Boolean.

### Test seams

A narrow native-failure test seam is acceptable when it exists only to drive
otherwise unreachable failure paths, is consumed one-shot, cannot poison a later
real call, and is documented as test infrastructure rather than consumer API.

Where practical, the test should inspect native state independently rather than
trusting only the result returned by the function under test.

---

## 🛡️ Error policy

Every procedure should have an error policy that matches its role.

| Policy | Use |
|---|---|
| **Raise** | Caller must know the operation failed |
| **Best effort** | High-frequency UI, cleanup, timer/callback paths where unrelated work should continue |
| **Structured result** | Operation can partially succeed and the caller needs the outcome |
| **Fail closed** | Ownership/safety cannot be established reliably |

### Rules

**Capture `Err` before doing anything else.**

At an error label:

```vb
ErrNumber = Err.Number
ErrDescription = Err.Description
ErrSource = Err.Source
ErrLine = Erl
```

Then perform cleanup/logging.

Every `On Error` statement resets the `Err` object, and so does `Err.Clear`. Any
cleanup that arms or disarms an error handler therefore destroys the original
failure before you can report it. Raise from the captured locals, never from the
live `Err`. This is not a corner case: production source carries 243 `Err.Clear`
sites and 87 raises that read the live `Err`, and two of them shipped as defects
in `v1.2.0` ([#48](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/48)).

**Anything called from an error handler must not replace the original failure.**

**Keep `On Error Resume Next` narrow.**

**No unsolicited low-level `MsgBox`.**

Interactive entry points may display deliberate summaries when that is part of
the contract.

**Preserve caller handlers where VBA semantics require re-arming.**

Some production procedures use `On Error GoTo 0`. If the regression suite relies
on its own `SuiteFail` handler, an immediate re-arm after such a call is
intentional.

---

## ✒️ Source style

The existing exported source is the style reference.

### Module hygiene

- `Option Explicit` in every module, class, and form;
- do **not** add `Option Private Module` to `M_DatePicker.bas` while Excel/Office
  callbacks require public resolution;
- preserve private-instancing metadata for the DatePicker classes;
- keep form `.frm` and `.frx` synchronized.

### Procedure banners

For substantial procedures, use the established sections where relevant:

```text
PURPOSE
WHY THIS EXISTS
INPUTS
RETURNS
BEHAVIOR
ERROR POLICY
DEPENDENCIES
NOTES
UPDATED
```

The banner should explain the contract, not merely repeat the procedure name.

### Procedure body structure

Prefer clear internal sections such as:

```text
DECLARE
INITIALIZE
VALIDATE
RESOLVE
APPLY
CLEANUP
ERROR HANDLER
```

Use the vocabulary that matches the procedure.

### Comments

Comment intent and invariants, not syntax.

Useful:

```vb
'Acquire ownership before any application-wide registration can be displaced
```

Not useful:

```vb
'Set variable to True
Flag = True
```

### Naming

| Prefix | Meaning |
|---|---|
| `DP_` | Supported DatePicker-facing entry point/type/constant where documented |
| `M_Settings_` | Settings/persistence infrastructure |
| `M_WriteBack_` | Target resolution and write engine |
| `M_GridIcon_` | Worksheet icon lifecycle |
| `M_ContextMenu_` | Right-click integration |
| `M_KeyboardShortcut_` | `Application.OnKey` integration |
| `M_Timer_` | Footer clock / `OnTime` |
| `M_Window_` | Native UserForm window helpers |
| `Ribbon_` | RibbonX callback |
| `TST_DP_` | Regression harness |
| `gDP_` | Shared runtime state |
| `m...` | Private module/class state |
| `cDatePicker...` | Class modules |

### Performance

High-frequency Excel/UI paths should avoid unnecessary object-model chatter.

Prefer bulk range operations, cached references, reuse of existing Shapes and
controls, and one resolution per operation.

Correctness still outranks micro-optimization.

A faster wrong-scope write is worse than a slower safe one.

---

## 🧹 Repository hygiene

The root `.gitattributes` is authoritative for line endings and binary handling.

Current policy includes:

```text
*.bas / *.cls / *.frm   CRLF working tree
Markdown/config         LF
*.frx                   binary, non-line-mergeable
Office packages         binary
images/media            binary
```

Do not normalize `.frx` as text.

Do not manually convert VBA exports to LF in the working tree.

Let Git apply the repository policy.

The `.gitignore` intentionally excludes generated Office binaries and local
artifacts while preserving authoritative source.

Never commit Office lock files, editor backups, logs/dumps, local secrets,
signing keys, client data, generated `.xlam` / `.xlsm` build output, the workbook
used as your VBE editing host, or unrelated formatting churn.

Published release evidence — certification worksheets, hash manifests — belongs
in the GitHub Release process rather than being committed from a workstation.

### One-time normalization

If `.gitattributes` rules are intentionally changed:

```bash
git add --renormalize .
```

Commit that normalization separately from behavioral changes.

---

## 📖 Documentation expectations

Documentation belongs in the **same change** as the behavior it describes.

| Change | Documentation impact |
|---|---|
| Supported API | README API section + relevant Wiki page + CHANGELOG |
| Write-back behavior | README write-back/safety section + Wiki + CHANGELOG |
| Settings namespace | README settings section + Wiki + CHANGELOG |
| Access-path behavior | README entry-point/keyboard section + Wiki + CHANGELOG |
| Provider lease | README ownership/limitations + Wiki + CHANGELOG |
| UserForm behavior | README where user-visible + relevant UI Wiki page |
| Regression harness | README testing section + contributing/release guidance where needed |
| Release process | Release documentation/checklist + CHANGELOG as appropriate |
| Security behavior | SECURITY.md and coordinated disclosure process |

### Wiki status

The Wiki rewrite for `v1.2.1` is complete.

Pages corrected in `v1.2.1` are stamped with the certified commit:

```text
Applies to:      v1.2.1
Reviewed commit: 7d55cc7
```

Pages untouched since the `v1.2.0` review keep that baseline, which remains
accurate for them:

```text
Applies to:      v1.2.0
Reviewed commit: 6435c91
```

Stamp a page you edit with the commit you verified it against — never carry a
previous page's stamp forward.

When a later source change makes a Wiki statement stale:

1. update the affected page;
2. update its version/review metadata;
3. keep README and Wiki language consistent;
4. do not leave a knowingly stale “reviewed commit” stamp.

The branch source is authoritative for code behavior; documentation must be
updated to match it.

---

## 🔍 Documentation accuracy rules

Do not write documentation that is stronger than the code.

Wrong:

```text
restores the previous third-party keyboard shortcut
```

when Excel provides no getter for that assignment.

Wrong:

```text
supports multiple providers safely
```

when the current contract is one participating current-version provider and
mixed-version coexistence remains unprotected.

Wrong:

```text
release-certified
```

when only the source regression pack was run.

Precision is a feature.

---

## 🚀 Pull requests

Keep a pull request focused on **one coherent concern**.

A reviewer should be able to answer:

```text
What problem does this solve?
What contract changes?
What contract stays unchanged?
What evidence proves the change?
What remains unverified?
```

### Pull-request checklist

```text
[ ] Related issue linked for non-trivial work
[ ] Scope is focused
[ ] Branch/base is correct
[ ] Public API / SemVer impact assessed
[ ] Write scope impact assessed
[ ] Formula behavior assessed
[ ] Application-state ownership assessed
[ ] Provider-ownership impact assessed
[ ] Settings-persistence impact assessed
[ ] Error/partial-success behavior assessed
[ ] Error handlers capture Err.Number/Source/Description before any cleanup, On Error statement or Err.Clear
[ ] 32/64-bit impact assessed where relevant
[ ] Debug → Compile VBAProject passed
[ ] TST_DP_RunAll result recorded for production changes
[ ] TST_DP_RunAll_WithUISmoke run when UI/form changed
[ ] Manual scenario evidence recorded where automation cannot prove the claim
[ ] Cleanup / recovery checked
[ ] Documentation updated in the same PR
[ ] CHANGELOG.md updated for user-visible behavior
[ ] Demo impact assessed
[ ] Release-binary impact assessed
[ ] No workbook binary or confidential data committed
```

### Suggested PR evidence block

```text
Environment
-----------
Excel:
Office bitness:
Windows:
Deployment: embedded / add-in
Other DatePicker provider loaded: none / version

Validation
----------
Compile: PASS
TST_DP_RunAll:
TST_DP_RunAll_WithUISmoke:
Manual checks:

Known boundary
--------------
<what this evidence does not prove>
```

Record only environments actually tested.

---

## 🧪 Regression changes

When fixing a defect:

```text
reproduce
    ↓
write or extend regression
    ↓
prove old behavior fails where practical
    ↓
fix
    ↓
prove regression passes
```

Tests should assert the invariant, not merely the exact implementation.

Good:

```text
formula remains unchanged by default
```

Less useful:

```text
helper X was called
```

Good:

```text
failed rollback produces RecoveryRequired=True
```

Where possible, verify side effects independently.

---

## 🧯 Recovery-path changes

A recovery path deserves explicit tests because it normally executes only when
something has already failed.

For changes to `DP_RepairRuntime`, provider-lease recovery, window-style
rollback, scratch-sheet recovery, grid-icon purge, timer teardown, context-menu
cleanup, or keyboard removal, document:

1. normal state;
2. failure/stranded state;
3. recovery action;
4. post-recovery observable state;
5. what cannot be recovered automatically.

Never make a recovery routine more destructive merely to make it “succeed”.

Failing closed can be correct behavior.

---

## 🚦 Access-path changes

Right-click, grid icon, keyboard shortcut, Ribbon and caller-supplied macros are
different entry surfaces.

Do not create hidden coupling between:

```text
ShowRightClick
ShowGridIcon
EnableKeyboardShortcut
```

This remains valid:

```text
False
False
False
```

Changing one setter should not silently enable another access path.

If a new access path is introduced, document ownership, registration lifetime,
teardown, conflict behavior, persisted setting (if any), testability, and its
interaction with zero-built-in-access configuration.

---

## 🔐 Security and confidential material

Security vulnerabilities follow [SECURITY.md](SECURITY.md).

Do not open a public issue containing exploit details or credentials.

Do not attach a real client/business workbook to a pull request merely to
demonstrate a bug.

Excel files can contain hidden metadata, VBA, external connections, cached
values, names, links and private information.

Use a sanitized minimal reproduction.

By contributing code or documentation, you confirm that you have the right to
submit it under this repository's license.

---

## 🧭 Release certification boundary

`TST_DP_RunAll` is the standard source regression pack.

It is **not by itself proof that a published `.xlam` or demo `.xlsm` is the exact
tested artifact**.

A release candidate should separately establish, as applicable:

```text
source compiles
standard regression PASS
UI smoke PASS
regression pack executed inside the packaged .xlam, not only in the source host
manual lifecycle checks PASS
provider-lease scenario checks where affected
release assets rebuilt from intended source
asset version/naming correct — check the exact published filename, it has varied
artifact SHA-256 computed after the artifact was built AND tested
tag points at the tested commit, not at a later merge commit
release notes/changelog correct
provenance/evidence recorded to the extent supported by the current release process
```

Two of those lines exist because `v1.2.1` certification found the gaps the hard
way. The packaged-artifact run had never been possible — a result sheet was built
into `ThisWorkbook`, so the pack could not run inside an add-in at all — which
means every earlier release's regression figure was source-level evidence only.
And a defect found in the harness is a release blocker, not test-only noise:
three candidate commits were voided during `v1.2.1` certification, all four
defects were in `test/`, and `git diff` across them over `src/` is empty.

The repository is continuing to improve automated workflow health and exact-SHA
release evidence.

Until those mechanisms are complete, do not represent their future guarantees as
already implemented.

---

## 📊 Operational workflows are not software-quality gates

Repository analytics/traffic automation may exist for operational reporting.

That does not make it static analysis, VBA regression execution, release
certification, or artifact provenance.

Keep analytics credentials and release-quality automation conceptually and
operationally separate.

A failed traffic collection job is not evidence that the DatePicker source is
defective.

A passing traffic job is not evidence that the DatePicker source is correct.

---

## 📜 Semantic versioning guidance

Assess version impact against the supported consumer contract.

### Patch

Typical patch:

```text
bug fix
diagnostic correction
safe internal hardening
documentation correction
```

### Minor

Typical minor:

```text
new backward-compatible capability
new optional trailing parameter
new supported entry point
new opt-in behavior
```

### Major

Potential major:

```text
rename/remove supported public member
change required parameter order
change enum numeric meaning
change established default incompatibly
remove stable persistence identifier
```

Do not label a change “internal” merely because most callers will not notice it.

If it changes the supported contract, discuss the SemVer impact explicitly.

---

## 🤝 Review culture

A strong review comment states location, risk, evidence, and whether the change
is required or optional.

Example:

> `M_WriteBack_Apply`: the new branch increments `WrittenCount` after an
> assignment that Excel can silently refuse. That breaks the structured-result
> invariant for non-writing object-model outcomes. Please classify the cell
> before the attempt or verify the write independently and add a regression
> case.

That is more useful than:

> This is wrong.

Review the software precisely and the contributor respectfully.

See [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md).

---

## 📄 License

By contributing, you agree that your contribution is licensed under the
project's [MIT License](LICENSE).

Do not submit third-party code whose license is incompatible or whose origin
cannot be established.

If adapting another implementation, identify the source and license in the pull
request.

---

## 👤 Maintainer

Maintained by **Daniele Penza**.

For design questions, architectural proposals, broader API changes, or work that
will require several commits, open an issue before implementation.

For security concerns, use the private process in
[SECURITY.md](SECURITY.md).

---

<div align="center">

### Contribution principle

**Make the intended scope explicit. Preserve what the caller owns. Report what actually happened. Test the failure path.**

</div>
