<div align="center">

# 🤝 Contributing

**Thank you for improving VBA-DATETIMEPICKER**

[![Conduct](https://img.shields.io/badge/read_first-code_of_conduct-6f42c1?style=flat-square)](CODE_OF_CONDUCT.md)
[![Security](https://img.shields.io/badge/read_first-security_policy-d73a49?style=flat-square)](SECURITY.md)
[![Harness](https://img.shields.io/badge/gate-TST__DP__RunAll-217346?style=flat-square)](#-required-validation)
[![Style](https://img.shields.io/badge/style-house_conventions-0969da?style=flat-square)](#-source-style)

</div>

---

## 🧭 Before you start

Read [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md) and [SECURITY.md](SECURITY.md).
**Open an issue before non-trivial work** — scope agreed in advance is far
cheaper than scope discovered in review.

Tiny fixes — typos, comment corrections, obvious one-line bugs — can go straight
to a pull request.

This project holds these above convenience, and a change that trades one away
needs to say so explicitly:

| Priority | Why it is non-negotiable |
|---|---|
| 🎯 **Predictable write scope** | The component writes into user data. A write the user did not ask for is worse than no write at all. |
| 🔒 **Caller-owned application state** | `Application.EnableEvents`, `OnKey`, `OnTime` and context menus belong to whoever set them. Changing one underneath a running macro breaks it silently. |
| 🧹 **Recoverable runtime** | Every registration the component makes must be removable. A stuck grid icon or an orphaned timer outlives the workbook that created it. |
| 🧾 **Deterministic diagnostics** | A failure that is not reported did not happen, as far as the caller is concerned. |
| ⚙️ **32-bit and 64-bit parity** | A defect that only appears on the other bitness is invisible to the person who wrote it. |
| 🧪 **Permanent regression coverage** | A fix without a test is a fix with a scheduled regression. |
| 📖 **Readable exported source** | The `.bas` file is the review artifact; the VBE is not. |

---

## ⚡ Quick reference

```text
Debug → Compile VBAProject           compile the imported project
TST_DP_RunAll                        the regression gate
TST_DP_RunAll_WithUISmoke            adds the form smoke suite
DP_RepairRuntime                     recover a wedged session
```

There is no CI that executes VBA — a hosted runner has no Excel — so the harness
is a manual step on a real host. Automating the parts that *can* run without
Excel is tracked in
[#15](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/15).

---

## 📁 Project layout

```text
VBA-DATETIMEPICKER/
├─ src/
│  ├─ modules/
│  │  └─ M_DatePicker.bas          public API, settings, write-back, integrations
│  ├─ classes/
│  │  ├─ cDatePickerManager.cls    Application event manager
│  │  └─ cDatePickerLabelHook.cls  per-label WithEvents router
│  ├─ forms/
│  │  ├─ UF_DatePicker.frm         modeless UserForm code-behind
│  │  └─ UF_DatePicker.frx         binary companion — never edit by hand
│  └─ ribbon/
│     └─ customUI14.xml            optional RibbonX layout
├─ test/
│  └─ M_cDP_Test.bas               regression harness
├─ demo/
│  ├─ M_DEMO_BUILDER.bas           worksheet-building primitives
│  └─ M_DP_DEMO.bas                composes the DatePicker demo sheet
├─ dist/
│  └─ README.md                    binaries ship as release assets
├─ CONTRIBUTING.md
├─ README.md
└─ …
```

> [!IMPORTANT]
> `UF_DatePicker.frx` is binary and is excluded from normalization by
> `.gitattributes`. It must travel with the `.frm` on every import and export.
> A `.frm` committed without its matching `.frx` produces a form that compiles
> and renders wrongly.

### Binary policy

**No workbook or add-in binary is tracked.** Everything under version control is
text that can be read in a diff.

| Artifact | Tracked | Policy |
|---|---|---|
| `src/forms/UF_DatePicker.frx` | **yes** | Part of the form source. Not optional and not a build artifact. The only tracked binary, and it must travel with its `.frm`. |
| `DATETIMEPICKER-demo-vx.y.z.xlsm` | no | Built from `demo/M_DP_DEMO.bas` by `DP_Demo_CreateDemoSheet`. Published as a release asset. |
| `DATETIMEPICKER-vx.y.z.xlam` | no | Build output, excluded by `.gitignore`. Published as a release asset. |

Release asset names carry the version, so a downloaded file is identifiable
without opening it:

```text
DATETIMEPICKER-vx.y.z.xlam
DATETIMEPICKER-demo-vx.y.z.xlsm
```

Names contain no spaces. GitHub replaces spaces with dots on upload, so a
spaced name is published under a different name than the one documented.

> [!IMPORTANT]
> Change the demo by editing `demo/M_DP_DEMO.bas`, never by editing a workbook.
> A hand-edited workbook produces a release asset that no committed source can
> reproduce, and nothing in the repository would show it happened.

Both release assets are rebuilt from the tagged source at release time. Neither
carries a manifest binding it to the commit that produced it yet; that is
[#16](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/16).

---

## 🌿 Branch workflow

Do not make routine development changes directly on `main`.

```text
fix/enablevents-caller-state
test/application-state-suite
docs/readme-known-limitations
ci/pin-actions
feature/explicit-fill-column
release/v<major>.<minor>.<patch>
```

| Prefix | For |
|---|---|
| `fix/` | A defect with an issue |
| `feature/` | Backward-compatible new capability |
| `test/` | Regression coverage only |
| `docs/` | Prose only, no code effect |
| `ci/` | Workflow, gate or tooling |
| `chore/` | Repository configuration and hygiene |
| `release/` | Integration branch for a version |

Confirm the current branch in GitHub Desktop before every commit.

> [!WARNING]
> Editing files through the GitHub web interface while holding unpushed local
> commits diverges the branch. If it happens, `git pull --rebase` replays your
> work on top; setting `pull.rebase true` once avoids the prompt entirely.

---

## 🔁 Import, edit, compile, test, export

Recommended import order:

```text
src/modules/M_DatePicker.bas
src/classes/cDatePickerManager.cls
src/classes/cDatePickerLabelHook.cls
src/forms/UF_DatePicker.frm          with .frx alongside
test/M_cDP_Test.bas
demo/M_DEMO_BUILDER.bas              when working on the demo
demo/M_DP_DEMO.bas                   with M_DEMO_BUILDER.bas
```

Workflow:

1. Confirm the current branch.
2. Import the required components into a controlled workbook.
3. Compile with `Debug → Compile VBAProject`.
4. Run `TST_DP_RunAll`.
5. Run `TST_DP_RunAll_WithUISmoke` when the form changed.
6. Perform the manual checks for whatever you touched.
7. Re-export each changed component over its matching repository path.
8. Review the GitHub Desktop diff.
9. Update documentation and `UPDATED` dates.
10. Commit and push.
11. Open a pull request against the agreed base.

Only commit source that actually changed. Never commit the workbook you edited
in.

> [!CAUTION]
> The harness runs with `Application.EnableEvents = False` and restores it
> afterwards. An aborted run can leave Excel in that state and leave its scratch
> and result worksheets behind. The next run now detects this before it changes
> anything and reports `FAIL_DIRTY_START` rather than passing against an
> environment it did not establish. Restart Excel and delete the leftover sheets
> before treating a later run as valid.

---

## ✅ Required validation

For production code changes:

```text
Debug → Compile VBAProject
TST_DP_RunAll
TST_DP_RunAll_WithUISmoke        when the form changed
DP_Show / DP_Close               manual
DP_RepairRuntime                 manual
```

Quote the harness summary line in the pull request:

```text
INFO | Harness | Summary | State=PASS; Run=214; Passed=214; Failed=0; CleanupFailures=0
```

`State` is the verdict. It is one of:

```text
PASS                the run started clean, passed, and tore down completely
FAIL                an assertion failed
FAIL_CLEANUP        assertions passed but teardown did not complete
INCOMPLETE_SKIPPED  a dispatched suite never returned
FAIL_DIRTY_START    the run began in an environment a previous run left behind
```

Only `PASS` is a pass. `FAIL_DIRTY_START` outranks the others because it makes
the rest of the report untrustworthy rather than merely bad — restart Excel,
delete the leftover sheets, and run again before reading anything else.

A run that ended early has no summary line at all, so its absence is itself the
verdict.

Record only environments actually tested.

> [!NOTE]
> Suite-level failures currently report `Error 0 -` because the handler resets
> `Err` before reading it. If you hit one, the real error number is not in the
> output. Tracked in
> [#18](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/18).

---

## 🔒 Public API compatibility

Preserve unless an explicitly approved breaking release requires otherwise:

- public procedure names — `DP_Start`, `DP_Show`, `DP_Close`, `DP_Preload`,
  `DP_Hide`, `DP_Today`, `DP_Now`, `DP_RepairRuntime`, `DP_Stop`;
- parameter order and optional defaults;
- enum values — `DP_WriteAction`, `DP_ClockMode`, `DP_SizeMode`;
- settings getter and setter semantics;
- the `VBA_DATETIMEPICKER` registry application name;
- the `VBA_DATETIMEPICKER` command-bar tag;
- `DP_RepairRuntime` as the recovery path.

The two stable legacy identifiers exist for backward compatibility. Renaming
either silently discards every user's saved settings, or orphans every context
menu entry a previous version installed.

---

## ⚙️ Application state

The component touches surfaces it does not own. Every one of them belongs to
whoever set it.

| Surface | Rule |
|---|---|
| `Application.EnableEvents` | Read it, report it, never change it. `DP_RepairRuntime` is the single sanctioned exception. |
| `Application.OnKey` | One registration, removable, and never assumed to be the only one in the session. |
| `Application.OnTime` | Cancelled with the exact scheduled time and qualified procedure name it was created with. |
| Cell context menus | Tagged `VBA_DATETIMEPICKER`, removed by tag, never by index. |
| Worksheet shapes | Named consistently and purged only at documented boundaries. |
| Registry settings | Written under the stable application name, on explicit save only. |

> [!CAUTION]
> `M_Picker_EnsureManager` used to force `Application.EnableEvents = True` on
> every call, which broke any business macro that had deliberately suppressed
> events. It no longer does. Do not reintroduce that pattern anywhere reachable
> from `DP_Start`, `DP_Show`, or `DP_Preload` — see
> [#2](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/2).

---

## 🎯 Write scope

Write-back is the part of this component that touches user data, so it gets
stricter treatment than the UI.

- The visible Excel selection is the strongest signal of intended scope.
- Expanding beyond the selection requires explicit user intent, not inference.
- Formulas are not collateral. They are preserved by default and replaced only
  when a caller passes `OverwriteFormulas:=True`. Existing literal values are
  fair game; a formula is not.
- A partial write must be reportable, not silent.
- `M_WriteBack_Apply` captures and restores the caller's `EnableEvents` state,
  including on the failure path. Preserve that.

### The provider lease

Only one DatePicker copy may own an Excel session. The lease is a hidden
`Temporary` `CommandBar` named `__VBA_DATETIMEPICKER_RUNTIME_PROVIDER_LEASE__`,
carrying one hidden `Temporary` control whose `Tag` identifies the marker and
whose `Parameter` holds the owner's ephemeral token.

Four rules, all settled:

- **Acquire before the first shared registration.** A copy that registers and
  then discovers the conflict has already displaced the owner's keyboard
  shortcut. `DP_Start` claims the lease first or exits.
- **Guard teardown, not just startup.** Refusing a second provider at startup
  protects nothing while its `DP_Stop` or `DP_RepairRuntime` still dismantles the
  owner. Both verify ownership before touching anything.
- **Never delete a lease you cannot prove you own.** Release requires a local
  token, a lease still carrying it, and an exact match. An unreadable lease is
  ambiguous, never free — a bar this component cannot interpret belongs to
  something.
- **The lease is runtime state and is never persisted.** A registry-backed lease
  would survive an Excel restart and block startup forever. `Temporary:=True`
  means Excel removes it at shutdown, which is the correct lifetime.

`DP_ForceReleaseProviderLease` is the single deliberate exception: an operator
command that deletes the lease regardless of ownership. It must never be called
automatically, and no lifecycle routine calls it.

> [!NOTE]
> The regression harness claims the lease during setup, through
> `TST_DP_ResetDatePickerArtifacts`. That is the one place the guard is overridden
> and it is deliberate: the lease outlives the VBA project that created it, so
> re-importing a module strands a lease no project can release, and a regression
> run is a single-provider environment by definition.

### Settings persistence

`DP_SETTINGS_APP_NAME` is a stable legacy identifier and the default application
name. It does not change.

Four rules, all settled:

- **Resolution stays centralized.** Every registry read and write obtains its
  application name from `M_Settings_GetEffectiveAppName`. A namespace applied at
  28 call sites would eventually be wrong at one of them.
- **The namespace locks after load.** `M_Settings_SetNamespace` raises once
  settings are in memory. Repointing afterwards would write values read from one
  namespace into another.
- **Runtime identity is never persisted here.** The one-provider lease in
  [#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37) is runtime
  state that dies with Excel. A registry-backed lease would survive a restart and
  block startup forever. A settings namespace is persistent by design; a lease
  must not be.
- **Nothing is migrated automatically.** A new namespace starts from defaults.
  Copying the shared namespace across would reintroduce the coupling — the
  integration toggles and `HolidayCallback` especially — that the namespace exists
  to remove.

> [!NOTE]
> Persistence scope and runtime ownership are different identities with different
> lifetimes. [#14](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/14)
> will need both; it must not use one as the other.

### The keyboard shortcut

`Application.OnKey` holds one assignment per key for the Excel session, and Excel
exposes no getter for it. Two decisions follow from that, and both are settled —
do not reopen them without new platform capability.

**Registration reflects explicit configuration only.** The shortcut is bound when
the user enabled it, and never because other access paths were disabled. An
earlier rule forced it on when right-click and the in-grid icon were both off,
which took a session-wide binding from whoever held it on behalf of a user who
had not asked for it. Zero built-in interactive access paths is a permitted
configuration; `DP_Show`, RibbonX and caller-supplied buttons remain.

That rule lived in six places, expressed two different ways — three direct
assignments and three through a local in the setters. If you are changing this
area, search for the *invariant* rather than the assignment.

**Teardown restores Excel default.** `M_KeyboardShortcut_Remove` calls
`Application.OnKey KEY` with no macro. This cannot restore a displaced
third-party binding, because that binding was never observable. It is the least
harmful of the three available behaviors:

```text
Application.OnKey KEY       Excel default restored; displaced binding lost
Application.OnKey KEY, ""   key swallowed; worse
do nothing                  stale callback into a possibly unloaded project; worse
```

Do not add code that claims to restore an unknown predecessor. It cannot be
captured, so any such claim would be false.

### The write result

Every write reports through `DP_WriteResult`: the counts, the worksheet-qualified
addresses behind the skipped and failed cells, the resolved target, and the table
expansion metadata the resolver already produces.

```text
M_WriteBack_Apply(iType, [NoTableGrow])          As DP_WriteResult
M_WriteBack_ResolveAndApplyTarget ..., Result    accumulates, attaches metadata
M_WriteBack_ApplyResolvedTarget   ..., Result    accumulates across areas
M_WriteBack_PopulateRange         ..., Result    accumulates one range
```

`M_WriteBack_Apply` is a `Function` because VBA does not permit a user-defined
type as an `Optional` argument, so a `ByRef` output would have to be required and
every caller that ignores it would still have to pass one. Bare-call syntax still
compiles. The private stages below it have no such constraint and take a required
`ByRef` accumulator.

Four rules hold this together:

- **The invariant is load-bearing.** A completed result satisfies
  `AttemptedCount = WrittenCount + LockedSkippedCount + FailedCount`.
  `TST_DP_AssertWriteResultBalances` asserts it on every write path, and is the
  one place that changes when a new skip classification is added.
- **Every exit populates the result.** The fast bulk write returns before the
  per-cell counters exist, so it contributes its own count. A path that returns
  without populating reports a successful write as nothing written.
- **A silent refusal is not a write.** Excel declines some assignments through
  the object model without raising — an array formula is the known case. So
  `WrittenCount` counts cells that reported success; it is never derived by
  subtracting skips and failures from the attempted count, because that treats
  anything that did not raise as written. Anything the engine cannot write has to
  be identified *before* the attempt, never inferred from its outcome.
- **The engine collects facts and displays nothing.** Deciding whether a human is
  told belongs to the entry point that was invoked, which calls
  `M_WriteBack_ReportShortfall` once after the complete result is available.
  `M_WriteBack_ApplyResolvedTarget` calls `M_WriteBack_PopulateRange` once per
  area, so a message raised lower down fires once per area — and a modal dialog
  inside the write path stops the regression harness and forces every
  programmatic caller through UI it did not ask for.
- **Extend the result, do not add a second mechanism.** A new write policy adds
  fields and extends `M_WriteBack_DescribeShortfall`. `EventsDisabledByCaller`
  lives here for the same reason, rather than becoming a second reading of caller
  event state.

> [!NOTE]
> A predicted scope is checked against `AttemptedCount`, never `WrittenCount`.
> `DP_FillTableColumn` predicting 247 cells and writing 244 because three were
> locked is a correct prediction and a legitimate partial write. A prediction
> that does not match `AttemptedCount` means the target changed between preview
> and application, and is a different kind of problem.

> [!IMPORTANT]
> A selected cell inside an Excel Table data column receives the date on its own.
> `DP_FillTableColumn` is the only route to the whole column, and it reports the
> scope before writing.
>
> That default was inverted in `v1.2.0` ([#13](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/13)).
> Any change to write-back must state its effect on it, and the `WriteBack` suite
> covers the three paths — omitted argument, explicit `NoTableGrow:=False`, and
> `DP_FillTableColumn` — independently. Keep all three.

> [!NOTE]
> `#13` changed the *size* of the resolved target. What happens inside it is
> settled separately: literal values are overwritten, formulas are preserved
> unless the caller opts in, and array-formula cells cannot be written at all.
> The three are distinct classifications in `DP_WriteResult`, because a preserved
> formula is a policy decision the caller can reverse and a failure is not.

---

## 🪟 WinAPI changes

- All declarations exist in both the `#If VBA7` and pre-`VBA7` branches.
- `SetLastError 0` before a call whose zero return is ambiguous; read
  `Err.LastDllError` on the statement immediately after.
- `M_Window_GetUserFormHwnd` is the **single** handle resolver. Do not add a
  second `FindWindow` call anywhere.
- `WS_CAPTION` is manipulated in `M_Window_RemoveTitleBar` and nowhere else.
- Test with more than one workbook window open.

> [!NOTE]
> MSForms re-applies `WS_CAPTION` whenever the `Caption` property is assigned.
> Any approach that writes to `UserForm.Caption` will restore the title bar on a
> borderless form. This has been tried and reverted; see
> [#3](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/3).

---

## 🛡️ Error policy

Every routine declares its policy in its banner, and the code must match.

| Policy | Handler label | Use for |
|---|---|---|
| Raises | `ErrorHandler` | Anything a caller must know failed |
| Best-effort | `FailSafe` / `SafeExit` | High-frequency UI, cleanup, teardown, timer callbacks |

Rules that apply to both:

- Capture `Err.Number` and `Err.Description` **before** any `On Error`
  statement. Any `On Error` resets the `Err` object, so a handler that suppresses
  cleanup errors first has already destroyed its own diagnostic.
- Reserve `On Error Resume Next` for genuinely best-effort sequences.
- Anything reachable from an error handler must not itself raise.
- Do not introduce an unsolicited production `MsgBox`.

Note that several module routines end with `On Error GoTo 0`, which clears the
*caller's* handler on return. Where the harness or another caller depends on its
own handler surviving, it re-arms immediately after the call. Preserve those
re-arms; they are not redundant.

---

## ✒️ Source style

New code must match existing conventions. Read a few procedures in
`M_DatePicker.bas` first — they are the reference.

**Module hygiene**

- `Option Explicit` at the top of every module, class and form.
- Do not add `Option Private Module` to `M_DatePicker.bas`; Excel UI callbacks
  depend on public procedures there.
- Keep `cDatePickerManager` and `cDatePickerLabelHook` private-instanced
  (`VB_Exposed = False`).

**Procedure banners**

```vb
'------------------------------------------------------------------------------
'                              PROCEDURE TITLE
'------------------------------------------------------------------------------
' PURPOSE
' WHY THIS EXISTS
' INPUTS
' RETURNS
' BEHAVIOR
' ERROR POLICY
' DEPENDENCIES
' NOTES
' UPDATED
'   YYYY-MM-DD
'------------------------------------------------------------------------------
```

Bump `UPDATED` on every routine you change. A stale date is worse than none —
it asserts something false.

**Body structure**

- Open with `DECLARE`, then sectioned sub-banners (`INITIALIZE`, `VALIDATE`, …).
- Declare `Const PROC_NAME As String = "..."` in routines that report errors.
- Put a short intent comment above each meaningful statement — and only where it
  adds something. `'Exit before the error handler` above `Exit Sub` is noise.

**Naming**

| Prefix | Scope |
|---|---|
| `DP_` | Public entry points |
| `M_` | Module-internal helpers, grouped by area — `M_Settings_`, `M_WriteBack_`, `M_GridIcon_`, `M_Window_`, `M_Timer_` |
| `UF_` | UserForm routines |
| `cDatePicker*` | Classes |
| `Ribbon_` | RibbonX callbacks |
| `TST_DP_` | Test harness |
| `gDP_` | Global runtime state |
| `m` | Private module or instance state |

**Performance**

- Bulk-read and bulk-write ranges; fall back to per-cell only when bulk cannot
  complete.
- Use 1-based arrays.
- Cache repeated lookups rather than re-reading per cell or per render.

---

## 📖 Documentation expectations

Documentation belongs in the same pull request, not a follow-up.

| Change | Update |
|---|---|
| Public API | `README.md` API table and the `Public-API` wiki page |
| Install method | `README.md` and the `Installation-and-Import` wiki page |
| UI internals | The `UserForm-UI-Layer` wiki page |
| Any user-visible change | `CHANGELOG.md` |
| Release process | The `Release-Checklist` wiki page |
| Wiki edits | `UPDATE_NOTES` |

> [!IMPORTANT]
> The wiki was written for `v1.1.0` and has not been fully verified against
> `v1.1.1`. Affected pages carry a notice. Until the rewrite
> ([#17](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/17)) lands,
> `README.md` and the tagged source are authoritative where they disagree.

---

## 🧹 Repository hygiene

Never commit:

- the built `DATETIMEPICKER.xlam` — excluded by `.gitignore`;
- Excel lock and owner files (`~$*`);
- the workbook you used to edit source;
- personal settings, scratch sheets or local paths;
- credentials, client data or production data.

`.gitattributes` stores VBA source as LF in the index and checks it out as CRLF.
Markdown is LF throughout. Do not fight this — export normally and let git
normalize.

---

## 🚀 Pull requests

1. Keep it **small and focused** — one logical change.
2. Fill in the template. Delete sections that do not apply rather than writing
   "N/A" fifteen times.
3. State the problem, the approach, and how you tested it.
4. Quote the harness summary line.

The maintainer reviews selectively and may adopt, adapt or decline a
contribution to keep the codebase coherent. Clear, well-scoped pull requests are
the most likely to be merged.

---

## 📄 License

By contributing, you agree that your contributions are licensed under the
project's **MIT License**.

## 👤 Maintainer

Maintained by **Daniele Penza**. For anything that is not a code change — design
questions, larger proposals, general feedback — open an issue to start the
conversation.
