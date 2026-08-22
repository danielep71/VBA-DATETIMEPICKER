<div align="center">

# 📄 Changelog

**All notable changes to VBA-DATETIMEPICKER**

[![Semantic Versioning](https://img.shields.io/badge/Versioning-semver-6f42c1?style=flat-square)](https://semver.org/)
[![Format](https://img.shields.io/badge/Format-Keep_a_changelog-0969da?style=flat-square)](https://keepachangelog.com/)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#)

</div>

---

Versioning applies to the **public VBA API** — every `DP_…` procedure, the
`M_Settings_…` getters and setters, the `DP_WriteAction`, `DP_ClockMode` and
`DP_SizeMode` enums, and the two stable legacy identifiers `VBA_DATETIMEPICKER`
used as the registry application name and the command-bar tag.

Internal module boundaries are not covered by it. A release that changes nothing
public may still require every component in `src/` to be re-imported together.

<details>
<summary><strong>Section legend</strong></summary>

<br>

| Section | Contains |
|---|---|
| ➕ **Added** | New members, suites, workflows or files |
| 🔧 **Changed** | Behavior or contract changes to something that already existed |
| 🐛 **Fixed** | Defects, each citing its issue |
| 📖 **Documentation** | Corrections and additions to prose, with no code effect |
| ✅ **Verified — no action required** | Reported defects that did not reproduce, with the evidence |
| 🧪 **Validation** | The evidence the release was actually tested on |
| 🔗 **Compatibility** | What upgrading requires, and what becomes newly observable |
| ⚠️ **Known limitations** | What is deliberately not fixed, and where it is tracked |

Release types follow semver: 🩹 **patch** corrects defects, ✨ **minor** adds
backward-compatible capability, 💥 **major** may break callers.

</details>

---

## [Unreleased]

> Targeting `v1.2.0`

### ➕ Added

- Added run states to the regression harness. A run now reports one of
  `PASS`, `FAIL`, `FAIL_CLEANUP` or `INCOMPLETE_SKIPPED`, in the Immediate
  Window summary and on the result sheet:

  ```text
  INFO | Harness | Summary | State=PASS; Run=150; Passed=150; Failed=0; CleanupFailures=0
  ```

  Assertion totals alone cannot express a run that passed every assertion but
  failed to clean up, or one that ended before every dispatched suite returned.
  Both looked like a pass when only `Passed` and `Failed` were reported.

- Added `TST_DP_VerifyFinalState`, which checks after teardown that the picker
  form is not still loaded, no `DP_GridIcon` shapes remain in the host workbook,
  and `Application.EnableEvents` matches the pre-run snapshot. It reports rather
  than repairs: silently fixing a leak would hide the defect that caused it.

- Added a dirty-start guard. A run that aborts before teardown now leaves a flag
  that the next run reports, so the cause is visible at the point of the defect
  rather than as an unexplained setup failure later.

- Added `demo/M_DP_DEMO.bas`, which builds the demo worksheet from code. The
  demo previously existed only as content inside `demo/DATEPICKER.xlsm`:
  `demo/M_DEMO_BUILDER.bas` is a toolkit of primitives, and nothing in the
  repository called them to produce the DatePicker demo. A change to the demo
  therefore appeared in a pull request as a binary diff and nothing else.

  Two entry points — `DP_Demo_CreateDemoSheet` builds or rebuilds,
  `DP_Demo_EnsureDemoSheet` returns the sheet and builds only when it is
  missing. Both take an optional target workbook and default to
  `ActiveWorkbook`.

  The Excel Table section is a real `ListObject` rather than a formatted range.
  A formatted range would not exercise the table write-back path, which is the
  reason that section exists. Its Expiry Date column is deliberately empty:
  selecting one of those cells is the shortest route to observing the
  write-scope behaviour described in
  [#13](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/13).

### 🐛 Fixed

- Fixed the regression harness discarding cleanup failures. Teardown runs under
  `On Error Resume Next` so that one failing step does not prevent the rest from
  being attempted, and nothing checked the outcome afterwards. A failure to
  restore `Application` state, delete the scratch worksheet or release the
  manager was silently dropped, and the run still reported its assertion totals
  as if nothing had gone wrong.

  Because the harness mutates process-wide Excel state — it sets
  `ScreenUpdating`, `EnableEvents` and `DisplayAlerts` to `False` for the
  duration — an aborted run left Excel in that state and left its worksheets
  behind. The next run then failed during setup:

  ```text
  FAIL | Harness | TST_DP_RunAllInternal | Fatal error 1004 -
        Method 'Add' of object 'Sheets' failed
  ```

  Each of the five teardown steps is now checked immediately after it runs.
  A failure is recorded as a `FAIL` row against the `Cleanup` suite, counted,
  and the first detail carried into the summary.
  ([#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19))

- Fixed `TST_DP_RunSuite_Manager` and `TST_DP_RunSuite_UISmoke` destroying their
  own diagnostics. Both performed cleanup before recording the escaping error,
  and any `On Error` statement resets the `Err` object — not only `Err.Clear` —
  so both reported `Error 0 -` with no number and no description.

  The reported scope was wider than the defect. Reading all thirteen handlers
  shows eleven read `Err` as the first statement and report correctly, and
  `TST_DP_RunSuite_WriteBack` releases object references first, which is
  assignment rather than an `On Error` statement, so `Err` survives. The
  `Error 0 -` originally observed came from `ApplicationState`, which was
  corrected when it was written.
  ([#18](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/18))

- Fixed the demo-sheet routines resolving their worksheet from `ThisWorkbook`.
  Four routines did so — `Ribbon_Demo`, `DP_DemoSheet_Show`,
  `DP_DemoSheet_HideVeryHidden` and `DP_DemoSheet_GetSafeVisibleSheet` — and
  their headers recorded the choice as deliberate, which it was for an embedded
  copy. In the `.xlam` it is wrong: `ThisWorkbook` is the add-in, which has no
  worksheets, so the Ribbon **Demo** button raised subscript-out-of-range on
  every click.

  A new private resolver decides the host workbook once, so the four routines
  cannot disagree about which workbook they are operating on. Embedded, it
  returns `ThisWorkbook` and behaviour is unchanged. As an add-in, it returns
  the first open workbook already holding the demo sheet, and adds a new
  workbook when none does.

  The add-in deliberately does not build into whichever workbook happens to be
  active. Adding an unrequested sheet to a user's live workbook is a worse
  outcome than opening a new one.
  ([#23](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/23))

### 🔗 Compatibility

- No public procedure was added, removed or renamed in `src/`.
- `DP_Demo_CreateDemoSheet` and `DP_Demo_EnsureDemoSheet` are new, and live in
  `demo/`. Versioning covers the public VBA API of the component in `src/`;
  `demo/` is example material that ships as a workbook.
- Embedded behaviour is unchanged. The add-in gains a Demo button that works.
- The harness summary line gained `State=` and `CleanupFailures=` fields. Any
  tooling that parses it by position rather than by name will need updating.

### ⚠️ Note on the add-in

Building the demo from the add-in requires `M_DP_DEMO.bas` and
`M_DEMO_BUILDER.bas` to be present in the `.xlam`, which adds roughly 6,500
lines of demo-construction code to a component whose purpose is a date picker.
That is a deliberate trade for a working Demo button; the alternative was
hiding the button when running as an add-in.

---

## [1.1.1] - 2026-08-22

> 🩹 **Patch** · correctness and disclosure release · public API unchanged

### 🧭 Release intent

A corrective release following an internal code and repository assessment of
`v1.1.0`.

The defects fixed here share a shape. A bootstrapper that re-enabled Excel
events a caller had deliberately suppressed; two copies of a window lookup that
had quietly diverged; five WinAPI call sites reading an error slot through a
path the language does not guarantee. Each one worked in the common case and
failed in the case nobody had reached yet, and none of them announced anything.

Two defects the assessment reported did **not** reproduce and were closed on
evidence rather than remediation. Two more were found while fixing the ones that
did.

The two most serious findings — the table write-scope default and the absence of
a runtime ownership model — are **not** fixed here. They need design work that
does not belong in a patch. This release documents both instead, in the README
and below, so that shipping with them open is a stated decision rather than an
omission.

#### At a glance

| Issue | What it was |
|:--:|---|
| [#2](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/2) | 🔴 `M_Picker_EnsureManager` re-enabled Excel events underneath the caller |
| [#3](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/3) | 🟠 Two caption-based window lookups had diverged |
| [#4](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/4) | 🟡 WinAPI error capture did not use `Err.LastDllError` |
| [#6](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/6) | 🟡 README named files that do not exist |
| [#7](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/7) | 🟡 README claims exceeded what the release demonstrates |
| [#8](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/8) | 🟠 Governance documents contradicted the tracked binary policy |
| [#11](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/11) | 🟡 Traffic workflow used a movable action tag and an unscoped token |
| [#1](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/1) | ✅ Did not reproduce — `.gitattributes` adoption was already complete |
| [#5](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/5) | ✅ Did not reproduce — the wiki was written for `v1.1.0`, not before it |

🔴 P1 · 🟠 P2 · 🟡 P3 · ✅ verified, no action required.

---

### 🐛 Fixed

- Fixed `M_Picker_EnsureManager` forcing `Application.EnableEvents = True` on
  every call. `DP_Start`, `DP_Show` and `DP_Preload` all route through it, so a
  business macro that had deliberately suppressed events had them re-enabled
  underneath it — silently, mid-transaction. The bootstrapper now observes the
  caller's state and reports it through a new optional
  `ByRef EventsDisabledByCaller As Boolean` output, and `DP_RepairRuntime` is
  the only routine in the component that writes to `Application.EnableEvents`.

  A manager hooked while events are disabled is a valid, self-correcting state:
  the `WithEvents` reference stays live and Excel resumes dispatching as soon as
  the caller restores events. `Is_Hooked` derives from the manager's own hook
  flag and its stored `Application` reference and does not consult
  `EnableEvents`, so a suppressed-event session does not cause repeated manager
  recreation. ([#2](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/2))

- Fixed `M_Window_BeginUserFormDrag` resolving its own window handle instead of
  calling `M_Window_GetUserFormHwnd`. Two copies of the same `FindWindow`
  fallback existed, and they had already diverged — the resolver exits early on
  a blank caption, the drag path did not and would call `FindWindow` with an
  empty string. The drag path now delegates, `FindWindow` call sites drop from
  four to two, and the caption policy and blank-caption guard apply to every
  WinAPI-dependent routine.
  ([#3](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/3))

- Fixed WinAPI error capture calling `GetLastError` directly at five sites. VBA
  populates `Err.LastDllError` from the same thread-local slot and guarantees it
  is read against the failing `Declare`d call; reaching the slot through another
  declared call does not carry that guarantee. All five converted, both
  `GetLastError` declarations removed from the `VBA7` and pre-`VBA7` branches,
  and the four `SetLastError 0` pre-clears preserved — they are what makes a
  zero return distinguishable from a genuine failure where the API's return
  value is ambiguous.
  ([#4](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/4))

- Fixed the traffic export workflow reporting success while failing. The step
  that opens an alert issue passed `check=False` to both its `gh` invocations
  and never inspected the return code, so a denied API call was discarded and
  the job exited zero. Both now capture stderr, emit `::error::` and exit
  non-zero. The duplicate-alert probe was the more dangerous of the two: a
  failure read as "no existing alert", which would have created a duplicate
  issue every day once the permission was restored.
  ([#11](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/11))

- Fixed the README naming files that do not exist. The repository structure and
  installation steps referenced `src/ribbon/customUI.xml` and
  `demo/DatePicker Demo.xlsm`; the tracked files are `customUI14.xml` and
  `DATEPICKER.xlsm`.
  ([#6](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/6))

- Fixed `CONTRIBUTING.md` and the pull request template asserting that the
  repository stores exported source "not a binary workbook" while
  `demo/DATEPICKER.xlsm` is tracked and `.gitignore` excludes only
  `/dist/*.xlam`. Both now carry a table distinguishing the tracked demo
  workbook, the untracked add-in, and `UF_DatePicker.frx` — which is form source
  rather than a build artifact and was covered by neither policy.
  ([#8](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/8))

### ➕ Added

- Added `TST_DP_RunSuite_ApplicationState` to the regression harness, wired into
  both `TST_DP_RunAllInternal` and the `TST_DP_RunSuiteSafe` dispatcher so it
  runs unconditionally rather than behind the UI-smoke flag. Nine assertions:
  bootstrap, `DP_Start`, `DP_Show` and `DP_Preload` each preserve a disabled
  event state; `EventsDisabledByCaller` reports it; the manager still hooks and
  `Is_Hooked` stays `True` while events are suppressed; write-back restores both
  the disabled and the enabled case; and `DP_RepairRuntime` still force-enables.

- Added a **Known limitations** section to `README.md` covering the two P1
  defects that ship open in this release, and the operating conditions under
  which the component is supported.

- Added a **Repository automation credentials** section to `SECURITY.md`: the
  control table for `TRAFFIC_TOKEN`, a 90-day rotation rule with its triggers,
  the rationale for isolating an `Administration: read` token, and what the
  public `traffic-history` branch does and does not contain.

### 🔧 Changed

- `TST_DP_RunSuite_LifecyclePair` asserted that `Application.EnableEvents` is
  `True` after `DP_Start`. That test encoded the defect and only ever passed
  because the bootstrapper forced events on. It now captures the caller's state
  before `DP_Start` and asserts it is unchanged. Its failure on the first run
  after the fix was the confirmation that the fix worked.

- The traffic export workflow is scoped and pinned. `actions/checkout` is pinned
  to the full commit SHA that `@v6` resolved to, so a moved tag cannot change
  what runs; `TRAFFIC_TOKEN` moved to a dedicated `analytics` environment and
  was removed from repository scope; and `traffic-history` is protected against
  deletion and force-push. The token was regenerated during this work — it had
  already expired, so the export had been failing before this release began.
  ([#11](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/11))

- `README.md` claims now match what the release demonstrates. The headline read
  as though no add-in exists while the repository distributes one;
  "Enterprise-friendly" was replaced with "Manager-driven" and
  "Deployment-friendly", neither of which asserts anything untested; and the
  table-column feature bullet carries an explicit warning about its blast
  radius. The Roadmap section was removed — milestones are public and current, a
  hand-maintained list is neither.
  ([#7](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/7))

### 📖 Documentation

- `CONTRIBUTING.md` and the pull request template were rewritten. Beyond the
  binary-policy correction, they now carry what this release established:
  `Application.EnableEvents` is caller-owned; write scope requires explicit
  intent; `M_Window_GetUserFormHwnd` is the single handle resolver and
  `WS_CAPTION` is manipulated in exactly one routine; `Err` must be captured
  before any `On Error` statement; and the harness recovery procedure for an
  aborted run.

  The template gained the three sections it lacked — application state, write
  scope, and a validation environment record.

- The wiki carries a notice on `Home`, `Public-API` and `Manager-and-Events`
  covering the `EnableEvents` change and the corrected file paths. A full
  rewrite against `v1.1.1` is deferred; until it lands, `README.md` and the
  tagged source are authoritative where they disagree.
  ([#17](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/17))

### ✅ Verified — no action required

Two reported defects did not reproduce. Both are recorded rather than dropped: a
release that silently omits findings which turned out to be wrong is less
trustworthy than one that says what was checked.

- **`.gitattributes` adoption was already complete.** The report described it as
  incomplete and recommended a renormalization commit.
  `git add --renormalize .` reported a clean tree, `git ls-files --eol` showed
  every tracked text file stored LF in the index, and `git check-attr` confirmed
  `UF_DatePicker.frx` resolves to `binary` / `merge: binary` and is excluded from
  normalization — which it must be, since a normalized `.frx` corrupts the form.
  ([#1](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/1))

- **The wiki was written for `v1.1.0`, not before it.** The report described
  every page as documenting a superseded architecture, and named five specific
  contradictions. Four did not reproduce: the pages list the current component
  names, reference no BMP payload, call the current startup procedures including
  the `v1.1.0` additions `DP_Preload` and `DP_Hide`, and document
  `Application.OnTime` rather than `SetTimer`. The fifth — Ribbon toggles absent
  from the shipped XML — also did not reproduce; `customUI14.xml` exposes exactly
  the three callbacks the wiki documents.

  The wiki is 21 files, not the eight reported, and `UPDATE_NOTES` records it as
  updated *for* `v1.1.0`.
  ([#5](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/5))

- **The governance documents referenced wiki pages that all exist.** The report
  listed `Public-API`, `Installation-and-Import`, `UserForm-UI-Layer`,
  `Release-Checklist` and `UPDATE_NOTES` as nonexistent. All five are present.
  ([#8](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/8))

### 🧪 Validation

Validated manually in desktop Microsoft Excel for Windows.

```text
Debug → Compile VBAProject               clean
TST_DP_RunAll                            150 assertions, 150 passed, 0 failed
Manual  Application.EnableEvents = False → DP_Show → still False
Manual  DP_Show / DP_Close / drag / WinAPI styling toggled off and on
```

Two defects in the harness itself were found during this release and are
**not** fixed here:

- Suite-level failures report `Error 0 -`, because the handler executes
  `On Error Resume Next` before reading `Err`, and any `On Error` statement
  resets the `Err` object. Every suite except the new one is affected.
  ([#18](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/18))
- An aborted run leaves `Application` state and worksheets behind, and the next
  run then fails during setup with
  `1004 — Method 'Add' of object 'Sheets' failed`. Observed during this release.
  ([#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19))

There is no workflow that compiles or tests this component. A hosted runner has
no Excel, so validation remains a manual step on a real host; automating the
parts that can run without Excel is tracked in
[#15](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/15).

### 🔗 Compatibility

| Question | Answer |
|---|---|
| Existing calls affected | ✅ none |
| Backward compatible | ✅ yes |
| Release type | 🩹 patch |
| Components to replace | ⚠️ `M_DatePicker.bas` and `M_cDP_Test.bas` |

- No public procedure was added, removed or renamed.
- No existing parameter changed name, position, type or default.
- No enum member or value changed.
- `M_Picker_EnsureManager` gained an **optional** `ByRef` parameter. All four
  call sites compile unchanged.
- Error `vbObjectError + 512` is retired. Any external code testing for that
  specific number will no longer see it; the condition it guarded — Excel events
  could not be re-enabled — is no longer treated as an error. Errors `513` and
  `514` are unchanged.
- **One behavior is newly observable.** A caller that relied on `DP_Show`,
  `DP_Start` or `DP_Preload` to clear a stuck `Application.EnableEvents = False`
  no longer gets that side effect. This is intentional, and `DP_RepairRuntime`
  is the replacement.

### ⚠️ Known limitations

Two P1 defects ship open. Both are safe to work with once you know they exist,
and both are documented in the README.

- **A selected cell inside an Excel Table data column can write the entire data
  column.** This applies to calendar selection, `DP_Today` and `DP_Now`, and it
  overwrites existing values and formulas in that column. Keep version history
  or backups for workbooks that use table write-back. Deferred to `v1.2.0`,
  where the fix adds an explicit fill command, a resolved-target preview, and a
  structured write result.
  ([#13](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/13))

- **The runtime has no ownership model.** The component registers process-wide
  Excel surfaces — the keyboard shortcut, the context-menu entry, `Application`
  events, the live-clock timer, worksheet icons and registry settings — without
  one. Two copies in a single Excel session interfere: either can remove the
  other's registrations and delete the other's worksheet icons. Do not run the
  embedded source and the `.xlam` in the same session, and do not embed the
  picker into multiple workbooks that will be open at once. `v1.2.0` adds
  duplicate detection; the full broker is `v1.3.0`.
  ([#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37),
  [#14](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/14))

Also open, and lower impact:

- Window handle resolution matches on caption, which is not a durable window
  identity. This only becomes ambiguous when two DatePicker forms are loaded at
  once — the scenario above. A caption-token approach was implemented and
  reverted: MSForms re-applies `WS_CAPTION` on every `Caption` assignment, so it
  restored the title bar on a borderless form.
  ([#14](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/14))
- A window style transaction that fails after committing the style write leaves
  the form half-applied, and a later call cannot repair it.
  ([#20](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/20))
- Accessibility, high-DPI and high-contrast behavior are neither tested nor
  documented.
  ([#29](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/29))

---

## [1.1.0] - 2026-06-17

> ✨ **Minor** · add-in distribution and lifecycle API

Backfilled from the repository history. This entry is less detailed than those
above it because it was reconstructed rather than written at release time.

### ➕ Added

- Prebuilt `DATETIMEPICKER.xlam` add-in, published as a GitHub Release asset for
  users who want the picker across every workbook without importing source.
- `DP_Preload` — loads and hides `UF_DatePicker` at startup so the first real
  open is instant. Best-effort; a failed preload never blocks normal use.
- `DP_Hide` — hides the form instead of unloading it, keeping it warm for fast
  reuse. Pairs with `DP_Preload`.
- `UF_DayGrid_BuildIndexMap` — a single case-insensitive dictionary mapping both
  the background-label and text-label names of each day cell to the same numeric
  index, replacing a per-render date cache that was cleared and rebuilt.
- `UF_DayCell_ApplyDateStateFast` — paints one day cell from already-resolved
  label references, avoiding repeated cache probing during a full grid refresh.
- Daily traffic export workflow, writing a CSV history to the orphan
  `traffic-history` branch.
- `dist/README.md` pointing users to the Releases page.

### 🔧 Changed

- The regression harness moved from `src/classes/` to `test/`, where it belongs.
- Weekend-highlight and other render settings are cached for the duration of a
  full grid refresh rather than re-read per cell.

### 🧪 Validation

Validated manually in desktop Microsoft Excel. No machine-readable evidence was
produced for this release.

---

## [1.0.0] - 2026-05-16

> 🎉 **Initial release**

Backfilled from the repository history.

### ➕ Added

- Modeless Date / Time Picker UserForm, built at runtime.
- `cDatePickerManager` — application-level event manager coordinating worksheet
  selection, workbook lifecycle and UI refresh.
- `cDatePickerLabelHook` — `WithEvents` router giving runtime-created labels
  click, hover and optional drag events.
- Public API: `DP_Start`, `DP_Show`, `DP_Click`, `DP_Close`, `DP_Today`,
  `DP_Now`, `DP_RepairRuntime`.
- Registry-backed settings under the `VBA_DATETIMEPICKER` application name, with
  an in-form settings panel across Display, Behavior and Integration pages.
- Entry points: Ribbon callbacks, right-click menu, in-grid worksheet icon, and
  the `Ctrl + Shift + D` keyboard shortcut.
- Optional WinAPI styling — borderless form, mouse positioning, and drag from a
  passive header surface.
- Regression harness with per-area suites.
- Demo workbook and builder.

---

[Unreleased]: https://github.com/danielep71/VBA-DATETIMEPICKER/compare/v1.1.1...HEAD
[1.1.1]: https://github.com/danielep71/VBA-DATETIMEPICKER/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/danielep71/VBA-DATETIMEPICKER/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/danielep71/VBA-DATETIMEPICKER/releases/tag/v1.0.0
