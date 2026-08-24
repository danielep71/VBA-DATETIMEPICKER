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

- Added a one-provider runtime lease, so two DatePicker copies in one Excel
  session are detected and the second is refused rather than silently displacing
  the first.

  Every copy registers the same application-wide resources under the same fixed
  identifiers — the `Ctrl + Shift + D` binding, the context-menu tag, the grid
  icon name — and either copy's teardown removes them all. Nothing detected this.

  The lease is a hidden `Temporary` `CommandBar` carrying one hidden `Temporary`
  control whose `Parameter` holds an ephemeral owner token:

  ```text
  __VBA_DATETIMEPICKER_RUNTIME_PROVIDER_LEASE__
  ```

  A `CommandBar` is visible to every VBA project in the process, needs no WinAPI
  — which the lease could not use, since WinAPI is disableable by setting and by
  platform — and `Temporary:=True` means Excel removes it at shutdown. A
  registry-backed lease would have had the opposite lifetime: it would survive a
  restart and block startup permanently.

  `DP_Start` claims the lease **before** the first shared registration. A copy
  that registered first and discovered the conflict afterwards would already have
  displaced the owner's shortcut.

  `DP_Stop` and `DP_RepairRuntime` verify ownership before acting. That is the
  more important half: refusing a second copy at startup protects nothing while
  its teardown still dismantles the owner. `DP_Stop` releases the lease last,
  after removing its own registrations.

  Ownership is never assumed. Release requires a local token, a lease still
  carrying it, and an exact match; an unreadable lease is reported as ambiguous
  and left alone, because a marker this component cannot interpret belongs to
  something. Acquisition re-reads the marker after writing it and only then
  claims ownership.

  Stale leases fail closed. A VBA project reset destroys the owner token while
  the lease survives, so the former owner can no longer prove ownership and every
  guarded entry point refuses. Restarting Excel clears it, since the lease is
  temporary. Automatic reclamation of a stale lease belongs to
  [#14](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/14).

  `DP_ForceReleaseProviderLease` is the one deliberate exception — an operator
  command that deletes the lease regardless of ownership, for the case where the
  project was reset and no other copy is running. It is never called
  automatically, and the refusal message names it.

  The lease protects `v1.2.0` against `v1.2.0` only. A copy from an earlier
  release has no lease code, registers unconditionally, and never checks whether
  another copy owns the session — so a mixed-version session is unprotected in
  both directions. A released build cannot be given a check it never had, so this
  is documented in README **Known limitations** rather than worked around.
  ([#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37))

- Added `DP_WindowStyleResult`, the structured outcome of applying the borderless
  window style. `M_Window_RemoveTitleBar` returned nothing, so complete success,
  a safe abort before any change, and a half-applied style were indistinguishable
  at the two `UF_DatePicker.frm` call sites.

  ```vb
  Public Type DP_WindowStyleResult
      Attempted           As Boolean
      Applied             As Boolean
      Committed           As Boolean
      RolledBack          As Boolean
      RecoveryRequired    As Boolean
      FailedStep          As String
      LastApiError        As Long
  End Type
  ```

  The routine is now a `Function`, following the precedent
  [#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21) set for
  operations with meaningful partial outcomes. It is a separate type from
  `DP_WriteResult` because a native-window transaction and a worksheet write-back
  are different domains. Bare-call syntax still compiles, so both call sites are
  unchanged.
  ([#20](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/20))

- Added the `WindowStyle` regression suite, covering every native failure point
  the transaction has to survive: safe non-attempts, the successful path and its
  repeat, both pre-commit failures, both post-commit failures with successful
  rollback, both rollback failures, and one-shot fault consumption for the
  primary and rollback seams.

  None of those paths can be produced through ordinary input — no test can make
  `SetWindowPos` fail on a window that has just accepted a style write — so the
  suite drives them through a narrow test-only seam in `M_DatePicker`:

  ```text
  two Private injection values
  one Public setter, M_Window_Test_SetFaultInjection, requiring an argument
  consumed one-shot at entry, before the window is touched
  ```

  The setter is technically `Public` only because the regression module is a
  separate VBA module and cannot assign `Private` state in `M_DatePicker`. It is
  internal test infrastructure, classified `internal` under
  [#25](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/25), absent from
  the README supported API, and has no effect unless deliberately armed. Failure
  points are private constants duplicated in both modules rather than a public
  enum, so the seam does not enlarge the public surface.

  An injected failure **skips** the native call it is failing. Performing the
  call and then overwriting its result would leave the window in the state of a
  success while the result described a failure — the opposite of what these paths
  exist to reproduce. A skipped style write leaves the window with its original
  style; a skipped rollback restore leaves it genuinely committed, which is what
  makes `RecoveryRequired` true rather than merely reported.

  One-shot consumption is the other safety property: the injected values are
  copied to locals and cleared before any native call, so a test run that aborts
  mid-way cannot leave a later real call poisoned. The harness disarms the seam
  again on its own cleanup and failure paths.

  The suite verifies native window state directly, through its own
  `GetWindowLongPtr` declaration, rather than inferring it from the result the
  routine under test returned. That covers the claims most worth proving
  independently: a pre-commit failure leaves the style unchanged, a successful
  rollback restores it, and a failed rollback does not.

  The suite asserts a resolvable window handle before exercising any injected
  failure. A preloaded hidden UserForm has one on supported hosts, which is why
  these cases run in the standard pack; a missing handle reports as a setup
  failure rather than passing quietly as a non-attempt.
  ([#20](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/20))

- Added `FAIL_DIRTY_START`, a fifth harness run state, and the preflight that
  produces it. The harness detected an incomplete predecessor and recorded an
  `INFO` row about it, but the detection never reached
  `TST_DP_ResolveRunState`. A run could inherit a dirty environment, pass every
  assertion against it, and report `PASS`:

  ```text
  dirty predecessor detected
      + current assertions pass
      + current cleanup succeeds
      → State=PASS
  ```

  The new state outranks every other outcome, including assertion failures. A
  dirty start does not make the report bad, it makes it untrustworthy — the
  failures may belong to the environment the previous run left rather than to the
  code under test, and reporting `FAIL` first would send someone to debug an
  artifact.

  Detection also moved to the real entry boundary. It ran inside
  `TST_DP_RunAllInternal`, after the public entry points had already resolved the
  host workbook and built the result sheet template — and building that template
  over a previous run's leftovers is one of the ways the failure presents.
  `TST_DP_Preflight` now runs before any mutation, reads without writing, and
  fails closed: an environment it cannot inspect is reported dirty rather than
  assumed clean.

  Detection takes two independent kinds of evidence, because they survive
  different kinds of abort:

  ```text
  mTST_DP_RunInProgress   an abort that left module state intact
  leftover scratch sheet  an abort that cleared it, such as a project reset
  ```

  The module flag alone could not see the abort it existed to detect. A VBA
  project reset zeroes module state, so the next run read `False` and reported a
  clean start while sitting in the previous run's wreckage. The worksheet is the
  evidence that survives.

  A dirty run does not execute suites. It records why, tears down, and reports.
  Gathering results in an environment the run did not establish would describe
  the predecessor's leftovers rather than the code under test, and a harness
  whose subject is evidence quality should not manufacture evidence it cannot
  interpret.

  The verdict is acted on before the workbook is touched, not after. Preflight
  ran before the mutation but its result was only consumed further downstream, so
  a dirty start was detected and then crashed on the very mutation the detection
  existed to prevent. Both entry points now refuse before building anything,
  reporting through the Immediate window and a message box — the result sheet
  cannot be created by a run that has just declined to touch the workbook.
  ([#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19))

- Added setup diagnostics to the regression harness. A failure before the run
  started surfaced as an unhandled `1004` naming a method and nothing else — no
  result sheet exists yet, and the run's own fatal handler had not been reached.

  ```text
  TST_DP_ReportSetupFailure    entry-point handler; names the error and its cause
  TST_DP_DescribeHostWorkbook  IsAddin, ProtectStructure, ReadOnly, sheet counts,
                               and whether each harness worksheet exists
  TST_DP_ReportEnvironment     public probe, runnable without starting a run
  ```

  `TST_DP_RunAllInternal` also tracks its setup steps, so a fatal names which of
  the six it happened in rather than only what failed:

  ```text
  Fatal error 1004 - Method 'Add' of object 'Sheets' failed
      | Step=Prepare scratch sheet
      | Host=...; Worksheets=3; TST_DP_SCRATCH exists=False
  ```
  ([#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19))

- Added the `HarnessSelfCheck` suite, which proves the run-state machine instead
  of assuming it. It drives `TST_DP_ResolveRunState` across all five outcomes,
  forces a cleanup step to fail through a one-condition injection seam in
  `TST_DP_CheckCleanupStep`, and runs the preflight against the live environment.

  `FAIL_CLEANUP` and `FAIL_DIRTY_START` are states a passing run never reaches on
  its own, and
  [#15](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/15) gates a
  release on them. A state that has never been observed to occur is not evidence
  of anything.

  Every counter the suite manipulates belongs to the run it is part of, so each
  is saved before the first mutation and restored before the first assertion.
  Assertions run against locals captured during the probe, never against live
  module state. A staged cleanup failure records as `INFO` rather than `FAIL`,
  because it is a probe of the routine and not a defect in the run.
  ([#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19))

- Added `DP_WriteResult`, the structured outcome of a write-back. Every write
  routine returned nothing, so a caller could not determine how many cells were
  targeted, how many were written, whether anything was skipped or failed, which
  cells those were, what range the selection resolved to, whether a one-cell
  table selection expanded to a data column, or whether Excel events were already
  disabled on entry. The only external signal was an exception, which made
  *nothing happened* and *some of it happened* indistinguishable.

  ```vb
  Public Type DP_WriteResult
      AttemptedCount          As Double
      WrittenCount            As Double
      LockedSkippedCount      As Double
      LockedSkippedAddresses  As String
      FailedCount             As Double
      FailedAddresses         As String
      ResolvedTargetAddress   As String
      ExpandedToTableColumn   As Boolean
      TableName               As String
      ColumnName              As String
      AreasCount              As Long
      EventsDisabledByCaller  As Boolean
  End Type
  ```

  A completed result satisfies:

  ```text
  AttemptedCount = WrittenCount + LockedSkippedCount + FailedCount
  ```

  Cell counts are `Double` rather than `Long` because they derive from
  `Range.Cells.CountLarge`, which a full-worksheet target overflows a `Long`.
  The skip and failure counts share the type so the invariant cannot overflow on
  one side.

  The result consumes what the resolver split
  ([#13](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/13)) already
  produced rather than re-deriving it: `M_WriteBack_ResolveAndApplyTarget`
  attaches the expansion metadata the resolver returns, and the resolver
  signature is unchanged.

  `EventsDisabledByCaller` records the same operational fact
  `M_Picker_EnsureManager` already surfaces
  ([#2](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/2)), carried in
  this result rather than as a second interpretation of caller event state.
  `M_WriteBack_Apply` derives it from the state it already captures, and still
  restores exactly the state it observed on entry.
  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21))

- Added address reporting for every cell a write did not write.
  `LockedSkippedAddresses` and `FailedAddresses` carry worksheet-qualified
  addresses in the form `SheetName!A1`, so a caller can distinguish *three
  failures somewhere in the target* from `Scratch!G6, Scratch!G9, Scratch!G12`.

  The lists are capped at 25 entries and end with an ellipsis beyond that. The
  counts stay exact — the cap bounds the reported string so a failed write across
  a long table column cannot build an unbounded string inside the write loop.
  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21))

- Added `M_WriteBack_ReportShortfall` and `M_WriteBack_DescribeShortfall`.
  `M_WriteBack_ReportShortfall` shows one consolidated summary for a whole
  operation and is called by the interactive entry points;
  `M_WriteBack_DescribeShortfall` formats the counts and addresses behind it. The
  formula and existing-value skips of
  [#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22) extend these
  two rather than adding a parallel report.

- Added structured-result coverage to the `WriteBack` suite, spanning the
  regressions the result has to survive: contiguous bulk write, single-cell
  write, discontiguous multi-area write, partial protected-range write, exact
  failed-address reporting, and caller event state both enabled and disabled.

  `TST_DP_AssertWriteResultBalances` asserts the accounting invariant on every
  path, and is the single place that changes when `#22` adds
  `FormulaSkippedCount` to the right-hand side.

  `TST_DP_ExpectPartialWriteReport` protects the scratch sheet with one locked
  cell inside the target and asserts two written, one skipped, and
  `Scratch!H6` as the address behind it. `TST_DP_ExpectFailedAddressReport`
  covers the other classification: an array formula inside the target makes two
  cells reject the write, and the result must name exactly those two. It asserts
  its own setup took effect and that the array survived the write, so a silent
  setup failure cannot be mistaken for a reporting defect. Both release what they
  changed on every path, including a failed assertion, so neither sheet protection
  nor an array formula can leak into a later suite.

  The `ApplicationState` suite now asserts `EventsDisabledByCaller` on both
  branches, against the same call whose `Application.EnableEvents` restoration it
  already checked.

- Added three-way write-scope coverage to the `WriteBack` suite. The existing
  expansion test passes `NoTableGrow:=False` explicitly, so it kept passing
  through the behaviour change and proved nothing about the new default. Eleven
  assertions now cover each path independently:

  ```text
  omitted argument          anchored cell written, rest of column untouched
  NoTableGrow:=False        whole data column written, unchanged legacy path
  DP_FillTableColumn        whole data column written, deliberately
  ```

  Plus the anchors `DP_FillTableColumn` must refuse: a cell outside any table, a
  header cell, a totals-row cell, and a multi-cell selection. Each asserts that
  nothing was written rather than only that no error was raised.

  The legacy expansion test is retained rather than replaced. Proving the defect
  is fixed and proving the intentional bulk capability still exists are separate
  claims.

- Added a **Fill Table Column** button to the demo sheet, and
  `DP_Demo_FillTableColumn` behind it. `DP_FillTableColumn` takes the date to
  write, so a worksheet button cannot call it directly; the wrapper supplies
  today's date and leaves confirmation on.

  The demo's Expiry Date column now demonstrates both scopes side by side:
  picking a date in one of its cells writes that cell only, while the button
  fills the column after reporting how many cells it would affect.

- Added `DP_FillTableColumn`, the explicit way to write one date to every cell
  of an Excel Table data column:

  ```vb
  DP_FillTableColumn ValueToWrite As Date, Optional ConfirmFill As Boolean = True
  ```

  Filling a table column is a legitimate operation. It used to happen
  implicitly — selecting one cell inside a table and picking a date wrote the
  whole column with nothing to indicate the scope. Now the scope comes from the
  command the user invoked.

  With `ConfirmFill` left at its default the routine describes the resolved
  scope before writing anything:

  ```text
  Fill 247 cells in Trades[Expiry Date] with 25-Aug-2026?
  ```

  A selection outside a table data body is an ordinary usage condition rather
  than a failure: the routine reports what is required and exits cleanly. Header
  cells and totals-row cells are deliberately rejected, matching the rule the
  write engine already used — the anchor must be inside `DataBodyRange`.

  `ConfirmFill:=False` suppresses both prompts, which is what makes the routine
  callable from the regression harness.

- Added `M_WriteBack_TryResolveTableColumn`, which reports whether the selection
  is a table data cell rather than raising when it is not. Keeping the expected
  negative distinct from a genuine fault also keeps the two distinguishable when
  the structured write result
  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21)) is added.

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

- Fixed the settings-namespace regression suite leaking registry keys. It created
  two temporary application namespaces to prove isolation and left both behind
  after every run, while reporting no cleanup failure:

  ```text
  VBA_DATETIMEPICKER__TST_DP_NS_A
  VBA_DATETIMEPICKER__TST_DP_NS_B
  ```

  `DeleteSetting` was called with an application name **and** a section, which
  removes that section and leaves an empty application key. The verification then
  looked for the probe value, found it absent, and concluded the namespace was
  gone — so the check meant to catch the leak was looking in the wrong place.

  Deletion now passes the application name alone, and verification uses
  `GetAllSettings`, which returns an unassigned `Variant` when the application key
  is absent and an array when it exists. That detects an empty leftover key,
  which a value lookup cannot.

  Found by inspecting the registry during the manual validation, not by the suite,
  which passed while leaking.
  ([#26](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/26))

- Fixed every deployment sharing one persisted settings namespace. All registry
  reads and writes used the fixed application name:

  ```text
  HKCU\Software\VB and VBA Program Settings\VBA_DATETIMEPICKER
  ```

  so the persistence scope was the Windows user, not the workbook, the add-in or
  the deployment. Two copies that never ran at the same time could still change
  each other's saved preferences — and `M_Settings_Load` ends by calling
  `M_Settings_Save` to persist the values it normalized, so merely *reading*
  settings rewrote the shared location.

  The coupling matters most for `HolidayCallback`, which is persisted as a
  callback name and later executed through `Application.Run`. A callback chosen
  for one workbook could be inherited by another purely through shared storage.
  The integration toggles — right-click, in-grid icon, keyboard shortcut, WinAPI
  — describe one deployment rather than a user preference and had the same
  problem.

  A caller can now isolate its own settings:

  ```vb
  M_Settings_SetNamespace "TreasuryTool"
  ```

  ```text
  no namespace configured   VBA_DATETIMEPICKER
  namespace "TreasuryTool"  VBA_DATETIMEPICKER__TreasuryTool
  ```

  Resolution is centralized in `M_Settings_GetEffectiveAppName`, and all 28 call
  sites obtain the name from it — `DP_SETTINGS_APP_NAME` now appears only inside
  the resolver. A namespace applied at each call site would eventually be wrong
  at one of them.

  The namespace locks once settings are loaded. Changing it afterwards is refused
  with a descriptive error, because values read from one namespace would then be
  written into another. Validation trims, bounds the length, and rejects control
  characters, path separators and the reserved `__` separator, so two valid
  namespaces can never resolve to one effective name.

  Deliberately **not** done: no automatic migration from the shared namespace
  into a new one, because importing it would carry across exactly the settings
  the namespace exists to separate. No namespace derived from workbook name or
  path, which change on rename, move or copy and would make settings appear to
  vanish. No version-specific namespace, which would make every upgrade look like
  a reset.
  ([#26](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/26))

- Fixed harness setup aborting a run over a worksheet Excel had already created.
  `Worksheets.Add` in `TST_DP_PrepareScratchSheet` intermittently reports `1004`
  **having created the worksheet anyway**. Captured:

  ```text
  Step=Add scratch worksheet | WorksheetsBefore=3 | WorksheetsAfter=4
      | OriginalSource=VBAProject | Cleanup=Sheet74 (deleted)
  ```

  The workbook mutation happens; only the call's completion does not. Every
  occurrence previously leaked an unnamed worksheet, and preflight could not see
  it — it looks for `TST_DP_SCRATCH` by name and the leftover is called
  `SheetNN`.

  Scratch-sheet setup is now one transaction. Protecting the `Add` alone was not
  enough: the rename, initialization, formatting and activation can each fail
  after a successful `Add` and leak the same worksheet.

  On failure, the worksheet this call created is identified by **object
  identity** against a pre-call snapshot, not by name. A worksheet name can
  contain any delimiter a string encoding would use, Excel compares names
  case-insensitively, and a rename would make a pre-existing sheet look new.

  What happens next depends on what is found:

  ```text
  no new worksheet          genuine failure; raise
  exactly one, validated    adopt it, record the anomaly, continue
  exactly one, unusable     delete that one; raise the original failure
  more than one             ambiguous; delete nothing; raise
  ```

  Adoption is confined to a failure of the `Add` itself, and the candidate must
  answer a rename, a write and a column-width change before it is accepted —
  Excel has been reported to create a worksheet that then cannot be renamed. A
  failure at the rename step is not adopted, because repeating an operation that
  was already refused would be a retry rather than a completion.

  Deleting every worksheet absent from the snapshot would assume this routine
  created them. That assumption is unavailable while a re-entrancy hypothesis is
  open, so two or more candidates are reported and none is deleted.

  A recovered run records what happened rather than looking clean:

  ```text
  INFO | Harness | Scratch sheet | Worksheets.Add reported 1004 after creating
  Sheet74. The worksheet was validated and adopted; WorksheetsBefore=3,
  WorksheetsAfter=4.
  ```

  Setup also rejects a protected workbook structure by name rather than surfacing
  a bare `1004` from the `Add`, checks the scratch name against `Sheets` rather
  than `Worksheets` because chart sheets share the name namespace, and anchors
  the `Add` on `Sheets(Sheets.Count)` so a trailing chart sheet is handled.

  The fresh-sheet lifecycle is deliberately preserved. `TST_DP_SCRATCH` is
  created at setup and deleted at successful teardown, because
  [#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19) uses its
  presence as the dirty-start evidence that survives a VBA project reset. A
  persistent worksheet would collapse that signal and would introduce residue
  between runs, which is a worse failure than an intermittent one because it
  fails quietly.

  The underlying Excel behavior is not explained by this change.
  ([#45](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/45))

- Fixed write-back destroying formulas. Nothing inspected whether a target cell
  held one: the per-cell writer refused missing cells, protected locked cells and
  array-formula cells, and wrote over everything else.

  Formula cells are now preserved by default and reported as a distinct
  classification:

  ```text
  cell is empty                        written
  cell holds a literal value           written
  cell holds a formula                 preserved and reported
  cell holds a date-returning formula  preserved and reported
  OverwriteFormulas:=True supplied     formulas written
  ```

  A formula that evaluates to a date is still a formula. Replacing a displayed
  date does not imply deleting what produced it.

  **Two gates, because one is not enough.** A rule in the per-cell writer alone
  would be bypassed on the most common multi-cell path:
  `M_WriteBack_TryBulkWriteRange` assigns across the whole target in one
  operation and never reaches per-cell inspection. The fast path is therefore
  refused whenever the target holds any formula, mirroring the array guard from
  [#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21).
  `Range.HasFormula` follows the same `True`/`False`/`Null` convention, and
  `Null` — the mixed case — refuses rather than being coerced to `False`. That
  coercion was the bypass this had to prevent.

  A single formula disables the fast path for the whole target. That is
  deliberate: partitioning the range would be faster on a long column containing
  one formula, at the cost of a scan, a block walk, and a new failure mode if the
  partition and the write disagree.

  The gate order is locked, then array, then formula. An array cell cannot be
  written at all, which is stronger and non-overridable, so it stays a failure
  whichever way `OverwriteFormulas` is set.

  `DP_WriteResult` gains `FormulaSkippedCount` and `FormulaSkippedAddresses`, and
  the accounting invariant extends to:

  ```text
  AttemptedCount = WrittenCount + LockedSkippedCount
                 + FormulaSkippedCount + FailedCount
  ```

  Every cell increments exactly one term, so it holds by construction.
  `M_WriteBack_DescribeShortfall` reports preserved formulas alongside the other
  classifications rather than through a second path.

  The override is a procedure parameter defaulting to protection, never a
  persisted setting, and no default UI path enables it — the same shape
  `NoTableGrow` already uses.
  ([#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22))

- Fixed a stale grid-icon reference raising `424 Object required` on teardown,
  and silently suppressing icon creation. Every use of the tracked icon was
  guarded by:

  ```vb
  If Not gDP_GridIconShape Is Nothing Then
  ```

  which tests the **variable**, not the object. The icon is an ordinary worksheet
  shape and can be destroyed without going through any routine that maintains the
  reference — the worksheet holding it is deleted, `M_GridIcon_PurgeAll` removes
  it by name from another workbook, or the user deletes it. The variable then
  still points at an object that no longer exists.

  The visible symptom was noise on every teardown:

  ```text
  M_GridIcon_Remove | Step=DeleteTrackedShape | Error=424 | Object required
  ```

  The more serious one was silent. `M_GridIcon_PreCreateHidden` guarded creation
  with the same test and jumped straight to hiding on a non-`Nothing` variable, so
  a stale reference meant **no icon was created at all** — and nothing reported
  it.

  `M_GridIcon_TrackedShapeIsLive` now asks the object model rather than the
  variable, clearing a reference whose shape no longer answers. A stale reference
  is an ordinary condition, not a failure to log. The check is applied at all
  eight sites that used the tracked reference, including the two that decided
  control flow from it rather than dereferencing it.

  Genuine deletion failures are still captured and reported; only the
  stale-reference case is removed from that signal.
  ([#44](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/44))

- Fixed the keyboard shortcut being enabled on the user's behalf. Disabling both
  the right-click entry and the in-grid icon forced `EnableKeyboardShortcut` back
  on, to avoid *a configuration with no practical access path*.

  `Ctrl + Shift + D` is registered through `Application.OnKey`, which holds one
  assignment per key for the whole Excel session. So the rule took a session-wide
  binding — possibly another add-in's — from a user who had explicitly turned the
  shortcut off. `M_Settings_SetEnableKeyboardShortcut(False)` also refused to
  store `False` in that configuration, so the setting could not be turned off
  even by asking directly.

  The premise was also shaky: `Ribbon_ShowPicker`, `DP_Show` and any
  caller-supplied button are access paths the rule did not count, and the
  component cannot detect at runtime whether the host package includes RibbonX.

  Registration now reflects explicit configuration only, and all three built-in
  interactive access paths may be disabled at once. The picker is then reached
  through `DP_Show`, `Ribbon_ShowPicker` or a caller-supplied entry point.

  The rule lived in **six** places, expressed two ways — three direct assignments
  to `gDP_EnableKeyboardShortcut` in `M_Settings_Load`, `M_Settings_Save` and
  `M_Settings_InitializeDefaults`, and three through a local in the setters. A
  search for the assignment finds half of them; the invariant is what has to be
  searched for.

  Teardown is unchanged and now documented rather than incidental.
  `M_KeyboardShortcut_Remove` restores Excel's default handling, which cannot
  bring back a displaced third-party binding — Excel exposes no getter for
  `Application.OnKey`, so that binding was never observable. It remains the least
  harmful of the three available behaviors: binding the key to an empty macro
  swallows it, and leaving the callback in place points at a project that may be
  unloading.
  ([#42](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/42))

- Fixed a failed frame refresh leaving the picker form half-styled.
  `M_Window_RemoveTitleBar` performs several native operations in sequence, and
  once `SetWindowLong` had cleared `WS_CAPTION` the original style was gone. A
  later `SetWindowPos` or `DrawMenuBar` failure was written to the Immediate
  Window and then **ignored**, leaving a window whose style said borderless and
  whose frame still showed a title bar:

  ```text
  style commit succeeds
      -> frame refresh fails
      -> routine logs and continues
      -> caller sees no failure
  ```

  Rollback was not merely absent, it was impossible: the original style was read
  into `WindowStyle` and then overwritten in place by the mask, so the value
  needed to restore it had been discarded.

  The operation is now transactional. The original style is kept in its own
  variable, a successful style write is the commit point, and a post-commit
  failure restores the original style and refreshes the frame through
  `M_Window_RollbackStyle`. A rollback that itself fails reports
  `RecoveryRequired` so the caller can rebuild the form rather than continue
  against a window in an unknown state. The original failure stays in
  `FailedStep` and `LastApiError` — a failed rollback must not overwrite the
  diagnostic that made rollback necessary.

  The closure invariant is derived on the common exit path rather than trusted to
  each branch, so a future branch cannot forget to mark a partially committed
  transaction unsafe:

  ```text
  Committed And Not Applied And Not RolledBack  ->  RecoveryRequired
  ```

  There is deliberately no shortcut for a caption bit that is already clear. The
  bit being clear proves a style write succeeded; it proves nothing about the
  frame refresh that should have followed, and skipping the refresh on that basis
  would make a half-applied window permanently unrepairable. That constraint is
  recorded in the routine banner.
  ([#20](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/20))

- Fixed composite teardown hiding which cleanup operation failed, or that any
  did. `TST_DP_ResetDatePickerArtifacts` wrapped five operations in one
  `On Error Resume Next` and cleared `Err` before returning, so the outer
  `TST_DP_CheckCleanupStep` could never observe a failure inside it:

  ```text
  M_ContextMenu_Remove fails silently
      -> the composite routine clears Err
      -> the outer cleanup step sees no failure
      -> CleanupFailures stays 0
  ```

  Teardown now invokes and checks each operation individually — `StopTimer`,
  `ClosePickerForm`, `RemoveContextMenu`, `RemoveKeyboardShortcut`,
  `PurgeGridIcons`. The composite routine survives as setup-only, where blanket
  suppression is appropriate: setup wants a usable starting state, teardown wants
  an account.

  Final-state verification also gained manager state and context-menu
  registration. Both are checked against their pre-run condition rather than
  against zero, because the harness restores the session it found — a DatePicker
  that was already running is entitled to its manager and its menu when the run
  ends. The context-menu check reports the count immediately after removal
  alongside the count at the end, since a mismatch has two causes needing
  different fixes: controls removal never took away, or controls something
  re-registered afterwards.

  Two teardown targets remain unverifiable from the harness and are recorded as
  such rather than quietly omitted. Excel exposes no getter for
  `Application.OnKey`
  ([#42](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/42)), and
  `mDP_TimerIsRunning` is `Private` to `M_DatePicker` with no way to enumerate
  `Application.OnTime` schedules. Both are covered instead by the per-operation
  accounting above.
  ([#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19))

- Fixed final-state verification covering three of the five Application settings
  the run snapshots. `TST_DP_CaptureApplicationState` captures
  `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `Calculation` and the status
  bar; `TST_DP_VerifyFinalState` asserted only `EnableEvents`, alongside the
  picker form and grid icons. A restore that silently failed on any of the other
  four was invisible.

  All are now verified, and each failure reports the expected and actual value
  rather than only naming the property.

  The status bar is the deliberate exception. Setting
  `Application.StatusBar = False` hands the bar back to Excel, but reading it
  immediately afterwards can still return the previous text, so neither a
  comparison against `False` nor a `VarType` test gives a stable answer at
  teardown. What is determinable — and what actually matters — is that the run
  left no message of its own, so that is what is asserted. The harness status
  text is now a single constant shared by the routine that sets it and the check
  that looks for it.
  ([#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19))

- Fixed the fast bulk write path never producing a written count. When a
  multi-cell target was written in one operation, `M_WriteBack_PopulateRange`
  returned before the section that calculates `WrittenCount`. Nothing depended on
  that section before, because the counts were discarded. A result fed only by
  the per-cell fallback would have reported the most common multi-cell success as
  nothing written.

  The bulk path now contributes its full target count explicitly, and the
  routine's notes record why, so the assignment is not read as redundant and
  removed. A failed bulk write is still not counted as a failed cell — the result
  describes the outcome of the logical write, not the optimization attempts
  behind it.
  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21))

- Fixed write-back reporting cells it had not written. Excel declines a value
  assignment to an array formula **silently** through the object model — it
  raises *You cannot change part of an array* only for an interactive edit. Three
  distinct silent outcomes were possible:

  ```text
  range assignment covering an array     replaces it, reports success
  range assignment overlapping an array  writes nothing, reports success
  cell assignment to an array cell       writes nothing, reports success
  ```

  On top of that, `WrittenCount` was derived by subtraction:

  ```vb
  WrittenCount = AttemptedCount - LockedSkippedCount - FailedCount
  ```

  That derivation treats every cell that did not raise as written, so a silent
  refusal was counted as a success. A write could report a full column written
  and change nothing.

  Two changes. `WrittenCount` now counts the cells that reported a successful
  write, rather than inferring them from what did not fail — every cell
  increments exactly one of written, locked or failed, so the accounting
  invariant holds by construction instead of by definition. And array cells are
  refused before the write on both paths: the fast path is skipped whenever
  `Range.HasArray` is `True` or `Null` for the target, and the per-cell writer
  refuses each array cell and records its address.

  This is a detectable inability to write, not a policy decision about formulas.
  The formula overwrite policy remains
  [#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22).
  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21))

- Fixed the skipped-cell message appearing once per target area. The summary for
  protected locked cells was raised inside `M_WriteBack_PopulateRange`, which
  `M_WriteBack_ApplyResolvedTarget` calls once for each member of
  `Target.Areas`. A discontiguous selection with locked cells in three areas
  produced three consecutive dialogs describing three fragments of one write.

  The write engine now resolves, writes, classifies and returns. It displays
  nothing at any level, including `M_WriteBack_Apply`. Notification moved to the
  interactive entry points — `M_Picker_SelectDate`, `DP_Now` and
  `DP_FillTableColumn` — which each call `M_WriteBack_ReportShortfall` once,
  after the complete result is available. `DP_Today` inherits this through
  `M_Picker_SelectDate`.

  A programmatic caller of `M_WriteBack_Apply` is no longer taken through modal
  UI, which is also what lets the harness exercise a partial write without a
  dialog stopping the run. Developer `Debug.Print` diagnostics remain.
  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21))

- Fixed six README statements describing `M_Picker_EnsureManager` behaviour that
  `v1.1.1` removed. The routine no longer forces
  `Application.EnableEvents = True`, but the public API table, the validation
  checklist, the troubleshooting table and a feature list still said it did.

  Two of them contradicted each other: **Known limitations** stated that
  `DP_RepairRuntime` is the only routine that deliberately re-enables events,
  while the troubleshooting table still offered `M_Picker_EnsureManager` as an
  equivalent recovery — sending a user whose events were disabled to a routine
  that, by design, does nothing to event state.

  The Known limitations section was audited when it was written; the rest of the
  file was not. The release documented its own behaviour change while
  contradicting it four sections above the disclosure.
  ([#41](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/41))
- Fixed four documents describing `demo/DATEPICKER.xlsm` as a tracked file after
  it was deleted. The `CONTRIBUTING.md` binary policy table was written
  specifically to resolve a contradiction about that file and had come to assert
  that a nonexistent file was tracked and must be kept in sync with source.

  The binary policy now leads with the rule rather than the exceptions: no
  workbook or add-in binary is tracked, and `UF_DatePicker.frx` is the only
  tracked binary in the repository.

- **Fixed a selected cell inside an Excel Table data column writing the entire
  column.** This applied to calendar selection, `DP_Today` and `DP_Now`, and it
  overwrote existing values in that column with no indication of the scope.

  The four `NoTableGrow` parameters now default to `True`, so an omitted
  argument means single-cell. `DP_Today` and `DP_Now` passed `False`
  explicitly — opting into expansion — and no longer do.

  Callers that already pass the argument are unaffected, in either direction:

  ```vb
  M_Picker_SelectDate SomeDate, False   'still expands to the column
  M_Picker_SelectDate SomeDate, True    'still writes one cell
  M_Picker_SelectDate SomeDate          'now writes one cell, previously expanded
  ```

  Only omitted arguments change, which is exactly the behaviour this release
  intends to change.

  No persisted setting governs this. The scope of a destructive write should not
  depend on a preference saved weeks earlier, or the same gesture would write
  one cell on one machine and a column on another.
  ([#13](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/13))

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

### 🔧 Changed

- The write-back chain now carries its result instead of discarding it.
  `M_WriteBack_Apply` became a `Function`; the three stages below it accumulate
  into a required `ByRef` result:

  ```text
  M_WriteBack_Apply(iType, [NoTableGrow])                    As DP_WriteResult
  M_WriteBack_ResolveAndApplyTarget iType, NoTableGrow, Result
  M_WriteBack_ApplyResolvedTarget   Target, iType, Result
  M_WriteBack_PopulateRange         oRange, iType, Result
  ```

  A `ByRef` output on `M_WriteBack_Apply` itself was considered and rejected. VBA
  does not permit a user-defined type as an `Optional` argument, so the parameter
  would have to be required, and every caller that does not want the result would
  still have to declare and pass one. The `Function` form avoids that churn. The
  private stages have no such constraint: their result parameter is required, and
  accumulating through the write path is what makes one discontiguous target
  produce one result.

  A `Function` called with bare-call syntax compiles unchanged. All eight
  `M_WriteBack_Apply` call sites were checked individually rather than assumed:

  ```text
  M_Picker_SelectDate       M_WriteBack_Apply DP_WriteAction_DatePicker, NoTableGrow
  DP_Now                    M_WriteBack_Apply DP_WriteAction_DatePicker
  DP_FillTableColumn        M_WriteBack_Apply DP_WriteAction_DatePicker, NoTableGrow:=False
  test/M_cDP_Test.bas       five calls, positional and omitted arguments
  ```

  None uses `Call`, parenthesises its arguments, or assigns the result, so none
  needed editing for the conversion. `M_WriteBack_Apply` already required an
  argument and so never appeared in the macro dialog; converting it changes
  nothing there.

  `M_WriteBack_TryWriteCell` replaced its two `ByRef` counters with the
  accumulating result, and now records the worksheet-qualified address behind
  every skip and failure rather than only counting them.
  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21))

- `DP_FillTableColumn` now checks its predicted scope against `AttemptedCount`.
  The routine already described the scope before writing:

  ```text
  Fill 247 cells in Trades[Expiry Date] with 25-Aug-2026?
  ```

  The predicted count was only resolved when prompting. It is now resolved on
  both paths and compared with `AttemptedCount` afterwards, which turns the
  prompt into a check.

  The comparison is deliberately against `AttemptedCount` rather than
  `WrittenCount`. A fill that legitimately skips protected cells is still a
  correct prediction:

  ```text
  Predicted scope / AttemptedCount   247
  WrittenCount                       244
  LockedSkippedCount                   3
  ```

  A prediction that does not match `AttemptedCount` means the target changed
  between preview and application, or that the two resolution paths diverged.
  That is reported as an inconsistency. The partial write is reported separately,
  through the same consolidated summary every other entry point uses.

  The routine is now a `Function` returning `DP_WriteResult`, so the harness can
  assert the comparison instead of reading a dialog. Both of its call sites are
  bare-call and unaffected: the regression suite, and `DP_Demo_FillTableColumn`
  behind the demo sheet button.
  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21))

- The demo workbook is now published as a release asset rather than tracked. It
  is generated from `demo/M_DP_DEMO.bas` by `DP_Demo_CreateDemoSheet` at release
  time, so a change to the demo is reviewable as a source diff.

  Release assets carry the version and contain no spaces, because GitHub
  replaces spaces with dots on upload and a spaced name is published under a
  different name than the one documented:

  ```text
  DATETIMEPICKER-vx.y.z.xlam
  DATETIMEPICKER-demo-vx.y.z.xlsm
  ```

- README, `CONTRIBUTING.md` and the pull request template brought current with
  the table write-scope change. `DP_FillTableColumn` is documented in the public
  API table, the write-scope callout describes the new default rather than
  warning about the old one, and **Table write scope** is removed from Known
  limitations.

  An operating note records what the fix did *not* change: the size of the
  resolved target moved, but formulas and existing values are still overwritten
  within it. That is
  [#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22).

  The pull request template's demo section now asks the questions that matter
  for generated output — rebuilt and visually checked, idempotent across two
  rebuilds, `OnAction` targets wired, and whether it still demonstrates current
  behaviour rather than superseded behaviour.

- Separated write-back target resolution from mutation.
  `M_WriteBack_ResolveAndApplyTarget` resolved the target and wrote to it in one
  routine, so nothing could ask what the target would be before it was written.

  ```text
  M_WriteBack_ResolveTarget         resolves and reports, writes nothing
  M_WriteBack_ApplyResolvedTarget   writes an already-resolved range
  ```

  The resolver returns the range plus expansion metadata: whether a single
  selected table cell was expanded to its data column, and the owning table and
  column names when it was. `M_WriteBack_ResolveAndApplyTarget` is retained as a
  thin delegation with its signature and default unchanged, so every existing
  caller is untouched.

  Write-action validation moved into the apply stage, so resolving a target for
  inspection does not require a valid write action. Error numbers are preserved:
  `513` unsupported action, `514` no selection, `515` non-range selection, `516`
  unresolved target, `517` empty target.

  No behaviour change. This is the seam that the write-scope fix
  ([#13](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/13)) uses to
  describe a resolved scope before writing it, and that the structured write
  result ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21))
  will later populate without re-deriving the target.

### 📖 Documentation

- The wiki's first `v1.2.0` pass corrects navigation, file paths and the
  `Application.EnableEvents` description. Nine pages changed.

  `_Sidebar` linked to a `Wiki Upload Guide` page that does not exist, and
  `UPDATE_NOTES` was a substantive page nothing linked to. The broken entry now
  points at the orphan: no broken links, no orphans.

  Eight references to `customUI.xml` across five pages became `customUI14.xml`,
  which is the file the project actually ships. The `<customUI>` XML root element
  in `Ribbon-Integration` is untouched — it is the element name, not a filename,
  and a blind replacement would have corrupted the example.

  `Manager-and-Events` and `Public-API` carried a warning banner above text that
  still said `M_Picker_EnsureManager must ensure Application.EnableEvents = True`.
  A banner over incorrect text is a holding measure, so the banners are gone and
  the text is rewritten: the routine preserves the caller's state and reports it
  through `EventsDisabledByCaller`.

  That makes `Is_Hooked = True` with events globally disabled a reachable state.
  The old page called it *the dangerous condition* the invariant existed to
  prevent; it is now the caller's decision, with `DP_RepairRuntime` as the
  documented way out. The pages say so rather than leaving it implied.

  `Home` claimed `v1.1.1` was in progress and cited `demo/DATEPICKER.xlsm`, a file
  deleted in this release. Its notice now records what has been verified against
  `v1.2.0` and names the two areas deliberately left alone:
  [#26](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/26) registry
  scope, and [#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37)
  multiple providers in one session.

  Page version metadata, the remaining `v1.2.0` API and behavior pages, and the
  final source-to-wiki verification pass are still outstanding.
  ([#17](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/17))

### 🔗 Compatibility

- No public procedure was added, removed or renamed in `src/`.
- `DP_Demo_CreateDemoSheet` and `DP_Demo_EnsureDemoSheet` are new, and live in
  `demo/`. Versioning covers the public VBA API of the component in `src/`;
  `demo/` is example material that ships as a workbook.
- Embedded behaviour is unchanged. The add-in gains a Demo button that works.
- The harness summary line gained `State=` and `CleanupFailures=` fields. Any
  tooling that parses it by position rather than by name will need updating.
- **`DP_FillTableColumn` is a new public procedure in `src/`.** Under the
  project's versioning policy that makes this release a minor rather than a
  patch, which it already was.
- **Table write scope changes for omitted arguments.** Code that relied on a
  calendar selection, `DP_Today` or `DP_Now` filling a table column must now
  either pass `NoTableGrow:=False` or call `DP_FillTableColumn`.
- **The keyboard shortcut is no longer forced on.** A configuration with
  right-click, in-grid icon and keyboard shortcut all disabled is now permitted
  and preserved. Anything relying on the shortcut re-enabling itself when the
  other two were disabled must now set `EnableKeyboardShortcut` explicitly.
- **Only one DatePicker copy may run per Excel session.** A second copy is
  refused at `DP_Start`, and cannot tear down the first through `DP_Stop` or
  `DP_RepairRuntime`. Sessions that previously loaded two copies — an `.xlam`
  alongside an embedded copy, for example — will now see the second refused with
  an explanation.
- **Settings persistence is unchanged by default.** An installation that
  configures no namespace reads and writes exactly where earlier releases did.
  `DP_SETTINGS_APP_NAME` keeps its value and meaning. Isolation is opt-in through
  `M_Settings_SetNamespace`, which must be called before anything loads settings.
- **Formula cells are no longer overwritten.** Code relying on write-back
  replacing a formula must now pass `OverwriteFormulas:=True` to
  `M_WriteBack_Apply` or `DP_FillTableColumn`. A fill that previously reported
  every predicted cell written may now report fewer, with the preserved
  addresses listed.
- **`DP_WriteResult` and `DP_WindowStyleResult` are new public types in `src/`.**
- **`M_Window_RemoveTitleBar` is now a `Function`.** Its argument list is
  unchanged and both `UF_DatePicker.frm` call sites use bare-call syntax, so
  neither needed editing.
- **`M_WriteBack_Apply` and `DP_FillTableColumn` are now `Function`s.** No
  argument list changed and no existing call site needed editing, because all of
  them use bare-call syntax. Code that wrapped either in `Call` with parentheses,
  or resolved either through `Application.Run`, needs checking — a `Function`
  returning a user-defined type cannot be called through `Application.Run`.
- **`M_WriteBack_PopulateRange` gained a required third argument.** It is an
  internal `M_` helper and outside the versioned public API, but any code calling
  it directly must now pass a `DP_WriteResult` accumulator.
- **A partial write is now reported once per operation, by the entry point.**
  The write engine displays nothing. A discontiguous selection with skipped cells
  in several areas produces one message instead of one per area, and a
  programmatic call to `M_WriteBack_Apply` produces none at all.

### ⚠️ Note for [#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22)

Building the failed-address regression established, empirically, that the fast
bulk write is not safe around formulas:

```vb
TargetRange.Value = WriteValue
```

Excel does not raise when it declines a write through the object model. An array
formula wholly inside the target is replaced without complaint; one partly
overlapping the target makes the assignment do nothing, also without complaint;
and a single-cell assignment to an array cell does nothing and returns normally.

`#21` handles the accounting consequence — count successes rather than infer them,
and refuse array cells on both paths. Two consequences carry into `#22`:

- A *skip formulas by default* rule cannot live in `M_WriteBack_TryWriteCell`
  alone, because the bulk path never reaches per-cell inspection. Formula
  detection has to gate the bulk write, exactly as the array check now does.
  `Range.HasFormula` follows the same `True`/`False`/`Null` convention.
- A refusal cannot be detected by catching an error, because there is no error to
  catch. Anything `#22` declines to overwrite has to be identified before the
  write is attempted, never inferred from its outcome.

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
