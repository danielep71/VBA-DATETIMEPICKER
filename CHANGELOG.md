<div align="center">

# 📄 Changelog

**All notable changes to VBA-DATETIMEPICKER**

[![Semantic Versioning](https://img.shields.io/badge/Versioning-semver-6f42c1?style=flat-square)](https://semver.org/)
[![Format](https://img.shields.io/badge/Format-Keep_a_changelog-0969da?style=flat-square)](https://keepachangelog.com/)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#)

</div>

---

Versioning applies to the **documented supported consumer API and its established
behavioral contracts**. A member that is technically `Public` only because Excel,
Office, RibbonX, `Application.Run`, CommandBars, `OnTime`, Shape callbacks, or
test infrastructure must resolve it is **not automatically supported API**.

Compatibility-sensitive contracts include documented API names/signatures,
public enums and result types, established defaults, write-back semantics,
settings persistence, runtime ownership, and the stable
`VBA_DATETIMEPICKER` identifiers.

<details>
<summary><strong>Section legend</strong></summary>

<br>

| Section | Contains |
|---|---|
| ➕ **Added** | New supported capabilities, types, files or test capabilities |
| 🔧 **Changed** | Final behavior or contract changes |
| 🐛 **Fixed** | Defects corrected, with issue links for engineering detail |
| 📖 **Documentation** | Current-state user, contributor and repository documentation |
| 🧪 **Validation** | Evidence actually produced |
| 🔗 **Compatibility** | Upgrade impact |
| ⚠️ **Known limitations** | Deliberate boundaries that remain |

</details>

---

## [1.2.1] - 2026-08-26

> 🩹 **Patch** · integrity hotfix correcting safety-contract gaps shipped in `v1.2.0`

### 🧭 Release intent

`v1.2.1` corrects defects in behavior `v1.2.0` already claimed. It contains no
refactor and no new features. Corrections are published against the released
`v1.2.0` notes rather than by rewriting them.

### 🐛 Fixed

- **Provider lease admission is now enforced on every runtime entry path**
  ([#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37)).
  At `v1.2.0` only `DP_Start` consulted the provider lease. `DP_Show` and
  `DP_Preload` reached `M_Picker_EnsureManager` without proving ownership, and
  `DP_Click`, `DP_OpenForActiveCell`, the keyboard path and `Ribbon_ShowPicker`
  all funnelled into that same unguarded path. A refused second copy could
  therefore create a manager, load settings, hook Application events, register
  `Application.OnKey`, load a form, and later remove the true owner's
  registrations during its own teardown.

  Admission is now a single acquire-or-verify gate applied at the lowest shared
  boundary each path reaches. It is idempotent for the current owner, fails
  closed for a foreign or unverifiable lease, and mutates no shared state on the
  refusal path. `M_Picker_EnsureManager` additionally fails closed, so direct
  calls to that technically public bootstrapper cannot bypass admission.

- **A disabled keyboard shortcut is honored by the settings-panel save path**
  ([#42](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/42)).
  `UF_SettingsPanel_Save` forced the shortcut back on whenever right-click and
  the grid icon were both disabled, then persisted that value and registered it
  through `Application.OnKey`. A user who had deliberately disabled all three
  built-in entry paths could not keep that configuration.

  The equivalent block had already been removed from
  `M_Settings_SetShowRightClick` and `M_Settings_SetShowGridIcon`, which is why
  all three module setters still carried an empty `PROTECT MANUAL ACCESS PATH`
  banner. The UserForm copy was missed, so setter-level coverage passed while
  the real panel still overrode the user's choice. The save path now resolves
  through a single pure seam that returns the current setting unchanged.

- **A write never reports partial mutation as an exception carrying no result**
  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21)).
  `M_WriteBack_PopulateRange` raised on both a zero-write area and an unexpected
  technical error, and both raises sat before the accumulation block. An area
  that raised discarded its own classification counts, and because the raise
  escaped the area loop it took every area already accumulated with it. A
  writable area followed by a failing one mutated the workbook and then returned
  no `DP_WriteResult` at all, and the totals depended on the order Excel
  enumerated `Target.Areas`.

  A zero-write area is now an outcome rather than an exception, so totals are
  complete and order-independent. An unexpected technical error is now decided by
  one rule: whether any cell, in this area or an earlier one, has produced an
  outcome. If none has, the original error is raised with its identity preserved,
  because nothing is destroyed by raising and a programming error reaching the
  engine must stay loud. If one has, the error is recorded in the result and the
  facts observed so far are returned.

  Population stops at the first technical failure rather than continuing to
  mutate the workbook, so the areas it stopped short of are absent from
  `AttemptedCount`. A technical-failure result is therefore **bounded** —
  `Written + Locked + Formula + Failed <= Attempted` — rather than balanced, and
  because the counts alone can then look complete, the failure is reported
  separately rather than inferred from them.

- **An unrecoverable window-style failure no longer presents the form**
  ([#47](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/47)).
  Both call sites of `M_Window_RemoveTitleBar` discarded its return value, so a
  `DP_WindowStyleResult` reporting `RecoveryRequired` was dropped in silence and
  the picker went on to show with a window style that had been partly applied
  and could not be rolled back.

  Both sites now consume the result. `UserForm_Initialize` records the failed
  step and the last API error and then fails the load, which is the only
  available action for an instance still under construction.
  `UserForm_Activate` records the same diagnostics, marks itself activated
  before anything else so the unload cannot re-enter the path, and unloads the
  form.

- **A handler that cleans up before re-raising now reports the error that
  actually occurred**
  ([#48](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/48)).
  Every `On Error` statement resets the `Err` object, and so does `Err.Clear`.
  Two handlers cleaned up and then raised from the live `Err` object, so the
  caller received error `0` with a blank description instead of the cause.

  `DP_FillTableColumn` restored `gDP_WriteValue` under `On Error Resume Next` and
  cleared `Err` before raising. The loss was conditional on the pending write
  value having been staged, which is true for every failure inside the fill
  itself. `UF_SettingsCheckBoxFont_Create` released its temporary font object
  under `On Error Resume Next` and then restored handling with `On Error GoTo 0`.
  It never called `Err.Clear` and did not need to, so its loss was unconditional.

  Both now capture number, source and description at handler entry before
  anything touches `On Error`, run cleanup separately, capture any cleanup
  failure separately, and re-raise the original with cleanup diagnostics appended
  when useful.

  All production source was audited mechanically: 243 `Err.Clear` sites, 87
  `Err.Raise` statements reading the live `Err` object, and exactly two where
  clearing precedes the raise — the two above. After the fix the same audit
  reports 85 live-`Err` raises and no defect, so no further occurrence exists.

### 🔧 Changed

- A refused provider now reports once per user action rather than once per
  delegating layer. `DP_Preload` refuses silently, because a background startup
  optimization must not raise a second message box after `DP_Start` has already
  reported.

- A write operation that writes no cell now returns a complete `DP_WriteResult`
  with `WrittenCount = 0` instead of raising. Callers report it through
  `M_WriteBack_ReportShortfall` as a normal shortfall message. Single-area
  writes are unaffected.

- `M_Picker_SelectDate` and `DP_Now` store the selected date only when at least
  one cell received it. A zero-write previously raised out of the engine and
  never reached those assignments, so this preserves the prior contract now
  that a zero-write returns normally.

- A window-style failure reported as `RecoveryRequired` is now terminal for that
  form load. The picker fails rather than presenting a window whose style is
  neither fully applied nor rolled back. A subsequent load is unaffected.

- `DP_WriteResult` gains `TechnicalFailureOccurred`, `TechnicalFailureStep`,
  `TechnicalFailureNumber` and `TechnicalFailureDescription`. The addition is
  backward compatible: existing field names, types and meanings are unchanged.
  Adding members to a public result type would ordinarily require a minor bump
  under the versioning policy stated at the top of this file. It is carried in a
  patch because the fields are the reporting mechanism for the #21 correction —
  a bounded result cannot be described without them — and because no supported
  member is added, removed, renamed or redefined. Recorded as a deliberate
  deviation rather than left to inference.

- `M_WriteBack_ReportShortfall` no longer stays silent when `WrittenCount`
  reaches `AttemptedCount` if a technical failure is flagged, and
  `M_WriteBack_DescribeShortfall` names the step the operation stopped at. The
  areas an aborted operation never reached are not counted, so the counts alone
  could otherwise describe a smaller operation than the user asked for.

### 📖 Documentation

- **The Wiki no longer teaches a shutdown that bypasses ownership**
  ([#17](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/17)).
  `Workbook-Lifecycle` recommended an explicit `BeforeClose` teardown that set
  `gDP_Manager = Nothing` and then called `M_ContextMenu_Remove`,
  `M_KeyboardShortcut_Remove`, `M_Timer_Stop`, `DP_Close` and
  `M_GridIcon_PurgeAll` directly. It never called `DP_Stop` and never released the
  provider lease.

  An owner that ran the recipe stranded ownership for the rest of the Excel
  process. A refused second copy that ran it removed the true owner's shortcut,
  menu entry and icons: `DP_Stop` and `DP_RepairRuntime` refuse a caller that
  cannot prove ownership, but the low-level helpers do not check, because they are
  the internals those APIs call once admission is already proven. The published
  recipe therefore bypassed the protection the one-provider lease exists to give.

  `BeforeClose` now calls `DP_Stop`. The old recipe is explained in place rather
  than deleted, so the record of what was published stands, and the low-level
  helpers are now identified as internal and diagnostic rather than substitutes
  for the ownership-aware lifecycle API.

- `Public-API` omitted `FailedCount` and `FailedAddresses` from the
  `DP_WriteResult` listing while the accounting invariant printed beneath it
  referred to `FailedCount`. Both fields existed in the source throughout; only
  the page was incomplete. They are restored, the four `TechnicalFailure` fields
  are documented, and the invariant is stated as holding for a result that
  completed, with the bounded technical-failure case described separately.

- `Public-API` address-cap wording now matches the code. The 25-address cap is
  applied per target area rather than per operation, so a discontiguous write can
  report more than 25 addresses in a category while the classification totals stay
  exact. Moving to one cap per operation is deferred to `v1.2.2` and is not yet filed as
  an issue.

- `Installation-and-Import` showed a repository tree containing a tracked
  `demo/DatePicker Demo.xlsm`. No demo workbook is tracked: `demo/` holds the VBA
  source that builds it and the built file is a Release asset, as the root
  `README` already states. The tree now matches the repository, including `test/`,
  `dist/` and `images/`, and no longer implies the Wiki is a folder inside it.

- `Testing-and-Demo-Guide` moves from 16 suites and `Run=302` to 24 standard
  suites and `Run=431`, and lists the six suites added by this patch against their
  issues. The 16 was already stale when published: the runner ran 18 standard
  suites at `v1.2.0`.

- Wiki claims about lease refusal and the keyboard shortcut needed no rewrite.
  They were published as `v1.2.0` behavior, were false against `v1.2.0` code, and
  are true against this patch. They are corrected here rather than on those pages,
  so released history is not rewritten.

### 🧪 Validation

- Standard regression, run on both the embedded `.xlsm` and the packaged
  `.xlam`: `State=PASS; Run=431; Passed=431; Failed=0; CleanupFailures=0`
- With UI smoke: `State=PASS; Run=434; Passed=434; Failed=0; CleanupFailures=0`
- New `RuntimeAdmission` suite — 28 assertions driving every public entry path
  under both a foreign and an owned lease.
- New `SettingsSaveResolution` suite — 12 assertions driving the seam the real
  save handler executes, sweeping all eight input combinations.
- New `MultiAreaWriteResult` suite — 35 assertions driving the real public write
  path with a two-area `Union`, repeating every scenario with the areas swapped.
- New `WindowRecovery` suite — 8 assertions driving the real form through an
  injected fault at `SetWindowPos` and at the rollback step, covering both the
  initialize and activate paths and confirming that the next load succeeds.
- New `WriteTechnicalFailure` suite — 35 assertions driving the real public write
  path with a fault armed after cells were written inside one area, on entering a
  later area after an earlier one completed, and before anything had been
  observed at all. One scenario uses an earlier area that only skipped a formula
  cell, which mutated nothing but still holds an outcome that must survive.
- New `ErrorPreservation` suite — 11 assertions driving the real public
  `DP_FillTableColumn` against a real Excel Table with a write fault armed,
  asserting the caller receives the original error rather than error `0`, that
  the description names both the failing operation and the original cause, and
  that cleanup still ran.
- Baseline 302 + 28 + 12 + 35 + 8 + 35 + 11 = 431 standard; 24 standard suites,
  25 registered including UI smoke.
- Certified at `7d55cc7`, tag `v1.2.1`. Three earlier candidates — `ab15c92`,
  `359d2ca`, `128a84f` — were voided during certification. All four defects found
  were in the regression apparatus, not the component: `git diff ab15c92 7d55cc7
  -- src/` is empty, so the certified source is byte-identical to the first
  candidate. One of those defects meant the regression pack had never been
  runnable inside a packaged `.xlam` at all, so the packaged figures above are
  the first this project has gathered from the artifact it ships. That narrows,
  but does not close, the packaging boundary recorded for `v1.2.0`
  ([#16](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/16)).
- Certification ran on a developer workstation with other add-ins loaded in the
  same Excel process, not on a clean VM.
- Manual two-provider refusal matrix passed for `.xlam + embedded` and
  `embedded + embedded` in a single Excel process, including a reversal check
  confirming ownership follows start order rather than packaging.

### 🔗 Compatibility

- No supported API name, signature or default changed.
- Provider-admission behavior change is limited to paths that were already
  specified as refusing; a single participating provider sees no difference
  there. The #21, #42 and #47 corrections do change what a single provider
  observes: a zero-write returns a complete `DP_WriteResult` instead of raising,
  a `RecoveryRequired` window-style failure fails the form load instead of
  presenting it, and saving the settings panel no longer re-enables
  `Ctrl + Shift + D`. In each case the new behavior is what `v1.2.0` already
  specified.
- `M_Lease_Test_SilenceRefusalReport` and `M_Lease_Test_RefusalReportCount` are
  internal test infrastructure, not supported API. They exist because
  `Application.DisplayAlerts` does not suppress `VBA.MsgBox`, so automated
  coverage of the real entry paths would otherwise block on a modal dialog.
  Both are to be classified `internal` under
  [#25](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/25).
- `M_Settings_ResolveKeyboardShortcutOnSave` is the settings-save resolution
  seam. It is internal implementation surface, not supported API, and is also to
  be classified `internal` under
  [#25](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/25).
- `M_WriteBack_Test_SetFaultInjection` is internal test infrastructure, not
  supported API, and is also to be classified `internal` under
  [#25](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/25). It exists
  because the path it covers cannot be produced on demand:
  `M_WriteBack_TryWriteCell` classifies every per-cell failure it can observe and
  raises nothing, so a genuine technical error inside the area loop is precisely
  what a test has no way to arrange. Injection is one-shot, disarms itself as it
  fires, and persists nothing.

- `#48` is delivered with two stated limitations rather than silent ticks.
  Cleanup failure is not injected by regression: the cleanup steps in both
  handlers are a `Date` assignment and an object release, neither of which can be
  made to fail without a seam for a failure mode that does not exist. The
  capture-and-append structure is implemented and uniform with
  `M_WriteBack_Apply`, but the append branch is defensive.
  `UF_SettingsCheckBoxFont_Create` has no direct regression because it is
  `Private` to the form and reached only from the settings-panel build path; its
  correction rests on review and on the mechanical audit.

### ⚠️ Known limitations

- Simultaneous active providers remain unsupported. This release closes
  admission bypasses under the exclusive-provider model; true coexistence
  remains [#14](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/14).
- Automatic reclamation of a stale lease is still deliberately absent. Recovery
  after a VBA project reset remains the explicit
  `DP_ForceReleaseProviderLease` call.
- A pre-existing `Ribbon_Demo` demo-sheet toggle defect is not fixed here: the
  callback builds the demo sheet visible, then reads it as already visible and
  hides it. It predates `v1.2.0`, is out of scope for an integrity hotfix, and is
  deferred to `v1.2.2`. It is not yet filed as an issue.

---

## [1.2.0] - 2026-08-25

> ✨ **Minor** · safety, observability and runtime-ownership release

> 📌 **Corrected by [1.2.1](#121---2026-08-26).** This entry is preserved as
> published. Five contracts recorded below were specified but not fully
> delivered in `v1.2.0`, and were corrected in the `v1.2.1` integrity hotfix:
> provider refusal was enforced only at `DP_Start` ([#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37)); the
> settings-panel save path still forced the keyboard shortcut back on
> ([#42](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/42)); the accounting invariant holds only for a result that
> completed ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21)); `DP_WindowStyleResult` was produced but
> discarded at both call sites ([#47](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/47)); and four Wiki pages from
> the rewrite were wrong or stale ([#17](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/17)). Read this entry
> together with `v1.2.1`.

### 🧭 Release intent

`v1.2.0` makes four contracts explicit:

1. **what the DatePicker writes** — safe selected-cell defaults, explicit broad
   Table scope and formula preservation;
2. **what it owns** — one participating current-version provider per Excel
   process;
3. **what happened** — structured worksheet/native outcomes and a harness whose
   `PASS` includes clean start, completion and cleanup;
4. **what the repository proves** — source, package validation and release
   provenance are kept as distinct claims.

This entry records the **final release state**. Detailed rationale, implementation
history, exact commits and closure evidence remain in the linked GitHub issues.

#### At a glance

| Area | Final v1.2.0 outcome | Issue |
|---|---|---|
| Table scope | Normal write-back is selected/resolved-cell; Table-column fill is explicit | [#13](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/13) |
| Wiki | Full v1.2.0 rewrite with page-level source baseline | [#17](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/17) |
| Harness errors | Real escaping error number/description preserved | [#18](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/18) |
| Harness verdict | Five states; dirty start and cleanup affect `PASS` | [#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19) |
| Native styling | Transactional result, rollback and `RecoveryRequired` | [#20](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/20) |
| Write observability | `DP_WriteResult`, counts and bounded addresses | [#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21) |
| Formula safety | Preserve by default; explicit override | [#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22) |
| Demo | Source-built and add-in-safe | [#23](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/23) |
| Settings | Optional stable persistence namespace | [#26](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/26) |
| Provider ownership | Second v1.2.0 provider refused safely | [#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37) |
| README drift | API/event/write-scope text reconciled | [#41](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/41) |
| Keyboard | Shortcut follows explicit configuration only | [#42](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/42) |
| Grid icon | Stale Shape references detected and cleared | [#44](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/44) |
| Harness setup | Partial-success `Worksheets.Add` handled transactionally | [#45](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/45) |

Repository findings [#33](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/33),
[#34](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/34) and
[#38](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/38) were
consolidated into their owning provenance/workflow follow-up work. The automated
traffic alert [#39](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/39)
was classified as operational analytics rather than DatePicker remediation.

---

### ➕ Added

- **`DP_FillTableColumn`** — explicit supported whole-Table-data-column write:

  ```vb
  Public Function DP_FillTableColumn( _
      ByVal ValueToWrite As Date, _
      Optional ByVal ConfirmFill As Boolean = True, _
      Optional ByVal OverwriteFormulas As Boolean = False) As DP_WriteResult
  ```

  Confirmation describes the resolved Table/column scope before mutation, and
  the applied target is checked against predicted `AttemptedCount`.
  ([#13](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/13),
  [#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21),
  [#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22))

- **`DP_WriteResult`** — one structured write outcome covering attempted,
  written, locked-skipped, formula-skipped and failed cells; worksheet-qualified
  addresses; resolved target/Table metadata; area count; and caller event state.

  Final accounting invariant:

  ```text
  AttemptedCount =
      WrittenCount +
      LockedSkippedCount +
      FormulaSkippedCount +
      FailedCount
  ```

  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21),
  [#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22))

- **`DP_WindowStyleResult`** — structured native-window outcome with
  `Attempted`, `Applied`, `Committed`, `RolledBack`, `RecoveryRequired`,
  `FailedStep` and `LastApiError`.
  ([#20](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/20))

- **Settings namespaces** — `M_Settings_SetNamespace` /
  `M_Settings_GetNamespace` isolate persistent preferences while retaining
  `VBA_DATETIMEPICKER` as the backward-compatible default. Namespace selection
  must happen before settings load; no automatic migration occurs.
  ([#26](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/26))

- **One-provider lease** — a temporary process-visible CommandBar lease is
  acquired before shared registrations. A second participating `v1.2.0`
  provider is refused; `DP_Stop` and `DP_RepairRuntime` are ownership-guarded.
  `DP_ForceReleaseProviderLease` is explicit operator-only recovery.
  ([#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37))

- **Source-built demo and stronger harness tooling** — the demo is generated from
  `M_DEMO_BUILDER.bas` / `M_DP_DEMO.bas`; the harness adds clean-start
  preflight, setup/environment diagnostics, self-checks and final-state
  verification. Its verdict states are `PASS`, `FAIL`, `FAIL_CLEANUP`,
  `FAIL_DIRTY_START` and `INCOMPLETE_SKIPPED`.
  ([#18](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/18),
  [#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19),
  [#23](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/23))

---

### 🔧 Changed

- **Table writes are safe by default.** Calendar selection, `DP_Today` and
  `DP_Now` no longer infer a whole Table data-column write from one selected
  Table cell. Use `DP_FillTableColumn` when broad scope is intended.
  ([#13](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/13))

- **Formulas are preserved by default.** Ordinary and date-returning formulas
  are reported as formula skips; ordinary formulas can be replaced only through
  an explicit `OverwriteFormulas:=True` call where supported. Array formulas
  remain non-overridable write failures.
  ([#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22))

- **Keyboard access is explicit.** `Ctrl + Shift + D` is no longer forced on
  when right-click and the grid icon are disabled. Zero built-in interactive
  access paths is a valid configuration. Removal restores Excel default handling,
  not an unobservable third-party predecessor.
  ([#42](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/42))

- **Borderless styling is transactional.** A post-commit native failure attempts
  rollback; rollback failure produces `RecoveryRequired=True` without replacing
  the original failure diagnostic.
  ([#20](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/20))

- **Distribution is source-first.** Generated `.xlam` and demo `.xlsm` files are
  release assets; committed VBA/Ribbon/UserForm/demo/test source remains the
  reviewable source of truth.
  ([#23](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/23))

---

### 🐛 Fixed

- **Write-back integrity:** fixed silent/non-raising Excel write outcomes around
  array formulas, incorrect bulk-path written counts, and per-area shortfall
  dialogs. `WrittenCount` now represents observed logical writes and one
  interactive operation produces one shortfall summary.
  ([#21](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/21))

- **Formula destruction:** both bulk and per-cell paths now enforce the same
  default preservation policy.
  ([#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22))

- **Harness error evidence:** suite handlers no longer erase the original error
  before recording it.
  ([#18](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/18))

- **Harness run integrity:** cleanup failures, incomplete suites and dirty
  predecessor state can no longer result in `PASS`; final observable state is
  verified after teardown where Excel exposes it.
  ([#19](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/19))

- **Scratch-sheet setup:** an observed `Worksheets.Add` partial-success condition
  is recovered only when one new worksheet can be identified and validated.
  Zero candidates fail; multiple candidates are ambiguous and none is deleted.
  No blind retry is performed.
  ([#45](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/45))

- **Settings isolation/cleanup:** deployments can opt into independent persistence,
  and regression namespace cleanup no longer leaves empty test application keys.
  ([#26](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/26))

- **Grid-icon lifecycle:** stale retained `Shape` references no longer raise
  teardown noise or suppress future icon creation; stale state is detected and
  cleared before use.
  ([#44](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/44))

- **Documentation/source drift:** README statements and the full Wiki were
  reconciled with the final v1.2.0 API and behavior.
  ([#17](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/17),
  [#41](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/41))

---

### 📖 Documentation

- Rebuilt `README.md` as the current technical landing page.
- Added `INSTALLATION.md` for source/add-in deployment, upgrade and recovery.
- Rebuilt `CONTRIBUTING.md` around v1.2.0 engineering contracts.
- Expanded `SECURITY.md` and `CODE_OF_CONDUCT.md` around trust boundaries and
  evidence-led collaboration.
- Rebuilt the pull-request, bug-report and feature-request templates around
  write scope, formulas, provider ownership, namespaces, native outcomes and
  harness evidence.
- Hardened `.gitignore` for the source-first repository policy.
- Completed the Wiki rewrite against source baseline:

  ```text
  6435c9170f1707a6269f2e307d158a0faf0cae21
  ```

  All 21 substantive pages carry:

  ```text
  Applies to:      v1.2.0
  Reviewed commit: 6435c91
  ```

  ([#17](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/17))

---

### 🧪 Validation

Final recorded standard source regression:

```text
State=PASS; Run=302; Passed=302; Failed=0; CleanupFailures=0
```

Additional recorded evidence includes formula/literal/array-formula write
classification, independent native-window state verification, settings namespace
isolation/cleanup, and a real two-project `v1.2.0` provider matrix covering
refusal, refused teardown/repair, ownership transfer, stale-lease failure,
Excel-restart recovery, WinAPI-disabled operation and clean release.
([#20](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/20),
[#22](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/22),
[#26](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/26),
[#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37))

The final packaged `v1.2.0` `.xlam` remains a **release-certification boundary**:
source regression does not by itself establish package-specific smoke or exact
source↔asset provenance.

---

### 🔗 Compatibility

`v1.2.0` is a **minor release**.

- No existing documented supported consumer member is removed or renamed.
- New supported capabilities/types include:

  ```text
  DP_FillTableColumn
  DP_WriteResult
  DP_WindowStyleResult
  M_Settings_SetNamespace
  M_Settings_GetNamespace
  DP_ForceReleaseProviderLease
  ```

- **Table scope:** existing code relying on omitted/default arguments to fill a
  whole Table column must move to explicit `DP_FillTableColumn` or an intentional
  internal broad-scope call.

- **Formulas:** code intentionally replacing ordinary formulas must explicitly
  opt in to formula overwrite on a supported write path.

- **Keyboard:** code relying on the shortcut re-enabling itself must now enable it
  explicitly.

- **Providers:** only one participating current-version provider may own an Excel
  process. A second is refused and cannot tear down/repair the owner.

- **Settings:** installations that configure no namespace keep the exact legacy
  persistence location/meaning. Isolation is opt-in and must be selected before
  settings load.

- **Partial writes:** skipped/failed cells are now observable through
  `DP_WriteResult`; interactive reporting is consolidated once per operation.

- **Harness consumers:** parse the named `State=...` verdict and fields rather
  than treating assertion totals alone as certification.

Technically `Public` callbacks/test seams and undocumented `M_...` helpers are not
automatically part of the supported consumer contract; code calling them directly
should be reviewed when upgrading.

---

### ⚠️ Known limitations

- **True multi-provider coexistence is not implemented.** v1.2.0 provides one
  owner plus refusal, not arbitration.
  ([#14](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/14))

- **Mixed-version sessions are unprotected.** A pre-v1.2.0 provider does not
  participate in the lease protocol.

- **`Application.OnKey` predecessor state is unobservable.** Teardown restores
  Excel default handling, not an unknown displaced macro.

- **Default settings are still Windows-user-global** unless a stable namespace is
  configured before settings load.

- **Exact source↔release-asset provenance is not yet guaranteed.**
  ([#16](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/16))

- **Automated software-quality / release certification is still incomplete.**
  Traffic analytics is operational telemetry, not DatePicker CI.
  ([#15](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/15))

- **The technically `Public` surface remains larger than the supported API.**
  Formal supported/callback/internal classification remains follow-up work.
  ([#25](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/25))

- **Accessibility, high-DPI and high-contrast behavior is not fully certified**
  across Office/display configurations.
  ([#29](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/29))

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

[Unreleased]: https://github.com/danielep71/VBA-DATETIMEPICKER/compare/v1.2.1...HEAD
[1.2.1]: https://github.com/danielep71/VBA-DATETIMEPICKER/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/danielep71/VBA-DATETIMEPICKER/compare/v1.1.1...v1.2.0
[1.1.1]: https://github.com/danielep71/VBA-DATETIMEPICKER/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/danielep71/VBA-DATETIMEPICKER/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/danielep71/VBA-DATETIMEPICKER/releases/tag/v1.0.0
