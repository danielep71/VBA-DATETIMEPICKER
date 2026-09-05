<!--
  Keep this pull request focused on one coherent outcome.
  Complete every common section. Delete optional profile blocks that do not apply.
  Use NOT RUN or NOT APPLICABLE with a reason; never manufacture PASS evidence.
  Record only checks and environments exercised against the exact candidate.
  Report vulnerabilities privately through SECURITY.md; do not disclose secrets,
  exploitable details, confidential workbooks, or restricted data in a pull request.
-->

<div align="center">

# 🔀 VBA DateTimePicker Pull Request

### Write-back safety · Runtime ownership · UI lifecycle · Exact evidence

[![Contributing](https://img.shields.io/badge/guide-CONTRIBUTING-217346?style=flat-square)](../CONTRIBUTING.md)
[![Security](https://img.shields.io/badge/security-private%20reporting-d73a49?style=flat-square)](../SECURITY.md)
[![Release](https://img.shields.io/badge/release-RELEASING-6f42c1?style=flat-square)](../RELEASING.md)
[![Changelog](https://img.shields.io/badge/changes-Unreleased-d97706?style=flat-square)](../CHANGELOG.md)

</div>

---

> [!IMPORTANT]
> Keep only the subsystem blocks touched by the pull request. The common evidence sections remain mandatory; do not retain hundreds of unchecked irrelevant boxes.

## 📌 Summary

<!-- State the observable outcome and why it is needed. Prefer one precise purpose. -->

## 🔗 Related issues

```text
Closes #
Related to #
```

Use a closing keyword only when this pull request satisfies the issue's complete acceptance criteria.

## 🧭 Change classification

- [ ] Defect correction
- [ ] Backward-compatible capability
- [ ] Breaking API, behavior, deployment, or migration change
- [ ] Internal refactor with no intended supported-behavior change
- [ ] Test, fixture, reference-data, or validation change
- [ ] Performance change
- [ ] Security or trust-boundary hardening
- [ ] Documentation-only change
- [ ] Repository tooling, workflow, or governance change
- [ ] Packaging or release preparation
- [ ] Write-scope, formula-safety, or data-integrity change
- [ ] Runtime ownership, event, timer, or recovery change
- [ ] UserForm, RibbonX, shape, keyboard, or WinAPI change

## 🎚️ Affected surface

- [ ] `DP_Start` / `DP_Stop` lifecycle
- [ ] Picker show, hide, preload, and close APIs
- [ ] Direct write-back and explicit table fill
- [ ] Settings namespace and persistence
- [ ] Application events, keyboard, context menu, grid icon, or timer
- [ ] `DP_RepairRuntime`, provider lease, and `DP_ForceReleaseProviderLease` recovery
- [ ] UserForm, label hooks, RibbonX, or WinAPI
- [ ] No runtime or supported surface — documentation/repository-only

---

## 📐 Scope and contract impact

### In scope

- <!-- Deliberate outcome -->

### Out of scope

- <!-- Reasonable adjacent work deliberately deferred -->

### Supported behavior and compatibility

```text
Supported behavior changed:       Yes / No
Backward compatible:              Yes / No / Uncertain
Suggested release impact:         none / patch / minor / major / uncertain
New supported members:
Removed or renamed members:
Changed signatures or defaults:
Changed results, errors, state, or side effects:
Migration required:
Known limitation introduced or retained:
```

Assess compatibility against documented behavior, not merely the VBA `Public` keyword. Infrastructure callbacks, Ribbon entry points, test seams, and `Application.Run` targets are not automatically supported API.

### Production source and package

`M_DatePicker`, `cDatePickerManager`, `cDatePickerLabelHook`, and the complete `UF_DatePicker.frm`/`.frx` pair; Ribbon XML is package metadata when used.

- [ ] Required source files and import order are unchanged.
- [ ] Required source files or order changed and `INSTALLATION.md` was updated.
- [ ] No production source/package impact.

## 🔧 Implementation notes

```text
Approach and key invariant:
Alternatives considered:
New dependency, reference, or generated input:
State ownership and cleanup:
Failure behavior:
```

Explain decisions a future reviewer cannot safely infer from the diff.

---

## ✅ Verification

### Candidate identity

| Evidence | Result |
| --- | --- |
| Exact PR HEAD SHA | <!-- Full 40-character SHA --> |
| Base branch and base SHA | <!-- Branch + full SHA --> |
| Working tree used locally | <!-- clean / dirty; explain --> |
| Source or package tested | <!-- Exact candidate source / artifact / N/A --> |

Evidence from another commit does not certify this candidate.

### Static and repository checks

- `git diff --check`
- Repository source/export inspection — no hosted static source gate is currently configured

| Check | Result / evidence |
| --- | --- |
| Hosted required checks | <!-- PASS / FAIL / NOT RUN + workflow URL --> |
| Local static command | <!-- Command + PASS / FAIL / NOT RUN --> |
| Formatting / `git diff --check` | <!-- PASS / FAIL --> |
| Machine-readable artifact | <!-- Name / URL / not produced --> |

### Excel and VBA execution

- [ ] Required and completed against the exact PR HEAD.
- [ ] Required but incomplete — reason and merge/release consequence stated.
- [ ] Not required — documentation/repository-only change with no executable or packaging impact.

Relevant entry points:

- `TST_DP_RunAll`
- `TST_DP_RunAll_WithUISmoke`
- Manual `DP_Show` / `DP_Close` / `DP_Stop` / `DP_Start` scenarios

| Evidence | Result |
| --- | --- |
| Tested commit SHA | <!-- Full SHA or N/A --> |
| `Debug → Compile VBAProject` | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Regression/certification entry point | <!-- Exact procedure --> |
| Completion state | <!-- PASS / FAIL / INCOMPLETE / NOT RUN --> |
| Cases / assertions / failures | <!-- Counts or N/A --> |
| Skipped / cleanup outcome | <!-- Counts and state or N/A --> |
| Focused and manual checks | <!-- Scenarios + result --> |
| Evidence file or workflow | <!-- Name / URL / N/A --> |

### Validation environment

```text
Excel product, version, and build:
Office bitness:                    32-bit / 64-bit
Windows version/build:
Workbook or add-in host:
Deployment model:
Deployment: embedded source / add-in
Other DatePicker provider and version
Settings namespace
Display scaling and monitor configuration
Worksheet protection and Excel event state
```

Record only tested environments. Source inspection does not constitute host execution, and one Office bitness does not execute the other conditional branch.

### Regression coverage

- [ ] Existing tests cover the changed success path.
- [ ] New or amended tests cover each corrected defect.
- [ ] Boundary, invalid-input, failure, fallback, and cleanup paths are covered as applicable.
- [ ] Test entry points and inventory/count metadata remain synchronized.
- [ ] Expected results come from the contract or an independent reference.
- [ ] No regression change is needed — rationale recorded below.

```text
Coverage rationale and new test names:
Unexecuted or deferred coverage:
```

---

## ⚠️ Risk, rollback, and recovery

- [ ] Low — documentation, metadata, or mechanically verified change.
- [ ] Medium — bounded runtime, tooling, or compatibility impact.
- [ ] High — numerical integrity, shared Excel state, native API, security, release, or breaking impact.

```text
Principal failure modes:
Residual risk after validation:
Rollback or revert procedure:
Excel-process, workbook, data, or artifact recovery:
Conditions that make rollback unsafe:
```

## 🔐 Security, data, and provenance

- [ ] No credential, secret, signing material, internal URL, or personal path is included.
- [ ] No client, employer, counterparty, student, personal, or restricted production data is included.
- [ ] Test data is synthetic, anonymized, or explicitly redistributable.
- [ ] External algorithms, code, datasets, and market/vendor data have attributable provenance and compatible licensing.
- [ ] Formula, command, path, callback, deserialization, and external-content injection surfaces were assessed.
- [ ] No security-sensitive detail belongs in private disclosure instead of this pull request.
- [ ] Generated evidence identifies its inputs, tool/runtime version, candidate SHA, and limitations.

```text
Security or privacy impact:
Source/data provenance:
New trust boundary:
```

## 📚 Documentation and release hygiene

- [ ] `README.md` reflects supported behavior and examples.
- [ ] `INSTALLATION.md` reflects paths, dependencies, import order, validation, upgrades, and removal.
- [ ] `CONTRIBUTING.md` reflects development and evidence requirements.
- [ ] `CHANGELOG.md` records material change under `[Unreleased]`.
- [ ] `SECURITY.md` reflects supported versions or trust boundaries.
- [ ] `RELEASING.md` reflects certification, package, provenance, or recovery changes.
- [ ] Source headers, API references, demos, Wiki pages, and counts remain synchronized.
- [ ] Version markers remain unchanged unless this is the deliberate release-stamp change.
- [ ] No documentation change is required — reason recorded below.

```text
Documentation impact:
Release, artifact, or migration impact:
```

---

## 🧩 Project-specific review

<details>
<summary><strong>🎯 Write-back, formulas, and table scope</strong></summary>

Keep when target resolution or cell writes can change.

- [ ] Ordinary interaction remains limited to the selected/resolved target.
- [ ] Whole-column table writes require explicit `DP_FillTableColumn` intent.
- [ ] Formula cells remain preserved by default.
- [ ] Destructive overwrite remains explicit and observable.
- [ ] Array-formula, protected, non-writable, merged, and partial-success paths are controlled.
- [ ] `DP_WriteResult` describes the real post-operation state.
- [ ] Caller event state is restored on every exit.

</details>
<details>
<summary><strong>🪪 Provider, settings, events, and timers</strong></summary>

Keep when runtime ownership or shared Excel registrations can change.

- [ ] Settings namespace is configured before settings load and retains its compatibility semantics.
- [ ] Every entry path enforces the single-provider lease before shared registration.
- [ ] A refused provider cannot remove the active provider's state.
- [ ] Manager startup, workbook lifecycle, VBA reset, repair, and teardown are covered.
- [ ] `Application.OnTime` cancellation uses the exact registered identity.
- [ ] Keyboard, CommandBar, event, and shape cleanup removes only owned resources.
- [ ] Recovery and force-release behavior fail closed where ownership is uncertain.

</details>
<details>
<summary><strong>🪟 UserForm, RibbonX, shape, and WinAPI</strong></summary>

Keep for visible UI or native-window changes.

- [ ] `UF_DatePicker.frm` and `.frx` remain synchronized.
- [ ] Label hooks and dynamically created controls are wired and released safely.
- [ ] Ribbon callback names and `customUI14.xml` remain synchronized.
- [ ] 32-bit and 64-bit declarations and valid-zero WinAPI returns are handled correctly.
- [ ] Positioning, borders, drag, scaling, multi-monitor, and modeless interaction are checked as applicable.
- [ ] No unsolicited production `MsgBox` or persistent worksheet artifact is introduced.

</details>
<details>
<summary><strong>📦 Regression, demo, and release package</strong></summary>

Keep for test harness, demo, add-in, or release-asset changes.

- [ ] Harness completion state distinguishes PASS, failure, cleanup failure, dirty start, and incomplete/skipped work.
- [ ] Demo-builder worksheet creation and partial-success cleanup are verified when affected.
- [ ] No fixed historical assertion count is treated as the current contract.
- [ ] The demo/add-in is built from the exact candidate and tested after final save.
- [ ] Source, form resources, Ribbon metadata, artifact hashes, and release notes identify one candidate.

</details>

---

## 👀 Reviewer focus

```text
Highest-risk decision:
Files and procedures to inspect first:
Evidence to challenge:
Known boundary not proved by this pull request:
Unresolved question or accepted trade-off:
```

## ☑️ Final author check

- [ ] The title describes the observable outcome.
- [ ] The pull request has one coherent purpose and no unrelated churn.
- [ ] Linked issue acceptance criteria are met or remaining work is explicit.
- [ ] Compatibility and release impact are assessed.
- [ ] Evidence belongs to the exact candidate claimed.
- [ ] Required checks are terminal and passing; incomplete work is not presented as PASS.
- [ ] Executable VBA was compiled and tested when required.
- [ ] Failure, cleanup, and recovery behavior were reviewed.
- [ ] The complete diff, including comments, metadata, binary companions, and documentation, was reviewed.
- [ ] No merge marker, stale placeholder, unexplained N/A, accidental binary, or private material remains.

---

**Review principle:** approve the smallest coherent change whose contract, evidence, risk, and recovery can all be explained from this pull request.
