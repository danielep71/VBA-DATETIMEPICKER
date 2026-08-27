<div align="center">

# 🔒 Security Policy

### Security, trust boundaries, responsible disclosure, and safe deployment for VBA-DATETIMEPICKER

[![Reporting](https://img.shields.io/badge/reporting-private-d97706?style=for-the-badge)](#-reporting-a-vulnerability)
[![Support](https://img.shields.io/badge/support-latest_tagged_release-217346?style=for-the-badge)](#-supported-versions)
[![Platform](https://img.shields.io/badge/platform-Excel_VBA_%2F_Windows-0078D6?style=for-the-badge)](#-security-model)
[![Scope](https://img.shields.io/badge/scope-source_runtime_and_release_assets-6f42c1?style=for-the-badge)](#-scope)
[![Automation](https://img.shields.io/badge/automation-credential_isolated-d73a49?style=for-the-badge)](#-repository-automation-credentials)

<br>

**Source-first trust · Explicit runtime ownership · Safe write scope · Formula preservation · Private disclosure · Release provenance**

<br>

[Supported versions](#-supported-versions)
&nbsp;·&nbsp;
[Report privately](#-reporting-a-vulnerability)
&nbsp;·&nbsp;
[Security scope](#-scope)
&nbsp;·&nbsp;
[Runtime boundaries](#-runtime-security-boundaries)
&nbsp;·&nbsp;
[Supply chain](#-supply-chain-and-release-integrity)
&nbsp;·&nbsp;
[Automation credentials](#-repository-automation-credentials)
&nbsp;·&nbsp;
[Verify a release](#-verifying-a-release)

</div>

---

**VBA-DATETIMEPICKER** distributes reviewable Excel/VBA source and, for
convenience, macro-enabled release binaries.

The production component runs with the privileges already granted to Microsoft
Excel and the current Windows user.

There is no:

- background service;
- package manager;
- bundled third-party DLL;
- privileged installer;
- credential store in the DatePicker runtime;
- network client in the production DatePicker source;
- automatic update mechanism;
- external executable required by the component.

The attack surface is therefore relatively small, but it is not zero.

The project interacts with security-relevant or integrity-sensitive surfaces:

```text
VBA macros
Excel Application events
Application.OnKey
Application.OnTime
Office CommandBars
worksheet Shapes
Excel Tables and worksheet write-back
registry-backed VBA settings
modeless MSForms UserForms
optional Windows API calls
macro-enabled .xlsm / .xlam release artifacts
GitHub Actions automation and repository credentials
```

Responsible disclosure matters because a defect in one of those surfaces can
affect more than the visible DatePicker UI.

> [!IMPORTANT]
> The DatePicker is **not a security boundary**.
>
> It does not enforce authorization, workbook permissions, segregation of duties,
> data-access controls, macro trust, or Windows security.
>
> Its safety mechanisms — formula preservation, explicit table scope, provider
> ownership, settings namespaces, structured results and native rollback — are
> designed to prevent accidental corruption and cross-component interference.
> They are not a sandbox against malicious code already running with Excel/VBA
> privileges.

---

## 🧭 Security model

The project assumes:

```text
Excel itself is trusted
the VBA project containing the DatePicker is trusted
the current Windows user is authorized to run the workbook/add-in
macros are enabled through an approved trust mechanism
```

The project does **not** assume:

```text
every workbook open in the same Excel process is cooperative
every add-in uses different application-wide resources
every worksheet is unprotected
every formula may be destroyed safely
every Application.OnKey assignment is observable
every VBA project reset leaves runtime ownership reconstructable
every release binary can be proven from source without explicit provenance data
```

That distinction explains several v1.2.1 design decisions:

- normal write-back defaults to the selected cell rather than inferred broad
  Table scope;
- formulas are preserved unless overwrite is explicitly requested;
- partial writes return structured outcomes;
- normal entry points preserve caller-owned `Application.EnableEvents`;
- a process-visible provider lease prevents two participating current-version
  copies from silently owning the same Excel resources;
- ambiguous provider ownership fails closed;
- native window styling records commit / rollback / recovery state;
- settings namespaces isolate persisted preferences but do not pretend to be a
  runtime ownership mechanism.

---

## 📦 Supported versions

Security fixes are normally applied to the **latest tagged release**.

| Source state | Security support |
|---|---|
| **Latest tagged release** | ✅ Supported |
| `release/*` before publication | ⚠️ Release-candidate testing / best effort |
| `main` | ⚠️ Development branch / best effort |
| Older tagged releases | ❌ Normally unsupported; upgrade first |
| Modified third-party forks/copies | ❌ Unsupported unless the issue reproduces in official source |
| Unofficial binary mirrors | ❌ Unsupported |

When reporting, identify the exact affected state with one of:

```text
release tag
full 40-character commit SHA
```

Do not report only:

```text
latest
current
main from yesterday
```

because those descriptions can point to different source after the report is
submitted.

### Security-fix policy

A confirmed issue may result in:

- a private or controlled fix branch;
- regression coverage;
- release notes / advisory;
- a corrected tagged release;
- guidance to remove or stop using an affected artifact.

Older tags are not normally patched in place.

---

## 📣 Reporting a vulnerability

**Do not open a public GitHub issue for a suspected vulnerability.**

Do not publish:

- exploit code;
- a weaponized workbook;
- credentials;
- a private token;
- a release-signing secret;
- a client workbook;
- sensitive environment details;

in a public issue, pull request, discussion, or Wiki page.

### Option 1 — GitHub private vulnerability reporting

Where the repository has private vulnerability reporting enabled:

```text
Repository
→ Security
→ Report a vulnerability
```

Submit the report there.

### Option 2 — email the maintainer

```text
danielep71@gmail.com
```

Suggested subject:

```text
Private security report — VBA-DATETIMEPICKER
```

### Include, where relevant

- affected release tag or full commit SHA;
- exact file/module/procedure;
- Excel version/build;
- Office 32-bit or 64-bit;
- Windows version;
- embedded-source vs `.xlam` deployment;
- whether another DatePicker provider is loaded;
- version of that second provider;
- settings namespace, if relevant;
- `Application.EnableEvents` state, if relevant;
- worksheet protection / Table / formula state, if relevant;
- whether WinAPI styling is enabled;
- minimal reproduction steps;
- observed behavior;
- expected behavior;
- practical confidentiality, integrity or availability impact;
- whether exploitation requires an already-trusted malicious workbook/macro;
- whether the issue affects official release artifacts;
- any proposed mitigation;
- whether public disclosure has already occurred.

### Safe reproduction material

Prefer:

```text
sanitized workbook
minimal reproduction module
screenshot with sensitive values removed
plain-text steps
regression output
```

Do not attach a production workbook merely because it reproduces the problem.

---

## ⏱️ What to expect

This is a solo-maintained open-source project.

Response times are therefore **best effort**, not a contractual SLA.

The expected process is:

1. acknowledge the report;
2. identify the affected source and environment;
3. reproduce where practical;
4. classify security impact;
5. determine affected versions/artifacts;
6. develop remediation;
7. add regression/recovery tests where applicable;
8. validate on the relevant Excel/Windows host;
9. prepare a corrected release or mitigation;
10. disclose publicly after users have had reasonable time to update.

Credit can be included in release notes or an advisory if the reporter wants it.

Anonymous credit is also acceptable.

Please allow reasonable remediation time before public disclosure.

---

## 🎯 What qualifies as a security issue

When uncertain, report privately.

The maintainer can safely reclassify a report as an ordinary defect.

### 1. Code execution / trust-boundary issues

Examples:

- official DatePicker source or release artifacts execute unexpected code;
- a repository-supplied `.xlsm` / `.xlam` contains undisclosed macros, links,
  connections, embedded payloads or executable behavior;
- untrusted worksheet/user input is used to construct arbitrary macro names,
  `Shell` commands, native API calls, or executable paths in a way the documented
  contract did not intend;
- a callback mechanism can be redirected to attacker-controlled code without the
  host deliberately configuring such a callback;
- a release asset contains code that materially differs from the source claimed
  to produce it.

### 2. Integrity issues

Examples:

- crafted input causes write-back outside the resolved/documented target;
- formula-preservation policy can be bypassed without an explicit destructive
  request;
- `DP_FillTableColumn` writes beyond the confirmed Table data-column scope;
- a partial write is reported as complete success in a way that can cause a
  caller to rely on false state;
- a non-owner DatePicker instance can tear down another provider's shared Excel
  registrations despite the documented `v1.2.1` ownership guard;
- a settings namespace defect causes one deployment to overwrite another
  deployment's persistent configuration unexpectedly;
- a native-window failure leaves the UserForm in an unsafe/unknown state while
  reporting successful application or successful rollback;
- a release or repository action modifies source/artifacts outside the
  documented workflow.

### 3. Confidentiality issues

Examples:

- diagnostics unexpectedly expose worksheet values, file paths, user data or
  secrets beyond documented behavior;
- a released workbook/add-in contains confidential, client-specific,
  machine-specific or personal information;
- a repository workflow logs or publishes secret material;
- the analytics token or another repository credential becomes available to an
  unrelated job;
- a sanitized demonstration path inadvertently captures hidden workbook metadata.

### 4. Availability issues

Examples:

- crafted state causes a persistent Excel hang, repeated crash or uncontrolled
  loop;
- provider ownership can be exploited to prevent normal use of Excel across
  sessions without a practical recovery path;
- a native-window operation repeatedly corrupts the form/process state beyond the
  documented best-effort UI effect;
- timer/event registration creates uncontrolled resource growth;
- a malformed input produces unbounded worksheet writes or pathological repeated
  UI creation.

### 5. Supply-chain issues

Examples:

- tampered GitHub Release assets;
- compromised repository workflow;
- moved/unpinned third-party action changing what executes;
- mismatch between release notes/source tag and published `.xlam` / `.xlsm`;
- compromised download instructions;
- malicious change to source-control or release metadata intended to redirect
  users to an unofficial binary.

### 6. Credential / automation issues

Examples:

- `TRAFFIC_TOKEN` exposure;
- expansion of the analytics workflow so high-privilege analytics credentials
  become available to code-executing build/test jobs;
- workflow permission escalation beyond documented need;
- publishing token-bearing logs;
- use of the analytics credential outside its intended GitHub traffic API scope.

---

## 🐞 Ordinary bugs

A serious defect is not automatically a security vulnerability.

Use a public issue for defects such as:

- the picker opens in the wrong screen position;
- a calendar label is misaligned;
- a date cell is not recognized correctly;
- a right-click entry is missing after startup;
- the grid icon appears late or in the wrong location;
- `Ctrl + Shift + D` is not registered when explicitly enabled;
- a documented formula skip count is wrong but no confidentiality/integrity
  boundary is bypassed;
- an error message is confusing;
- a settings control displays the wrong value;
- the borderless title-bar effect does not apply but Excel remains recoverable;
- performance is slower than expected but bounded;
- documentation is stale.

Some ordinary defects can become security-relevant when they cross a trust,
integrity, confidentiality or availability boundary.

If unsure, report privately first.

---

## 🛠️ Scope

### In scope — production source

- `src/modules/M_DatePicker.bas`;
- `src/classes/cDatePickerManager.cls`;
- `src/classes/cDatePickerLabelHook.cls`;
- `src/forms/UF_DatePicker.frm`;
- `src/forms/UF_DatePicker.frx`;
- `src/ribbon/customUI14.xml`.

### In scope — project validation / demo source

- `test/M_cDP_Test.bas`;
- `demo/M_DEMO_BUILDER.bas`;
- `demo/M_DP_DEMO.bas`.

### In scope — official release artifacts

- official add-in assets published on the Releases page (the exact filename has
  varied between releases — `DATETIMEPICKER v1.2.1.xlam` at `v1.2.1`);
- official `DATETIMEPICKER-demo-v<version>.xlsm` assets;
- release archives;
- release notes and provenance claims.

### In scope — runtime integrations

- worksheet write-back;
- Excel Table target resolution;
- formula-preservation / overwrite policy;
- `Application.EnableEvents` preservation;
- `Application.OnKey` registration/removal;
- `Application.OnTime` timer registration/cancellation;
- DatePicker CommandBar/context-menu integration;
- in-grid worksheet Shape lifecycle;
- settings persistence and namespaces;
- process-visible runtime provider lease;
- provider-ownership guards;
- `DP_ForceReleaseProviderLease`;
- optional WinAPI positioning / borderless styling;
- native transaction rollback / `RecoveryRequired`.

### In scope — repository automation

- `.github/workflows/daily-traffic.yml`;
- `TRAFFIC_TOKEN` handling;
- GitHub Actions permission boundaries;
- traffic-history write behavior;
- automated traffic-alert issue creation;
- third-party action pinning;
- accidental credential disclosure through workflow changes.

### Out of scope

- vulnerabilities in Microsoft Excel, Office, Windows, GitHub or the VBA runtime
  themselves;
- organization-controlled macro-security configuration;
- malicious VBA code not supplied by this repository;
- unrelated add-ins in the host process;
- user-created modifications that do not reproduce in official source;
- unofficial copies or mirrors;
- old unsupported tags where the issue does not affect a supported state;
- social engineering unrelated to project content;
- a user intentionally calling documented destructive APIs with trusted code;
- treating the provider lease as an authorization or sandbox boundary;
- treating settings namespaces as an access-control mechanism;
- ordinary bugs without concrete security impact.

A vulnerability in Excel, Windows or GitHub should be reported to the relevant
vendor/platform.

---

## 🛡️ Runtime security boundaries

### 1. Worksheet write-back

The DatePicker writes into user data.

That is the most important integrity-sensitive runtime function.

Normal behavior favors:

```text
explicit target
predictable scope
formula preservation
structured partial outcome
```

over convenience shortcuts.

#### Single-cell default inside Tables

A selected cell inside an Excel Table data column is not authorization to mutate
the entire column.

Normal interactive selection defaults to the selected/resolved cell.

Whole-column fill is explicit through:

```vb
DP_FillTableColumn
```

A defect that allows an unconfirmed broad write may be security-relevant when it
creates a practical integrity impact.

#### Formula preservation

Formulas are preserved by default.

An ordinary formula can be replaced only through an explicit overwrite-enabled
call where supported.

Array-formula cells remain non-overridable in the normal write engine.

This policy protects against accidental destructive mutation.

It does **not** prevent trusted VBA code from deliberately changing formulas by
other Excel APIs.

#### Structured result

`DP_WriteResult` distinguishes:

```text
attempted
written
locked skips
formula skips
other failures
technical-failure flag, step, number and description
resolved target
Table expansion metadata
caller event state
```

A partial write must not be represented as complete success.

A technical failure stops population rather than continuing to mutate the
workbook, so the result is bounded rather than balanced: it reports the outcomes
actually observed, and the failure is reported in its own fields rather than
inferred from the counts.

---

### 2. `Application.EnableEvents`

Normal DatePicker entry points preserve the caller's `Application.EnableEvents`
state.

This matters because a business macro may deliberately run a transaction with:

```vb
Application.EnableEvents = False
```

The DatePicker must not silently break that transaction by forcing events on.

The explicit recovery exception is:

```vb
DP_RepairRuntime
```

which may deliberately re-enable events.

> [!IMPORTANT]
> `DP_RepairRuntime` is a recovery tool, not a transparent ordinary entry point.

---

### 3. Application-wide keyboard shortcut

The optional DatePicker shortcut uses:

```vb
Application.OnKey
```

That binding belongs to the Excel process.

Excel exposes no supported getter for the previous assignment.

Therefore:

```text
DatePicker can register its own shortcut
DatePicker can remove its own shortcut
DatePicker cannot truthfully restore an unknown third-party predecessor
```

Removal returns the key to Excel default handling.

A third-party binding may already have been displaced when the DatePicker
shortcut was enabled.

The keyboard feature is therefore opt-in and independent from other access paths.

This is valid configuration:

```text
ShowRightClick = False
ShowGridIcon = False
EnableKeyboardShortcut = False
```

The component does not force the key on as a fallback.

> [!NOTE]
> Through `v1.2.0` the settings-panel save path did not honor this: it forced the
> shortcut back on whenever right-click and the grid icon were both disabled, and
> the panel exposes no keyboard control, so the change was invisible to the user
> ([#42](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/42)). The statement above holds from `v1.2.1`.

---

### 4. `Application.OnTime`

The live footer clock uses Excel scheduling.

Timer cleanup must use the exact schedule/procedure identity recorded when the
timer was created.

The DatePicker cannot enumerate all pending `OnTime` jobs globally and should not
pretend to own jobs it did not schedule.

A timer-cleanup defect is security-relevant only when it creates concrete
availability/integrity impact beyond an ordinary UI bug.

---

### 5. Context menu and worksheet Shape ownership

The component uses stable identifiers for its own UI artifacts.

It must remove only DatePicker-owned resources.

A cleanup routine must not delete unrelated:

```text
CommandBar controls
worksheet Shapes
application registrations
```

merely because they look similar.

The in-grid icon's retained object reference is validated for liveness before
reuse because deleting a worksheet/shape can stale the reference.

---

## 🪪 Runtime provider lease

`v1.2.0` introduces a process-visible provider lease because two DatePicker
copies share Excel surfaces. `v1.2.1` completes it: admission is enforced on
every runtime entry path, not only at `DP_Start`.

Lease name:

```text
__VBA_DATETIMEPICKER_RUNTIME_PROVIDER_LEASE__
```

The lease uses a hidden temporary CommandBar and a hidden marker control carrying
an ephemeral owner token.

### Security/safety goals

The lease is intended to prevent:

```text
provider A registers shared resources
provider B silently replaces them
provider A or B tears down resources owned by the other
```

### Core rules

**Acquire before registration**

The lease must be claimed before shared registrations are changed.

**Guard teardown**

`DP_Stop` and `DP_RepairRuntime` verify ownership before destructive lifecycle
work.

**Prove ownership before release**

A provider needs its local token and a matching process-visible marker.

**Ambiguous ownership fails closed**

Unreadable ownership is treated as occupied/ambiguous, not as free.

**Lease lifetime is process-scoped**

The lease is temporary and should disappear when Excel exits.

### Stale lease after VBA reset

A VBA project reset can erase the local token while leaving the process-visible
lease alive.

That state intentionally fails closed.

Safest recovery:

```text
close all Excel windows
restart Excel
```

### Operator force release

```vb
DP_ForceReleaseProviderLease
```

is intentionally more powerful.

It deletes the lease without proving ownership.

> [!CAUTION]
> This is an operator recovery command.
>
> It must never be called automatically.
>
> Use it only when the operator has established that no other DatePicker provider
> is alive.

### Not a security sandbox

The lease prevents accidental collision among participating implementations.

It does **not** protect against malicious VBA running with the same Excel user
privileges.

A malicious macro can use Excel/VBA/Office APIs directly.

---

## 🔀 Mixed-version sessions

The provider lease protects:

```text
v1.2.1 + v1.2.1
```

when both copies participate in the protocol on every entry path.

It cannot retrofit lease awareness into a previously released provider.

Therefore:

```text
v1.2.1 + pre-v1.2.1
```

is not guaranteed safe in either direction.

`v1.2.0` is a pre-`v1.2.1` provider for this purpose, not a participating peer.
It carries the lease protocol but admits it only at `DP_Start`, so on every other
entry path it behaves as a non-participant while appearing to be one
([#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37)). That is the more dangerous of the two failure modes,
because it looks protected.

This is a compatibility boundary, not an authentication failure.

For controlled deployments, remove/disable the older provider before relying on
`v1.2.1` ownership behavior. Do not rely on `v1.2.0` ownership behavior: it
admits the lease only at `DP_Start` ([#37](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/37)).

---

## ⚙️ Settings persistence and namespaces

Settings use VBA registry persistence under the stable legacy base name:

```text
VBA_DATETIMEPICKER
```

A deployment may select a stable namespace before first settings load.

Example:

```vb
M_Settings_SetNamespace "TreasuryTool"
DP_Start
```

A namespaced effective application name is derived from the stable base.

### Security-relevant properties

- namespace resolution is centralized;
- the namespace locks after settings load;
- no automatic migration copies the shared scope into a new namespace;
- namespaces are persistent deployment identities;
- namespaces are not runtime-provider identities.

### Not an access-control mechanism

A settings namespace is designed to prevent accidental preference coupling.

It does not encrypt settings.

It does not protect secrets.

Do not store credentials, API keys, passwords, private tokens or client secrets in
DatePicker settings.

---

## 🪟 WinAPI and native UserForm styling

Optional DatePicker UI behavior uses Windows APIs for form positioning and
borderless styling.

Security-sensitive review is required for changes to:

- window handle resolution;
- `FindWindow`;
- `GetWindowRect`;
- `GetWindowLong` / `GetWindowLongPtr`;
- `SetWindowLong` / `SetWindowLongPtr`;
- `SetWindowPos`;
- `DrawMenuBar`;
- `SendMessage`;
- style masks such as `WS_CAPTION`;
- error handling around ambiguous native return values.

### Transactional styling

Borderless styling is treated as a transaction.

Conceptually:

```text
read current style
write target style
refresh frame
redraw
```

If a post-commit step fails, rollback is attempted.

`DP_WindowStyleResult` reports:

```text
Attempted
Applied
Committed
RolledBack
RecoveryRequired
FailedStep
LastApiError
```

A half-applied style must not silently report clean success.

From `v1.2.1`, `RecoveryRequired` is also terminal for that form load: the picker
fails rather than presenting a window whose style is neither fully applied nor
rolled back. Through `v1.2.0` the result was produced and then discarded at both
call sites, so the form was presented anyway ([#47](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/47)).

### Native APIs are not an elevation mechanism

The APIs operate inside the user's existing Excel process and Windows session.

They do not grant higher operating-system privileges.

A vulnerability claim should distinguish:

```text
UI corruption / recoverability problem
```

from:

```text
privilege escalation / arbitrary code execution
```

---

## 🧩 RibbonX

Optional RibbonX is package metadata:

```text
src/ribbon/customUI14.xml
```

The XML references DatePicker callback procedures.

Security-sensitive changes include:

- adding callbacks;
- changing callback names to dynamically constructed values;
- adding external image/resource references;
- adding commands that invoke unexpected macros;
- packaging Ribbon XML into a release binary that does not match source review.

RibbonX should remain a thin interface into documented DatePicker entry points.

---

## 📦 Macro-enabled release artifacts

The repository is source-first.

Generated:

```text
DATETIMEPICKER v1.2.1.xlam
DATETIMEPICKER-demo-v1.2.1.xlsm
```

are executable Office artifacts and must be treated accordingly.

> [!WARNING]
> The `.xlam` asset separator has changed between releases — `v1.2.0` published
> `DATETIMEPICKER.v1.2.0.xlam` and `v1.2.1` publishes `DATETIMEPICKER v1.2.1.xlam`.
> Do not treat a filename as proof of authenticity. Confirm the exact name and
> the SHA-256 on the Release page.

Users should assume:

```text
macro-enabled Office file = executable content
```

even when the corresponding source repository is transparent.

Official release binaries are in scope for security reports.

---

## 🔗 Supply-chain and release integrity

### Trusted distribution

Obtain source and binaries from the official GitHub repository / Releases page.

Do not rely on:

- third-party mirrors;
- files forwarded by email without provenance;
- renamed binaries from unrelated repositories;
- binaries embedded in blog downloads;
- unofficial package sites.

### Source-first review

Where organizational policy requires it, review:

```text
.bas
.cls
.frm
Ribbon XML
relevant release metadata
```

before enabling the macro-enabled artifact.

Remember that `.frx` is a binary UserForm resource companion and must travel with
its matching `.frm`.

### Exact artifact provenance

`v1.2.1` publishes a SHA-256 for each release asset, recorded after the asset
was built **and tested**. That establishes **file identity**: it lets you confirm
the file you hold is the file that was certified. It does not establish that the
file was built from a given source commit.

The repository is continuing to strengthen exact-SHA release manifests and
machine-readable evidence. Provenance today is procedural, not cryptographic. Do
not overstate what a source regression proves about a later saved `.xlam` or
`.xlsm`.

This distinction matters:

```text
source at SHA X passed regression
```

is not automatically identical to:

```text
binary asset Y is cryptographically proven to have been built from SHA X
```

Release-provenance work is tracked separately from source behavior, as
[#16](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/16).

---

## 🔍 Verifying a release

For controlled use:

1. obtain the source/release asset from the official repository;
2. record the release tag;
3. record the relevant commit SHA where practical;
4. inspect release notes and `CHANGELOG.md`;
5. review the source relevant to the deployment;
6. treat `.xlam` / `.xlsm` as executable Office content;
7. compute the artifact's SHA-256 and compare it with the value published on
   the Release page:

   ```text
   certutil -hashfile "DATETIMEPICKER v1.2.1.xlam" SHA256
   ```

   A mismatch means the file is not the certified artifact. Treat it as a
   supply-chain report, not a download error.
8. scan binaries under organizational policy where required;
9. compile the imported source:

   ```text
   VBA Editor → Debug → Compile VBAProject
   ```

10. for developer validation, import the regression dependencies and run:

   ```vb
   TST_DP_RunAll
   ```

11. for form/UI changes, also run:

   ```vb
   TST_DP_RunAll_WithUISmoke
   ```

12. record the actual Excel version/build, Office bitness and Windows version;
13. perform package-level smoke testing on the actual release binary where that
    binary is the deployed artifact.

Latest recorded `v1.2.1` regression baseline, on both the embedded `.xlsm` and
the packaged `.xlam`:

```text
standard:       State=PASS; Run=431; Passed=431; Failed=0; CleanupFailures=0
with UI smoke:  State=PASS; Run=434; Passed=434; Failed=0; CleanupFailures=0
```

The assertion count is informative, not a permanent security constant.

Certification ran on a developer workstation with other add-ins loaded in the
same Excel process, not on a clean VM. Host-isolation effects are therefore not
part of the recorded evidence — which matters here, because the provider lease
and the `OnKey`/`OnTime` boundaries above are all about behavior in a shared
process.

Only a clean `PASS` is a passing harness run.

---

## 🧰 Safe-use guidance

### 1. Preserve macro security

- keep Excel macro security at the organization-approved level;
- do not disable Protected View globally;
- do not weaken Trust Center settings solely to use the DatePicker;
- use Trusted Locations or signed VBA only where organizational policy supports
  them;
- unblock a downloaded file only after establishing provenance.

### 2. Prefer embedded source in locked-down environments

Embedded source can reduce deployment complexity because recipients do not need a
separately installed add-in.

It does not make macros inherently safe.

The workbook still contains executable VBA and must be trusted as such.

### 3. Use one provider per Excel process

For v1.2.1, choose:

```text
embedded provider
```

or:

```text
.xlam provider
```

for a given Excel session.

Do not intentionally create a multi-provider setup and then force-release leases
until the conflict disappears.

### 4. Use stable settings namespaces where required

A namespace can prevent accidental preference sharing across deployments.

Do not store secrets there.

### 5. Treat broad write-back as explicit

Use:

```vb
DP_FillTableColumn
```

only when whole-column write scope is intended.

Review `DP_WriteResult` when programmatic callers need to know exactly what
happened.

### 6. Respect formula ownership

Normal write-back preserves formulas.

Use formula overwrite only when the caller deliberately owns that destructive
decision.

### 7. Keep a runtime recovery path

Know the distinction:

```text
DP_Close          close form/UI
DP_Stop           normal owner teardown
DP_RepairRuntime  deliberate runtime repair; may re-enable events
Excel restart     safest process-level reset
DP_ForceReleaseProviderLease
                  operator-only lease recovery
```

Workbook teardown must go through `DP_Stop`. The low-level helpers —
`M_ContextMenu_Remove`, `M_KeyboardShortcut_Remove`, `M_Timer_Stop`,
`M_GridIcon_PurgeAll` — are internal and diagnostic. They perform no ownership
check, because they are the internals the ownership-aware APIs call once
admission has already been proven.

Calling them directly from `Workbook_BeforeClose` strands the lease for an owner,
and lets a **refused** non-owner strip the true owner's shortcut, menu entry and
icons — exactly the damage the lease exists to prevent. A teardown recipe of that
shape was published through `v1.2.0` and is corrected in `v1.2.1`
([#17](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/17)).

### 8. Protect sensitive material in bug reports

Sanitize:

- screenshots;
- workbook names;
- worksheet values;
- external links;
- client names;
- file-system paths;
- registry screenshots;
- logs.

---

## 🔑 Repository automation credentials

The repository contains one GitHub Actions workflow:

```text
.github/workflows/daily-traffic.yml
```

Its purpose is repository analytics.

It is **not** software CI and it does not execute the DatePicker VBA tests.

### Workflow purpose

It:

- reads GitHub traffic data;
- appends aggregate history to the public orphan `traffic-history` branch;
- can create an alert issue for traffic spikes/new referrers.

### Credential model

| Control | Current setting |
|---|---|
| Secret | `TRAFFIC_TOKEN` |
| Secret type | Fine-grained personal access token |
| Token permission | `Administration: read` for this repository |
| GitHub environment | `analytics` |
| Workflow permissions | `contents: write`, `issues: write` |
| Third-party action | `actions/checkout` pinned to a full commit SHA |
| Traffic-history data | Aggregate repository analytics; public branch |
| Alert issue token | `${{ github.token }}` under the workflow's `issues: write` permission |

`TRAFFIC_TOKEN` is deliberately isolated in the `analytics` environment.

### Why the token is sensitive

GitHub's traffic API requires a high-value repository permission relative to the
simple read-only analytics task.

That credential should never be exposed to a job that:

```text
builds untrusted pull-request code
executes repository scripts supplied by contributors
runs future test tooling
packages release binaries
```

Keeping analytics and software-quality automation separate reduces the blast
radius of a compromised build/test path.

### Rotation policy

Rotate `TRAFFIC_TOKEN` at least every 90 days and immediately if:

- the workflow is modified unexpectedly;
- the token is printed or otherwise exposed;
- the token is used from another workflow;
- the `analytics` environment protection changes materially;
- repository access is broadened unexpectedly;
- the workflow permissions are widened;
- a collaborator with write access is added and the threat model changes;
- GitHub reports suspected token compromise.

### Workflow-permission review

The current workflow needs:

```yaml
permissions:
  contents: write
  issues: write
```

because it:

```text
pushes traffic history
opens alert issues
```

Do not broaden to repository-wide write permissions without explicit review.

### Third-party actions

`actions/checkout` is pinned to a full commit SHA.

Future third-party actions should also be pinned to immutable full SHAs rather
than floating tags.

### Traffic data privacy

The workflow stores aggregate data such as:

```text
views
unique views
clones
unique clones
referrers
popular paths
stars
forks
watchers
open issue counts
```

The `traffic-history` branch is public.

The workflow is intended to preserve the same type of aggregate repository
analytics exposed by GitHub, not visitor-identifying data.

If future API changes introduce personally identifying fields, the workflow must
not publish them without a separate privacy/security review.

### Keep analytics separate from quality automation

Software-quality automation is a separate concern.

Do not extend the traffic workflow into:

```text
VBA build
test execution
static source checks
release packaging
artifact signing
```

while the analytics credential is available to the job.

The repository's software-quality workflow work is tracked separately, as
[#15](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/15). No automated CI runs the VBA regression pack today.

---

## 🔐 Secret-handling rules

Never commit:

```text
TRAFFIC_TOKEN
personal access tokens
GitHub tokens
private signing keys
PFX/P12/PVK files
passwords
API credentials
client secrets
```

The repository `.gitignore` excludes common local environment and private-key
material, but `.gitignore` is not a security control.

A secret committed once must be considered compromised even if the commit is
later deleted.

If a credential is exposed:

1. revoke/rotate it immediately;
2. determine scope of access;
3. remove it from current source;
4. assess whether history cleanup is useful;
5. assume copies may exist in clones/logs;
6. review related workflow activity.

---

## 🧾 Logging and diagnostics

Diagnostics should reveal enough to reproduce a problem without exposing
unnecessary user data.

Prefer logging:

```text
procedure/stage
error number
error description
target address where operationally necessary
counts
host/version metadata
```

Avoid logging:

```text
entire worksheet contents
credentials
connection strings
private registry values unrelated to DatePicker
personal information
arbitrary workbook dumps
```

`DP_WriteResult` address lists are designed as operational diagnostics.

A consuming application that handles sensitive sheet structure should decide
whether to persist or externally transmit those addresses.

---

## 🧪 Regression harness security considerations

The regression harness manipulates real Excel state.

It can:

- create/delete test worksheets;
- change DatePicker settings temporarily;
- manipulate application-level registrations;
- exercise UserForm/native-window paths;
- use the same process-level runtime lease model as production.

Run it in a controlled workbook/Excel instance.

Do not run release-validation tests in an unsaved production workbook containing
irreplaceable data.

### Dirty-start protection

The harness refuses to call a contaminated predecessor environment a pass.

Possible state includes a leftover:

```text
TST_DP_SCRATCH
```

worksheet.

`FAIL_DIRTY_START` means the environment was not trusted as a clean basis for
evidence.

That is a safety feature, not a vulnerability.

### Worksheet-creation recovery

The harness contains recovery logic for an observed partial-success
`Worksheets.Add` failure.

It does not blindly retry when worksheet ownership is ambiguous.

That fail-closed behavior protects against deleting a worksheet the harness
cannot prove it created.

---

## 📣 Disclosure coordination

Please avoid public disclosure while:

- exploitability is still being assessed;
- a release fix is being prepared;
- users have not had reasonable time to update;
- a repository credential is still active;
- a malicious release artifact remains publicly downloadable.

The maintainer may ask for:

- additional environment detail;
- a sanitized reproduction;
- confirmation against a candidate fix;
- a reasonable embargo period.

The project does not require a reporter to surrender ownership of their research.

The goal is simply to reduce preventable user harm.

---

## 🧭 Security review checklist for maintainers

For a security-sensitive DatePicker change, review:

```text
[ ] Trust boundary stated
[ ] Exact input / caller control identified
[ ] Write scope assessed
[ ] Formula-destructive behavior assessed
[ ] Application.EnableEvents ownership assessed
[ ] OnKey / OnTime process-wide effects assessed
[ ] Context-menu / Shape ownership assessed
[ ] Provider lease ownership assessed
[ ] Mixed-version behavior stated
[ ] Settings persistence/namespace behavior assessed
[ ] WinAPI handle/style behavior assessed
[ ] Error path cannot replace original evidence
[ ] Partial success is observable
[ ] Recovery behavior documented
[ ] 32/64-bit path considered
[ ] Regression added where deterministic
[ ] Manual host validation recorded where required
[ ] Documentation updated
[ ] Release-artifact impact assessed
[ ] Credential/workflow impact assessed
```

A security review should distinguish:

```text
code correctness
operational safety
security impact
```

They overlap, but they are not identical.

---

## 📚 Related policies and documentation

- [Project README](README.md)
- [Installation Guide](INSTALLATION.md)
- [Contributing Guidelines](CONTRIBUTING.md)
- [Code of Conduct](CODE_OF_CONDUCT.md)
- [Changelog](CHANGELOG.md)
- [MIT License](LICENSE)
- [Project Wiki](https://github.com/danielep71/VBA-DATETIMEPICKER/wiki)

---

## 👤 Maintainer

Maintained by **Daniele Penza**.

Private security reports:

```text
danielep71@gmail.com
```

---

<div align="center">

## 🛡️ Security principle

**Trust the source you run. Keep destructive scope explicit. Fail closed when ownership is ambiguous. Treat release binaries and repository credentials as security-sensitive artifacts.**

</div>
