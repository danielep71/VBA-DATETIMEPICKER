<div align="center">

# 🧭 Code of Conduct

### Respectful, evidence-led technical collaboration for VBA-DATETIMEPICKER

[![Applies to](https://img.shields.io/badge/Applies_to-Everyone-217346?style=for-the-badge)](#scope)
[![Spaces](https://img.shields.io/badge/Spaces-Issues_PRs_Wiki-0969da?style=for-the-badge)](#scope)
[![Standard](https://img.shields.io/badge/Standard-Respectful_%2B_Evidence--Led-6f42c1?style=for-the-badge)](#technical-discussion-standards)
[![Enforcement](https://img.shields.io/badge/Enforcement-Maintainer-d97706?style=for-the-badge)](#enforcement)

<br>

**Technical rigor · Respectful disagreement · Reproducible evidence · Privacy-aware collaboration**

</div>

---

**VBA-DATETIMEPICKER** is a focused open-source Excel/VBA project.

This Code of Conduct exists to keep interaction around the project respectful,
technical, constructive, and welcoming — especially when a problem is difficult
to reproduce, depends on Excel session state, or exposes a limitation in the
platform itself.

People should feel comfortable:

- reporting defects;
- asking basic or advanced VBA questions;
- challenging design decisions;
- proposing safer alternatives;
- documenting behavior that differs by Office version, bitness, workbook state,
  or Windows configuration;
- saying that an earlier assumption was wrong;
- contributing even when they are unfamiliar with the project's conventions.

A technically demanding project benefits from disagreement.

It does not benefit from hostility.

---

<a id="our-pledge"></a>

## 🤝 Our pledge

Everyone who participates — by opening an issue, submitting a pull request,
commenting, reviewing, editing documentation or the Wiki, discussing a release,
or representing the project elsewhere — is expected to help create a
harassment-free experience for all.

That expectation applies regardless of:

- experience level;
- professional or academic background;
- age;
- disability;
- ethnicity;
- gender identity or expression;
- nationality;
- race;
- religion;
- sexual orientation;
- socioeconomic status;
- or any other personal characteristic unrelated to the technical contribution.

Technical rigor and respectful interaction are complementary requirements.

Neither excuses the absence of the other.

---

<a id="expected-behavior"></a>

## ✅ Expected behavior

Participants are expected to:

- be respectful and assume good faith unless evidence shows otherwise;
- focus criticism on code, behavior, documentation, tests, architecture, or
  process rather than on the person who produced them;
- describe Excel/VBA behavior precisely;
- distinguish **observed fact**, **inference**, **hypothesis**, and **platform
  limitation**;
- provide reproduction steps, logs, screenshots, test evidence, or minimal
  examples where practical;
- acknowledge uncertainty rather than presenting an assumption as verified;
- correct mistakes openly when new evidence changes the conclusion;
- give and receive constructive review comments professionally;
- respect privacy, confidentiality, client restrictions, and security boundaries;
- help newcomers understand the repository's workflow and vocabulary;
- allow maintainers reasonable time to investigate Excel behavior that may be
  session-specific or difficult to reproduce;
- respect that a contribution may be adopted, adapted, deferred, split into a
  follow-up issue, or declined to preserve project coherence and compatibility.

### Useful disagreement

A useful technical disagreement is specific enough that another person can
investigate it:

> "`Worksheets.Add` reported error 1004 after the workbook's worksheet count
> increased from 3 to 4. I reproduced the failure before any rename or formatting
> step. Host details and the exact regression output are below."

That can be tested.

### Unhelpful disagreement

A personal judgment cannot be tested:

> "This code is broken because the author does not understand Excel."

Both may arise from frustration with the same defect.

Only the first helps fix it.

---

<a id="unacceptable-behavior"></a>

## 🚫 Unacceptable behavior

Unacceptable behavior includes:

- personal attacks, insults, ridicule, or derogatory comments;
- harassment in public or private;
- discriminatory, demeaning, or sexualized language or imagery;
- threats, intimidation, or encouragement of violence;
- publishing another person's private information without permission;
- deliberate misrepresentation of another participant's work or statements;
- knowingly fabricating test results, screenshots, reproduction evidence, or
  implementation claims;
- sustained disruption of technical discussion;
- repeated bad-faith argument after the technical decision and evidence have
  been explained;
- spam, unrelated promotion, or commercial solicitation;
- attempts to pressure maintainers into unsafe disclosure or an unverified
  release;
- public disclosure of a suspected vulnerability before reasonable coordinated
  remediation;
- retaliation against someone who reports misconduct, a security concern, or a
  technical failure.

Disagreement is allowed.

Abuse is not.

---

<a id="technical-discussion-standards"></a>

## 🧪 Technical discussion standards

VBA-DATETIMEPICKER interacts with several Excel surfaces whose behavior can
depend on the host session:

```text
Application events
Application.OnKey
Application.OnTime
CommandBars
worksheet Shapes
UserForms
Excel Tables
protected worksheets
registry-backed settings
optional WinAPI
```

A report is therefore more useful when the environment is explicit.

### For bugs and behavioral reports

Where relevant, include:

- exact repository tag, branch, or commit SHA;
- Excel version/build;
- Office 32-bit or 64-bit;
- Windows version;
- whether the code is:
  - embedded in a workbook;
  - running from the `.xlam`;
  - running alongside another DatePicker copy;
- whether any pre-`v1.2.1` DatePicker provider is also loaded (a `v1.2.0` peer
  admits the lease only at `DP_Start`, so it counts);
- `Application.EnableEvents` state when relevant;
- worksheet protection state;
- whether WinAPI styling is enabled;
- whether right-click, grid-icon, and keyboard access are enabled;
- the selected cell/range and whether it is inside an Excel Table;
- reproduction steps;
- observed behavior;
- expected behavior;
- logs, screenshots, or minimal workbooks when safe to share.

Do not manufacture certainty where Excel exposes no readable state.

Examples include:

```text
Application.OnKey predecessor
pending Application.OnTime schedule enumeration
Excel-internal cause of a partial-success Worksheets.Add
```

Where the platform makes a state unobservable, say so and classify the evidence
appropriately:

```text
automated regression
manual validation
source inspection
platform contract
unresolved hypothesis
```

That distinction is part of this project's quality standard.

---

## ✅ Regression evidence

The current regression harness is:

```text
test/M_cDP_Test.bas
```

Primary runners:

```vb
TST_DP_RunAll
TST_DP_RunAll_WithUISmoke
```

For production changes, contributors should normally compile first:

```text
VBA Editor → Debug → Compile VBAProject
```

and then report the harness summary.

The harness uses these run states:

```text
PASS
FAIL
FAIL_CLEANUP
FAIL_DIRTY_START
INCOMPLETE_SKIPPED
```

A run that ends early prints no summary line at all. An absent summary is itself
a failure result, not missing output.

A statement such as:

```text
302 assertions passed
```

is incomplete without the run state and cleanup result.

Preferred evidence:

```text
State=PASS; Run=431; Passed=431; Failed=0; CleanupFailures=0
```

A non-`PASS` result must not be presented as release-quality evidence merely
because some or all assertions passed.

When the form itself changed, include the UI-smoke run where practical.

---

## 🔍 Review standards

Review comments should be actionable whenever possible.

A strong review comment identifies:

1. **where** the concern exists;
2. **what** behavior or invariant is at risk;
3. **why** it matters;
4. **what evidence** supports the concern;
5. whether the requested change is:
   - required;
   - recommended;
   - optional;
6. whether the concern affects:
   - correctness;
   - compatibility;
   - data safety;
   - runtime ownership;
   - diagnostics;
   - test validity;
   - documentation;
   - security;
   - maintainability;
   - style.

Example:

> "`M_Settings_SetNamespace` changes externally visible persistence behavior and
> is required for callers to select isolated settings. It should therefore be
> classified as supported API under #25 rather than internal."

This is preferable to:

> "I don't like this API."

---

## 🛡️ Data-safety discussions

This component can write into user worksheets.

Claims about write behavior deserve particular care.

When discussing write-back changes, distinguish at least:

```text
target resolution
selected-cell vs table-column scope
literal values
formula cells
protected locked cells
array-formula cells
partial writes
explicit destructive overrides
```

Do not describe a destructive behavior as "convenient" without also stating its
scope and safety implications.

The project deliberately prefers:

```text
explicit destructive action
```

over:

```text
implicit broad write
```

and:

```text
reported partial outcome
```

over:

```text
silent partial success
```

Technical discussion should preserve those distinctions.

---

## 🔐 Runtime-ownership discussions

Some Excel integration surfaces are application-wide.

The `v1.2.1` contract allows one current-version DatePicker provider to own the
session and refuses a second one on every entry path.

When reporting ownership problems, identify:

```text
provider A version
provider B version
embedded vs add-in
which provider started first
which provider called DP_Stop / DP_RepairRuntime
whether a VBA project reset occurred
whether Excel was restarted
```

Do not assume a mixed-version session is protected by the current provider
lease.

A pre-`v1.2.0` copy does not participate in that protocol at all, and a `v1.2.0`
copy participates only at `DP_Start`.

---

<a id="privacy-and-confidentiality"></a>

## 🔒 Privacy, confidentiality, and safe reproductions

Do not upload confidential business material merely to demonstrate a DatePicker
problem.

That includes:

- client workbooks;
- proprietary VBA;
- credentials;
- private signing keys;
- personal data;
- internal URLs;
- connection strings;
- production data extracts;
- non-public add-ins or modules you are not authorized to distribute.

Create a sanitized minimal reproduction instead.

Excel files can contain much more than visible worksheet values, including:

```text
defined names
external links
connection strings
Power Query metadata
cached values
comments
document properties
hidden sheets
VBA
```

A workbook that appears anonymized may still disclose information elsewhere in
the package.

If a private reproduction is genuinely necessary, coordinate privately with the
maintainer before sharing it.

---

## 🔑 Security issues are different from ordinary bugs

Suspected vulnerabilities must follow:

[SECURITY.md](SECURITY.md)

Do **not** publish exploit details, credentials, or a proof of concept in a normal
public issue before coordinated disclosure.

The Code of Conduct reporting channel is for participant behavior.

The Security Policy is for software vulnerabilities.

If an incident involves both, use the private channel and say so.

---

<a id="scope"></a>

## 🛠️ Scope

This Code of Conduct applies to:

- this GitHub repository;
- issues;
- pull requests;
- review comments;
- GitHub Discussions, if enabled;
- release threads;
- the Wiki;
- project-related email;
- project-related private communication between participants;
- public spaces where someone is representing the project.

Examples of project representation include:

- speaking on behalf of the project;
- using an official project account;
- presenting oneself as a maintainer or contributor in a project-related forum;
- moderating a project discussion.

The standards apply to both maintainers and contributors.

---

<a id="reporting-unacceptable-behavior"></a>

## 📣 Reporting unacceptable behavior

Report unacceptable behavior **privately** to the maintainer:

```text
danielep71@gmail.com
```

Where available, include:

- what happened;
- where it happened;
- dates or approximate times;
- links, screenshots, or quoted text;
- whether the behavior is ongoing;
- whether another participant witnessed it;
- any immediate safety, privacy, or confidentiality concern.

Do not post sensitive personal information in a public issue.

Reports will be handled as discreetly as reasonably possible.

Information will be shared only as needed to:

```text
understand the report
protect participants
enforce this policy
comply with platform or legal requirements where applicable
```

A good-faith report will not be treated as misconduct merely because the
maintainer ultimately concludes that no violation occurred.

---

<a id="enforcement"></a>

## ⚖️ Enforcement

The maintainer is responsible for interpreting and enforcing this Code of
Conduct.

Responses depend on seriousness, frequency, context, prior behavior, and risk to
participants or the project.

Possible actions include:

1. clarification or a private reminder;
2. a formal warning;
3. editing or removing comments or contributions;
4. closing or locking a discussion;
5. rejecting or reverting a contribution;
6. temporary restriction from project participation;
7. permanent blocking;
8. escalation to GitHub or another relevant platform.

Enforcement aims to be:

```text
proportionate
consistent
documented where appropriate
protective of participants
protective of the technical record
```

Retaliation against a reporter, witness, or participant in an investigation is
itself a violation.

---

## 🧩 Conflicts of interest

Participants should disclose a material conflict when it could reasonably affect
technical review.

Examples:

- ownership of a competing implementation;
- commercial interest in a dependency or integration being proposed;
- employment/client restrictions that materially limit what can be disclosed;
- inability to establish the origin or license of copied code;
- evaluating a contribution one personally authored under another identity or
  organization.

A conflict is not automatically disqualifying.

Undisclosed material influence is the problem.

---

## 📜 Source and licensing integrity

Contributors must have the right to submit the code, documentation, screenshots,
or other material they provide.

Do not submit:

- proprietary code copied from an employer or client;
- code with an incompatible license without clear attribution and discussion;
- screenshots containing confidential data;
- generated code presented as independently authored when its provenance or
  license is uncertain;
- binary workbook changes that bypass the repository's source-first architecture.

If material was adapted from another source, identify that source and its license
clearly enough for review.

---

## 🧱 Maintainer decisions

A maintainer may decline a contribution even when it is technically valid.

Reasons can include:

```text
scope
compatibility
maintenance burden
testability
platform risk
API stability
release timing
duplication
architectural direction
```

A declined contribution is not a judgment about the contributor.

When practical, the technical reason should be recorded so that the same design
question does not need to be rediscovered repeatedly.

Participants may challenge a decision respectfully with new evidence.

Repeatedly reopening the same argument without new evidence is not constructive.

---

## 🙏 Project scale and response expectations

VBA-DATETIMEPICKER is maintained by one person.

That affects response capacity, not the seriousness of this policy.

Response times are best-effort.

Complex reports may take longer when they require:

- a particular Office bitness;
- a clean Excel process;
- multiple VBA providers;
- project-reset behavior;
- protected-sheet state;
- WinAPI testing;
- a fresh-session reproduction.

Reasonable delay is not dismissal.

Repeatedly demanding immediate action is not an acceptable substitute for
technical evidence.

Where appropriate, GitHub's platform policies and community standards also
apply.

---

## 🧭 Practical principle

The standard for this project can be summarized as:

```text
be precise about the software
be generous toward the person
show the evidence
state the uncertainty
protect user data
```

That is the environment in which difficult Excel/VBA problems are most likely to
be solved well.

---

## 👤 Maintainer

Maintained by **Daniele Penza**.
