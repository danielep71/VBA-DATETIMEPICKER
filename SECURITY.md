# 🔒 Security Policy

<p align="left">
  <img alt="Reporting" src="https://img.shields.io/badge/reporting-private-orange">
  <img alt="Scope" src="https://img.shields.io/badge/scope-VBA_source_and_add--in-6f42c1">
  <img alt="Platform" src="https://img.shields.io/badge/platform-Excel_VBA-blue">
</p>

This project ships VBA source and a packaged Excel add-in (`DATETIMEPICKER.xlam`)
that run with the user's Excel and Windows privileges. Security reports are taken
seriously even though the surface is small. This page explains what is supported,
how to report an issue privately, and what is in scope.

---

## 📦 Supported versions

<p align="left">
  <img alt="Support" src="https://img.shields.io/badge/Support-Latest_release-217346">
</p>

Security fixes are applied to the latest release. Older tags are not patched —
please upgrade before reporting.

| Version | Supported |
| --- | --- |
| `v1.1.0` (latest) | ✅ |
| `v1.0.0` | ⚠️ upgrade recommended |
| earlier / unreleased | ❌ |

---

## 📣 Reporting a vulnerability

<p align="left">
  <img alt="Channel" src="https://img.shields.io/badge/Channel-Private_disclosure-orange">
  <img alt="Public_issues" src="https://img.shields.io/badge/Public_issues-Do_not_use-red">
</p>

**Please do not open a public issue for a security problem**, and do not post
proof-of-concept exploit code in a public thread.

Report privately through either:

- **GitHub private vulnerability reporting** — on the repository, go to the
  **Security** tab → **Report a vulnerability** (enable it under
  *Settings → Security* if it is not already on), or
- email the maintainer at:

```text
danielep71@gmail.com
```

Helpful details to include:

- Affected version (`v1.1.0`, etc.) and Excel version / bitness / OS
- A clear description of the issue and its impact
- Minimal reproduction steps
- Any suggested remediation, if you have one

---

## ⏱️ What to expect

<p align="left">
  <img alt="Response" src="https://img.shields.io/badge/response-best_effort-blue">
</p>

This is a solo-maintained project, so responses are best-effort rather than
guaranteed within a fixed window. You can expect:

- Acknowledgement of your report,
- An assessment of validity and severity,
- A fix in a new release when a valid issue is confirmed,
- Credit in the release notes if you would like it.

Please allow reasonable time for a fix before any public disclosure.

---

## 🎯 Scope

<p align="left">
  <img alt="In_scope" src="https://img.shields.io/badge/In_scope-This_project-217346">
  <img alt="Out_of_scope" src="https://img.shields.io/badge/Out_of_scope-Host_environment-orange">
</p>

**In scope**

- The VBA source under `src/` (module, classes, form, ribbon)
- The published `DATETIMEPICKER.xlam` add-in
- Registry use under the `VBA_DATETIMEPICKER` app name
- The WinAPI declarations used for window styling and positioning
- Worksheet write-back behavior

**Out of scope**

- Microsoft Excel, Office, Windows, or the VBA runtime themselves
- Issues that require the user to disable Excel macro security or run untrusted
  code unrelated to this project
- Copies of the add-in obtained from anywhere other than the official Releases
  page

---

## 🧰 Safe use guidance

<p align="left">
  <img alt="Trust" src="https://img.shields.io/badge/Download-Official_releases_only-217346">
</p>

- Download `DATETIMEPICKER.xlam` only from this repository's official
  [Releases](https://github.com/danielep71/vba-datetimepicker/releases) page.
- After downloading, right-click the file → **Properties** → **Unblock**, then
  enable it through the Excel Add-ins manager.
- Review the source before importing if your environment requires it — all code
  is published in plain text under `src/`.

---

## 🔑 Repository automation credentials

<p align="left">
  <img alt="Scope" src="https://img.shields.io/badge/Secret-TRAFFIC__TOKEN-d73a49">
  <img alt="Rotation" src="https://img.shields.io/badge/Rotate-every_90_days-orange">
</p>

One workflow runs against this repository: `.github/workflows/daily-traffic.yml`.
It reads the GitHub traffic API and appends a daily snapshot to CSV files on the
orphan `traffic-history` branch.

| Control | Setting |
|---|---|
| Secret | `TRAFFIC_TOKEN` |
| Type | Fine-grained personal access token |
| Permission | `Administration: read` on this repository only |
| Environment | `analytics` — not readable by any other workflow |
| Workflow permissions | `contents: write` only |
| Third-party actions | Pinned to full commit SHAs |

**Rotate `TRAFFIC_TOKEN` every 90 days**, and immediately if:

- the workflow file is modified by anyone other than the maintainer;
- the token is used from a workflow other than the traffic export;
- a `permissions:` block is widened;
- the repository gains a collaborator with write access.

### Why the analytics token stays isolated

`Administration: read` is a high-value permission — higher than this workflow's
purpose suggests, but it is the minimum the traffic API accepts. Keeping it in a
dedicated environment means a future workflow that compiles or executes
repository code cannot read it, whatever its own permissions are.

Software-quality automation is deliberately tracked as a separate workflow
([#15](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/15)) for that
reason. Do not extend the traffic workflow to run tests, builds or checks.

### Data retention

GitHub retains traffic data for 14 days. The `traffic-history` branch exists to
keep a longer record. It is public and holds the same aggregate data GitHub
already shows repository admins — view counts, clone counts, referrers, popular
paths. It contains no information identifying individual visitors.

---

## 👤 Maintainer

Maintained by **Daniele Penza**.

