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

## 👤 Maintainer

Maintained by **Daniele Penza**.

