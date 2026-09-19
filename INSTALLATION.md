<div align="center">

# 📦 Installation and Upgrade Guide

### Install, integrate, validate, upgrade, and remove the Excel Date / Time Picker

[![Deployment](https://img.shields.io/badge/Deployment-Source--first-0969da?style=flat-square)](#deployment-model)
[![Validation](https://img.shields.io/badge/Validation-Required-d97706?style=flat-square)](#validation)
[![Security](https://img.shields.io/badge/Security-Review_before_enabling-d73a49?style=flat-square)](SECURITY.md)
[![Version](https://img.shields.io/badge/Version-VERSION_file-6f42c1?style=flat-square)](VERSION)
[![License](https://img.shields.io/badge/License-MIT-217346?style=flat-square)](LICENSE)

<br>

**Back up · Import one coherent version · Compile · Validate · Preserve caller state**

</div>

---

This guide covers installation, validation, upgrade, recovery, and removal of
**VBA Date / Time Picker**.

> [!IMPORTANT]
> VBA source can execute with the user's Office permissions. Review the exact
> source or use a trusted tagged release, follow the organization's macro
> security policy, and never enable macros in an untrusted workbook.

---

## 🧭 Support baseline

| Item | Requirement |
|---|---|
| Host | Desktop Microsoft Excel for Windows |
| Office bitness | 32-bit and 64-bit Office |
| Version identity | Root `VERSION` file and the selected tag/commit |
| Source policy | Exported repository source is authoritative |
| Licence | MIT |
| Current deployment status | Source-first component with embedded-source and optional `.xlam` deployment. |

Compatibility claims apply only to environments actually certified for the
selected release. Read [README.md](README.md), [CHANGELOG.md](CHANGELOG.md), and
the release notes before installation.

<a id="deployment-model"></a>

## 🎯 Deployment model

The normal supported model permits one participating current-version provider per Excel process. A settings namespace distinguishes persistent configuration; it does not replace the provider lease.

Choose one supported model and keep its source identity explicit:

| Model | Use when | Trust boundary |
|---|---|---|
| Embedded source | The component must travel with a workbook or add-in | Destination project contains the reviewed source |
| Tagged source | You build or integrate the component yourself | Tag/commit and exported files define identity |
| Published binary | The project explicitly ships a workbook/add-in asset | Hash, tag binding, and package smoke evidence are required |
| Development source | Focused testing or contribution work | Not a supported release unless the project says otherwise |

Do not combine files from different tags, commits, release assets, local exports,
or copied workbooks.

---

## 📂 Production source package

| Order | Repository source | VBE component | Responsibility |
|---:|---|---|---|
| 1 | `src/modules/M_DatePicker.bas` | `M_DatePicker` | Public API, settings, write-back, integrations, provider lease, and WinAPI helpers |
| 2 | `src/classes/cDatePickerManager.cls` | `cDatePickerManager` | Application-level Excel event orchestration |
| 3 | `src/classes/cDatePickerLabelHook.cls` | `cDatePickerLabelHook` | Runtime MSForms label-event routing |
| 4 | `src/forms/UF_DatePicker.frm` with adjacent `.frx` | `UF_DatePicker` | Modeless Date / Time Picker UI and resources |

Additional source. Only the RibbonX metadata is genuinely optional; the rest is
required to compile or to test the project as it currently stands:

| Source | Purpose |
|---|---|
| `src/ribbon/customUI14.xml` | Optional RibbonX package metadata; not a VBE module |
| `test/M_cDP_Test.bas` | Regression harness |
| `demo/M_DP_DEMO.bas` | Demonstration entry points, **and required to compile `M_DatePicker`** |
| `demo/M_DEMO_BUILDER.bas` | Demonstration builder, **required by `M_DP_DEMO` and, independently, by the regression harness** |

> [!IMPORTANT]
> The four `src/` components above do not compile on their own. `M_DatePicker`
> exposes the Ribbon demo command, and that command calls
> `DP_Demo_EnsureDemoSheet`, which lives in `demo/M_DP_DEMO.bas`. Omitting the
> demo source leaves an unresolved reference at **Debug → Compile VBAProject**.
>
> The current compile and test dependencies are:
>
> ~~~text
> M_DatePicker.bas  ->  M_DP_DEMO.bas  ->  M_DEMO_BUILDER.bas
> M_cDP_Test.bas    ->  M_DEMO_BUILDER.bas
> ~~~
>
> So a compilable project needs both demo modules, and the regression harness
> additionally needs `M_DEMO_BUILDER.bas` in its own right: it builds its result
> sheet through `DEMO_Sheet_BuildTemplate` and exercises the fast-mode
> transaction through `DEMO_FastMode_Begin`, `DEMO_FastMode_End`,
> `DEMO_FastMode_Test_ArmFault`, `DEMO_FastMode_ResolveFailure` and
> `tDEMOFastModeState`. The harness makes no reference to `M_DP_DEMO`.
>
> This is recorded as current fact, not as intended architecture. Production
> source depending on demonstration source is a boundary problem: it predates
> `v1.2.0` and is not introduced by any `v1.2.2` change. Decoupling the harness
> from the demo builder is tracked as
> [#35](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/35); the
> production-to-demo dependency is tracked as
> [#86](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/86), under the
> module-boundary work in
> [#24](https://github.com/danielep71/VBA-DATETIMEPICKER/issues/24).

> [!CAUTION]
> A `.frm` and its `.frx` companion are one logical component. Keep them in
> the same directory during import, never import the `.frx` separately, and
> never process it as text.

---

## 🚀 Fresh installation

1. Back up the macro-enabled destination workbook and stop any existing DatePicker runtime.
2. Import the standard module, both classes, and the UserForm in the listed order. Keep `UF_DatePicker.frx` beside the `.frm`; do not import it separately.
   Import `demo/M_DP_DEMO.bas` and `demo/M_DEMO_BUILDER.bas` as well — the
   project does not compile without them.
3. Compile the complete VBA project.
4. Wire `DP_Start` and `DP_Stop` into the workbook lifecycle if the component must start automatically.
5. Set any deployment-specific settings namespace before the runtime loads.
6. Save, close, reopen, and exercise `DP_Show` and `DP_Close` in a clean Excel session.

### VBE import procedure

1. Open the destination workbook or add-in and press `Alt+F11`.
2. Select the intended project in Project Explorer.
3. Use **File → Import File…** for exported modules, classes, and forms.
4. Confirm component names match the repository source.
5. Run **Debug → Compile VBAProject**.
6. Save in a macro-capable format such as `.xlsm`, `.xlsb`, or `.xlam`
   when the project requires executable VBA.
7. Close and reopen the host before the clean-session smoke test.

Do not paste source into arbitrarily named modules when an exported component is
available. VBE attributes, component identity, form resources, and line endings
are part of a reproducible source installation.

---

## 📥 Published add-in installation

Use this path when the project ships a `.xlam` asset and your Excel policy
permits add-in installation. It is the **published binary** model from the
deployment table: identity comes from the asset hash and its tag binding, not
from source you imported.

### Download and verify

1. Open the
   [Releases page](https://github.com/danielep71/VBA-DATETIMEPICKER/releases)
   and choose the release you intend to run.
2. Download the `.xlam` asset. Take the exact filename from the Release page.
3. Verify its SHA-256 against the value published on that release before
   enabling macros:

   ~~~text
   certutil -hashfile "<asset filename>" SHA256
   ~~~

4. Stop if the hash does not match. A mismatch means the file is not the
   published artifact, whatever its name.

### Install and enable

1. Right-click the downloaded file, open **Properties**, and unblock it if
   Windows marked it as coming from the internet.
2. Copy it to your add-ins location, or to any trusted location your policy
   allows.
3. In Excel choose **File → Options → Add-ins**, set **Manage** to
   **Excel Add-ins**, and select **Go…**.
4. Use **Browse…** to select the file, then tick it in the list and confirm.
5. Close and reopen Excel so the add-in loads in a clean session.

### Validate

1. Confirm the DatePicker Ribbon group appears.
2. Select a date-formatted cell and confirm the entry path you expect —
   grid icon, context menu, keyboard shortcut or Ribbon — opens the picker.
3. Write a date back and confirm the target cell receives it.
4. Confirm no second provider is already active: a refusal message naming the
   provider lease means another copy owns the Excel process.

Do not assume anything about what a published `.xlam` contains. Package contents
are release-specific and are determined by that release's build and
certification process — the `v1.2.1` package, for instance, ran the regression
pack inside the packaged add-in as part of its certification. Consumer
validation should use the supported user-facing paths above; certification
evidence for the exact package is recorded in the release itself.

### Remove

1. **File → Options → Add-ins → Manage: Excel Add-ins → Go…** and clear the
   checkbox.
2. Close Excel, then delete the `.xlam` file.
3. Reopen Excel and confirm the Ribbon group and all entry paths are gone.

Clearing the checkbox alone leaves the file registered for re-enabling. Deleting
the file without clearing the checkbox leaves Excel reporting a missing add-in on
every start.

---

<a id="validation"></a>

## ✅ Validation

A successful import is not sufficient evidence that the installation is correct.

- Run `TST_DP_RunAll` when the regression harness is installed.
- For UI changes or package certification, run `TST_DP_RunAll_WithUISmoke`, then show and close the form manually.
- Verify ordinary single-cell write-back, explicit table-column fill, formula preservation, protected/array-formula behavior, and structured results.
- Verify provider-lease refusal when another participating provider owns the Excel process.
- Exercise the enabled keyboard, context-menu, grid-icon, and Ribbon entry paths.

### Minimum installation evidence

~~~text
Source tag or full commit SHA:
VERSION:
Files imported:
Excel version/build:
Office bitness:
Operating system:
Compile:
Consumer smoke:
Regression/certification:
Cleanup:
Skipped or unverified:
~~~

Treat a skipped, incomplete, cleanup-failed, or wrong-environment run as not
certified. Static checks and source review do not replace execution in Excel.

---

## ⬆️ Upgrade

Before upgrading:

1. read the complete version-to-version changelog;
2. back up the host and export any local modifications;
3. stop or clean up active component state;
4. identify every required production component;
5. decide whether stored configuration or generated assets are compatible.

- Stop the old provider before replacing source or an add-in.
- Replace the complete module/class/form package from one version; keep the `.frm` and `.frx` synchronized.
- Review write-back, settings, access-path, provider-lease, and formula-preservation changes in the changelog.
- Do not run an embedded copy and an add-in copy concurrently as independent providers.

After replacement, compile and repeat the full installation validation. Do not
claim an upgrade is non-breaking solely because VBA signatures compile.

### Local modifications

A locally modified copy is a fork. Diff it against the old and new exported
source, reapply changes deliberately, and retest. Do not overwrite it and assume
the local behavior survived.

---

## 🧯 Troubleshooting

| Symptom | Check |
|---|---|
| Compile error or missing procedure | Confirm every required component was imported from one version and optional dependencies are present. |
| Ambiguous name | Remove duplicate/legacy modules; do not paste new source beside old components. |
| Form missing controls or corrupt UI | Re-import the `.frm` with its exact adjacent `.frx`. |
| Behavior differs by workbook | Check caller, active-object, settings namespace, references, locale, and date-system assumptions. |
| 32/64-bit failure | Confirm the tested Office bitness and conditional WinAPI declarations. |
| Excel left altered after failure | Run the documented recovery/cleanup path; do not blindly force global state. |
| Security warning | Verify source origin, signature/hash where provided, trusted location policy, and macro settings. |
| Output differs from reference | Confirm exact version, inputs, parameterization, tolerance, environment, and reference independence. |

If recovery is uncertain, save user data separately, close Excel, reopen a clean
session, and reproduce with a minimal sanitized workbook before changing code.

Report suspected vulnerabilities privately under [SECURITY.md](SECURITY.md).

---

## 🗑️ Removal

1. Call `DP_Stop` from the owning provider and close the form.
2. Remove DatePicker callbacks, Ribbon/custom UI, shortcuts, context-menu entries, timers, grid icons, and lifecycle hooks as applicable.
3. Remove the production components, compile, save, and reopen Excel to confirm no runtime registration remains.

Removing files does not automatically remove workbook formulas, Ribbon XML,
registry settings, trusted-location configuration, cached add-ins, shortcuts,
scheduled callbacks, or other integrations. Remove only state the component
owns and document anything intentionally retained.

---

## 🔐 Security and privacy

- Obtain source and assets from the official repository or a verified release.
- Compare the selected tag, `VERSION`, release notes, and any published hash.
- Review VBA before enabling macros.
- Do not test with client, personal, regulated, or confidential workbooks.
- Inspect example and release workbooks for links, connections, names,
  properties, hidden content, and embedded code.
- Follow organizational macro, add-in, trusted-location, and signing policy.
- Report vulnerabilities through [SECURITY.md](SECURITY.md), not publicly.

---

## 📚 Related documentation

- [README.md](README.md) — capabilities, requirements, and public API
- [CHANGELOG.md](CHANGELOG.md) — version history and compatibility
- [CONTRIBUTING.md](CONTRIBUTING.md) — source and validation standards
- [RELEASING.md](RELEASING.md) — maintainer release and provenance procedure
- [SECURITY.md](SECURITY.md) — private vulnerability reporting
- [LICENSE](LICENSE) — MIT licence terms

---

### Installation principle

> Install one identifiable source version, compile it, exercise its real host
> behavior, and keep evidence of what was—and was not—validated.
