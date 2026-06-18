# 🤝 Contributing to VBA-DATETIMEPICKER

<p align="left">
  <img alt="Contributions" src="https://img.shields.io/badge/Contributions-Welcome-217346">
  <img alt="Language" src="https://img.shields.io/badge/Language-Excel_VBA-blue">
  <img alt="Style" src="https://img.shields.io/badge/Style-House_conventions-6f42c1">
  <img alt="Tests" src="https://img.shields.io/badge/Tests-TST__DP__RunAll-orange">
  <img alt="License" src="https://img.shields.io/badge/Contributions-MIT-green">
</p>

Thanks for your interest in improving the Date / Time Picker. This project values
small, surgical, well-documented changes that match the existing conventions over
large rewrites. This guide explains how to work with the VBA source and what a
change needs to include before it can be merged.

---

## 💬 Before you start

<p align="left">
  <img alt="Step" src="https://img.shields.io/badge/Step-Open_an_issue_first-217346">
</p>

Please **open an issue before starting non-trivial work** so the approach can be
agreed up front. Good issues to raise:

- a clear bug with reproduction steps (Excel version, 32/64-bit, OS)
- a focused enhancement with a concrete use case
- a documentation gap or inaccuracy

Tiny fixes (typos, comment corrections, obvious one-line bugs) can go straight to a
pull request without a prior issue.

---

## 🧰 Project layout and toolchain

<p align="left">
  <img alt="Source" src="https://img.shields.io/badge/Source-Exported_VBA-217346">
  <img alt="Tool" src="https://img.shields.io/badge/Tool-GitHub_Desktop-blue">
</p>

The repository stores **exported VBA source**, not a binary workbook:

```text
src/modules/M_DatePicker.bas         # coordination, state, write-back, settings
src/classes/cDatePickerManager.cls   # Excel Application event manager
src/classes/cDatePickerLabelHook.cls # per-label MouseMove/Click hook
src/forms/UF_DatePicker.frm (+ .frx)  # modeless UserForm UI
src/ribbon/customUI14.xml            # optional RibbonX layout
test/M_cDP_Test.bas                  # regression harness
demo/                                # demo workbook and builder
dist/                                # build output (the .xlam is NOT tracked)
```

You do not need the git command line. The maintainer works through **GitHub
Desktop**, and that is the recommended workflow for contributors too.

---

## 🔁 Edit and export workflow

<p align="left">
  <img alt="Flow" src="https://img.shields.io/badge/Flow-Import_Edit_Export-217346">
</p>

Because the source lives as exported files, the working loop is:

1. Import the relevant `.bas` / `.cls` / `.frm` files into an Excel workbook
   through the VBE (`File → Import File...`). Keep `UF_DatePicker.frm` and
   `UF_DatePicker.frx` together.
2. Make your change in the VBE and **compile** (`Debug → Compile VBAProject`)
   until it is clean.
3. Run the test harness (see below).
4. **Re-export** each changed component (`File → Export File...`) back over the
   matching file in `src/`, preserving the existing folder layout.
5. Commit the changed text files.

Only commit source files that actually changed. Do not commit the host workbook
you used for editing.

---

## 🧱 Coding standards (house style)

<p align="left">
  <img alt="Explicit" src="https://img.shields.io/badge/Option-Explicit_required-217346">
  <img alt="Banners" src="https://img.shields.io/badge/Doc-Banner_per_procedure-blue">
  <img alt="Prefixes" src="https://img.shields.io/badge/Naming-Prefix_namespaced-6f42c1">
</p>

New code must match the existing conventions. Read a few procedures in
`M_DatePicker.bas` before contributing — they are the reference.

**Module hygiene**

- `Option Explicit` at the top of every module, class, and form.
- Do not add `Option Private Module` to `M_DatePicker.bas`; public Excel UI
  callbacks (CommandBars, `Shape.OnAction`, `Application.OnTime`, RibbonX) depend
  on public procedures there.
- Keep `cDatePickerManager` and `cDatePickerLabelHook` private-instanced
  (`VB_Exposed = False`).

**Procedure banners**

Every procedure carries a banner doc-block in this shape:

```vb
'------------------------------------------------------------------------------
'                              PROCEDURE TITLE
'------------------------------------------------------------------------------
' PURPOSE
'   ...
' WHY THIS EXISTS
'   ...
' INPUTS / RETURNS / BEHAVIOR
'   ...
' ERROR POLICY
'   raises descriptive errors  -- or --  best-effort / safe-default
' DEPENDENCIES
'   ...
' NOTES
'   ...
' UPDATED
'   YYYY-MM-DD
'------------------------------------------------------------------------------
```

**Body structure**

- Open with a `DECLARE` section, then sectioned sub-banners (`INITIALIZE`,
  `VALIDATE`, etc.).
- Declare `Const PROC_NAME As String = "..."` for routines that report errors.
- Put a short intent comment **above** each meaningful statement.

**Naming and prefixes**

| Prefix | Scope |
| --- | --- |
| `DP_` | public entry-point API (`DP_Show`, `DP_Close`, `DP_Preload`, `DP_Hide`) |
| `M_` | module-internal helpers grouped by area (`M_Settings_`, `M_WriteBack_`, `M_GridIcon_`, `M_Window_`) |
| `UF_` | UserForm-level routines |
| `cDatePicker*` | classes |
| `Ribbon_` | RibbonX callbacks |
| `TST_DP_` | test harness |

Keep the stable legacy identifiers (`VBA_DATETIMEPICKER` registry app name and
command-bar tag) unchanged — they exist for backward compatibility.

**Error-handling contract**

- Use a labeled handler whose name states intent: `ErrorHandler` for routines
  that raise, and `FailSafe` / `SafeExit` for best-effort paths.
- The banner `ERROR POLICY` section must say whether the routine raises a
  descriptive error or is best-effort / safe-default — and the code must match.
- Reserve `On Error Resume Next` for genuinely best-effort or cleanup sequences,
  not for normal logic.

**Performance conventions**

- Bulk-load and bulk-write ranges via `.Value2`; fall back to per-cell writes
  only when bulk cannot complete (see `M_WriteBack_TryBulkWriteRange`).
- Use 1-based arrays.
- Cache repeated lookups rather than re-reading them per cell or per render.

---

## 🧪 Testing

<p align="left">
  <img alt="Harness" src="https://img.shields.io/badge/harness-M__cDP__Test-217346">
  <img alt="Entry" src="https://img.shields.io/badge/entry-TST__DP__RunAll-blue">
</p>

Import `test/M_cDP_Test.bas` and run the suite before submitting:

```vb
TST_DP_RunAll
```

Results are written to the Immediate Window and to a `TST_DP_RESULTS` worksheet;
a scratch sheet (`TST_DP_SCRATCH`) is created and removed automatically.

- All existing suites must still pass.
- If you change or add behavior, add or extend a suite (for example
  `TST_DP_RunSuite_WriteBack`) and use the existing `Assert*` helpers.
- Include a `UISmoke` run (`TST_DP_RunAll_WithUISmoke`) when touching the form.

---

## 📚 Documentation expectations

<p align="left">
  <img alt="Docs" src="https://img.shields.io/badge/docs-kept_in_sync-6f42c1">
</p>

Documentation is part of the change, not a follow-up. When your change affects a
public-facing surface, update the matching docs in the same pull request:

- **public API** changed → `Public-API` wiki page **and** the README API table
- **install method / UI internals** changed → `Installation-and-Import` /
  `UserForm-UI-Layer` wiki pages
- any user-visible change → a note in `UPDATE_NOTES`

This mirrors the docs-sync step in the **Release-Checklist** wiki page.

---

## 📦 What not to commit

<p align="left">
  <img alt="Excluded" src="https://img.shields.io/badge/excluded-binaries_and_locks-red">
</p>

- The built **`DATETIMEPICKER.xlam`** — it is a binary artifact published only as
  a release asset and is excluded by `.gitignore`.
- Excel owner/lock files (`~$*`).
- The workbook you used to edit the source.
- Personal settings, scratch sheets, or local paths.

---

## 🚀 Submitting changes

<p align="left">
  <img alt="PR" src="https://img.shields.io/badge/PR-small_and_focused-217346">
</p>

1. Fork the repository and create a branch for your change.
2. Keep the pull request **small and focused** — one logical change per PR.
3. In the PR description, state the problem, the approach, and how you tested it
   (which suites you ran and the result).
4. Confirm: project compiles cleanly, `TST_DP_RunAll` passes, banners and naming
   follow the house style, and docs are updated.

The maintainer reviews changes selectively and may adopt, adapt, or decline a
contribution to keep the codebase coherent. Clear, well-scoped PRs are the most
likely to be merged.

---

## 📄 License

By contributing, you agree that your contributions are licensed under the
project's **MIT License**.

---

## 👤 Maintainer

Maintained by **Daniele Penza**. For anything that is not a code change — design
questions, larger proposals, or general feedback — open an issue to start the
conversation.
