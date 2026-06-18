<!--
  Thanks for contributing! Keep PRs small and focused — one logical change each.
  See CONTRIBUTING.md for the house style and workflow.
-->

## 📋 Summary

What does this change do, and why?

## 🔗 Related issue

Closes #<!-- issue number, if any -->

## 🧩 Type of change

- [ ] 🐞 Bug fix
- [ ] ✨ New feature
- [ ] ♻️ Refactor (no behavior change)
- [ ] 📚 Documentation
- [ ] 🧪 Tests / tooling

## 🧪 How it was tested

Which suites did you run, and what was the result?

```text
TST_DP_RunAll            → 
TST_DP_RunAll_WithUISmoke (if the form changed) → 
```

## ✅ Checklist

- [ ] Project compiles cleanly (`Debug → Compile VBAProject`)
- [ ] `TST_DP_RunAll` passes (and `UISmoke` if the form changed)
- [ ] New/changed procedures have a banner doc-block and follow the naming and
      error-handling conventions
- [ ] Changed source files were **re-exported** to `src/` (no editing workbook
      committed)
- [ ] Docs updated where relevant (`Public-API`, README API table,
      `Installation-and-Import`, `UserForm-UI-Layer`, `UPDATE_NOTES`)
- [ ] No binaries or lock files committed (`.xlam`, `~$*`)

## 📎 Notes for the reviewer

Anything that needs special attention, trade-offs, or follow-ups.
