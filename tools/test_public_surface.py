"""Offline fixtures for the public-surface classification gate."""
import copy
import unittest
from typing import Any

import check_public_surface as surface

FACADE = "src/modules/M_Fac.bas"
FACADE_SOURCE = """Attribute VB_Name = "M_Fac"
Option Explicit
    Public Const DP_LIMIT As Long = 5 'Limit
    Public Enum DP_Mode
        DP_Mode_A = 0 'First
    End Enum
Public Sub DP_Open()
End Sub
Public Function DP_Get( _
    ByVal Mode As DP_Mode) As Long
End Function
Public Sub Ribbon_Open(ByVal control As IRibbonControl)
End Sub
Public Sub DP_Click()
End Sub
Public Sub M_Area_Test_Reset()
End Sub
Private Sub Wire()
    Dim MacroName As String
    MacroName = M_GetQualifiedMacroName("DP_Click")
End Sub
Private Function M_GetQualifiedMacroName(ByVal ProcedureName As String) As String
End Function
"""
FORM_SOURCE = """VERSION 5.00
Attribute VB_Name = "UF_X"
Attribute VB_Exposed = False
Option Explicit
Public Sub UF_Do()
    Debug.Print DP_LIMIT
End Sub
"""
TEST_SOURCE = """Attribute VB_Name = "M_T"
Option Explicit
Private Sub Run()
    M_Area_Test_Reset
End Sub
"""
README = """# API
<!-- public-surface:supported -->
| `DP_Open` | `Sub` |
| `DP_Get` | `Function` |
`DP_Mode`
<!-- /public-surface:supported -->
<!-- public-surface:callback -->
`Ribbon_Open` `DP_Click`
<!-- /public-surface:callback -->
"""
RIBBON = b'<customUI><button id="b" onAction="Ribbon_Open"/></customUI>'


def entry(name: str, kind: str, classification: str, channels: list[str], signature: str,
          **extra: Any) -> dict[str, Any]:
    return {"name": name, "module": "M_Fac", "kind": kind, "classification": classification,
            "channels": channels, "signatures": [signature], **extra}


MANIFEST: dict[str, Any] = {"schema_version": 1, "facade": FACADE, "symbols": [
    entry("DP_LIMIT", "Const", "internal", ["project-internal"], "Const DP_LIMIT As Long = 5"),
    entry("DP_Mode", "Enum", "supported", ["direct-vba"], "Enum DP_Mode", members=["DP_Mode_A = 0"]),
    entry("DP_Open", "Sub", "supported", ["direct-vba", "macro-dialog"], "Sub DP_Open()"),
    entry("DP_Get", "Function", "supported", ["direct-vba"], "Function DP_Get(ByVal Mode As DP_Mode) As Long"),
    entry("Ribbon_Open", "Sub", "callback", ["ribbon-onaction"],
          "Sub Ribbon_Open(ByVal control As IRibbonControl)"),
    entry("DP_Click", "Sub", "callback", ["onaction", "macro-dialog"], "Sub DP_Click()"),
    entry("M_Area_Test_Reset", "Sub", "internal", ["macro-dialog", "test-harness"], "Sub M_Area_Test_Reset()",
          pending="#55"),
]}


class PublicSurfaceTests(unittest.TestCase):
    def run_case(self, *, files: dict[str, str] | None = None, manifest: dict[str, Any] | None = None,
                 readme: str = README, ribbon: bytes = RIBBON) -> list[str]:
        project = {FACADE: FACADE_SOURCE, "src/forms/UF_X.frm": FORM_SOURCE, "test/M_T.bas": TEST_SOURCE}
        project.update(files or {})
        report = surface.analyze(project, manifest or MANIFEST, ribbon, readme)
        findings: list[str] = report["findings"]
        return findings

    def assertFinding(self, findings: list[str], text: str) -> None:
        self.assertTrue(any(text in f for f in findings), f"{text!r} not in {findings}")

    def changed(self, name: str, **fields: Any) -> dict[str, Any]:
        manifest = copy.deepcopy(MANIFEST)
        for item in manifest["symbols"]:
            if item["name"] == name:
                item.update(fields)
        return manifest

    def test_classified_fixture_passes(self) -> None:
        self.assertEqual(self.run_case(), [])

    def test_unclassified_addition_fails(self) -> None:
        source = FACADE_SOURCE + "Public Function DP_New() As Long\nEnd Function\n"
        self.assertFinding(self.run_case(files={FACADE: source}), "DP_New is not classified")

    def test_implicit_visibility_fails(self) -> None:
        source = FACADE_SOURCE + "Sub DP_Implicit()\nEnd Sub\n"
        self.assertFinding(self.run_case(files={FACADE: source}), "implicit visibility")

    def test_supported_signature_change_fails(self) -> None:
        source = FACADE_SOURCE.replace("Public Sub DP_Open()", "Public Sub DP_Open(ByVal Force As Boolean)")
        findings = self.run_case(files={FACADE: source})
        self.assertFinding(findings, "incompatible supported change")
        self.assertFinding(findings, "DP_Open: channel macro-dialog is recorded but absent")

    def test_supported_enum_member_change_fails(self) -> None:
        source = FACADE_SOURCE.replace("DP_Mode_A = 0", "DP_Mode_A = 1")
        self.assertFinding(self.run_case(files={FACADE: source}), "DP_Mode signature differs")

    def test_removed_supported_symbol_fails(self) -> None:
        source = FACADE_SOURCE.replace("Public Sub DP_Open()\nEnd Sub\n", "")
        self.assertFinding(self.run_case(files={FACADE: source}), "no longer declared (incompatible")

    def test_orphaned_ribbon_callback_fails(self) -> None:
        findings = self.run_case(ribbon=b"<customUI/>")
        self.assertFinding(findings, "orphaned callback")
        self.assertFinding(findings, "Ribbon_Open: channel ribbon-onaction is recorded but absent")

    def test_unbound_ribbon_name_fails(self) -> None:
        ribbon = b'<customUI><button onAction="Ribbon_Open"/><button onAction="Ribbon_Gone"/></customUI>'
        self.assertFinding(self.run_case(ribbon=ribbon), "Ribbon callback has no public procedure: ribbon_gone")

    def test_by_name_dispatch_must_be_declared(self) -> None:
        manifest = self.changed("DP_Click", classification="internal", channels=["macro-dialog"],
                                pending="#55")
        readme = README.replace("`DP_Click`", "")
        self.assertFinding(self.run_case(manifest=manifest, readme=readme), "dispatched by name but declares no")

    def test_test_seam_must_be_internal(self) -> None:
        manifest = self.changed("M_Area_Test_Reset", classification="supported",
                                channels=["direct-vba", "macro-dialog", "test-harness"])
        manifest["symbols"][-1].pop("pending")
        readme = README.replace("`DP_Mode`", "`DP_Mode` `M_Area_Test_Reset`")
        self.assertFinding(self.run_case(manifest=manifest, readme=readme), "test seam and must be internal")

    def test_internal_macro_needs_pending_issue(self) -> None:
        manifest = copy.deepcopy(MANIFEST)
        manifest["symbols"][-1].pop("pending")
        self.assertFinding(self.run_case(manifest=manifest), "reachable from Alt+F8")

    def test_computed_channel_must_match(self) -> None:
        manifest = self.changed("DP_Open", channels=["direct-vba", "macro-dialog", "project-internal"])
        self.assertFinding(self.run_case(manifest=manifest), "project-internal is recorded but absent")
        manifest = self.changed("DP_LIMIT", channels=[])
        self.assertFinding(self.run_case(manifest=manifest), "project-internal is present in source but missing")

    def test_supported_signature_needs_supported_types(self) -> None:
        manifest = self.changed("DP_Mode", classification="internal", channels=[])
        readme = README.replace("`DP_Mode`", "")
        self.assertFinding(self.run_case(manifest=manifest, readme=readme), "uses DP_Mode, which must also be")

    def test_readme_lists_must_match(self) -> None:
        self.assertFinding(self.run_case(readme=README.replace("`DP_Get`", "")),
                           "supported block differs from the manifest: missing ['DP_Get']")
        self.assertFinding(self.run_case(readme="# API\n"), "no public-surface:callback block")

    def test_implementation_modules_need_option_private_module(self) -> None:
        implementation = 'Attribute VB_Name = "M_Impl"\nOption Explicit\nPrivate Sub Work()\nEnd Sub\n'
        findings = self.run_case(files={"src/modules/M_Impl.bas": implementation})
        self.assertFinding(findings, "M_Impl.bas: implementation modules must use Option Private Module")
        private = implementation.replace("Option Explicit\n", "Option Explicit\nOption Private Module\n")
        self.assertEqual(self.run_case(files={"src/modules/M_Impl.bas": private}), [])
        facade = FACADE_SOURCE.replace("Option Explicit\n", "Option Explicit\nOption Private Module\n")
        self.assertFinding(self.run_case(files={FACADE: facade}), "facade must not use Option Private Module")

    def test_private_module_procedure_is_not_a_macro(self) -> None:
        implementation = ('Attribute VB_Name = "M_Impl"\nOption Explicit\nOption Private Module\n'
                          "Public Sub M_Impl_Run()\nEnd Sub\n")
        manifest = copy.deepcopy(MANIFEST)
        manifest["symbols"].append({"name": "M_Impl_Run", "module": "M_Impl", "kind": "Sub",
                                    "classification": "internal", "channels": [],
                                    "signatures": ["Sub M_Impl_Run()"]})
        self.assertEqual(self.run_case(files={"src/modules/M_Impl.bas": implementation}, manifest=manifest), [])

    def test_exposed_object_component_cannot_declare_public_members(self) -> None:
        form = FORM_SOURCE.replace("VB_Exposed = False", "VB_Exposed = True")
        self.assertFinding(self.run_case(files={"src/forms/UF_X.frm": form}), "declares public members: UF_Do")

    def test_conditional_declarations_merge(self) -> None:
        source = FACADE_SOURCE + ("#If VBA7 Then\nPublic Function DP_Hwnd() As LongPtr\nEnd Function\n"
                                  "#Else\nPublic Function DP_Hwnd() As Long\nEnd Function\n#End If\n")
        component = surface.parse_component(FACADE, source)
        hwnd = [s for s in component.symbols if s.name == "DP_Hwnd"]
        self.assertEqual(len(hwnd), 1)
        self.assertEqual(hwnd[0].signatures, ["Function DP_Hwnd() As Long", "Function DP_Hwnd() As LongPtr"])

    def test_dispatched_names_resolve_constants_and_literals(self) -> None:
        source = ('Const TICK As String = "M_Tick"\nx = M_GetQualifiedMacroName(TICK)\n'
                  "Application.OnKey \"^+d\", \"'Book.xlsm'!DP_Key\"\n.OnAction = \"DP_Menu\"\n")
        self.assertEqual(surface.dispatched_names(source), {"M_Tick", "DP_Key", "DP_Menu"})


if __name__ == "__main__":
    unittest.main()
