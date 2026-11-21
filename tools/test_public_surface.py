"""Offline fixtures for the public-surface classification gate."""
import copy
import unittest
from typing import Any

import check_public_surface as surface

FACADE = "src/modules/M_Fac.bas"
FACADE_SOURCE = """Attribute VB_Name = "M_Fac"
Option Explicit
    Public Enum DP_Mode
        DP_Mode_A = 0 'First
    End Enum
Public Sub DP_Open()
    M_Impl.DP_Open_Impl
End Sub
Public Function DP_Get( _
    ByVal Mode As DP_Mode) As Long
    DP_Get = M_Impl.DP_Get_Impl(Mode)
End Function
Public Sub Ribbon_Open(ByVal control As IRibbonControl)
    M_Impl.Ribbon_Open_Impl control
End Sub
Public Sub DP_Click()
    M_Impl.DP_Click_Impl
End Sub
"""
IMPL = "src/modules/M_Impl.bas"
IMPL_SOURCE = """Attribute VB_Name = "M_Impl"
Option Explicit
Option Private Module
    Public Const DP_LIMIT As Long = 5 'Limit
Public Sub DP_Open_Impl()
End Sub
Public Function DP_Get_Impl(ByVal Mode As DP_Mode) As Long
End Function
Public Sub Ribbon_Open_Impl(ByVal control As IRibbonControl)
End Sub
Public Sub DP_Click_Impl()
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
          module: str = "M_Fac", **extra: Any) -> dict[str, Any]:
    return {"name": name, "module": module, "kind": kind, "classification": classification,
            "channels": channels, "signatures": [signature], **extra}


MANIFEST: dict[str, Any] = {"schema_version": 1, "facade": FACADE, "symbols": [
    entry("DP_LIMIT", "Const", "internal", ["project-internal"], "Const DP_LIMIT As Long = 5", "M_Impl"),
    entry("DP_Mode", "Enum", "supported", ["direct-vba", "project-internal"], "Enum DP_Mode",
          members=["DP_Mode_A = 0"]),
    entry("DP_Open", "Sub", "supported", ["direct-vba", "macro-dialog"], "Sub DP_Open()"),
    entry("DP_Open_Impl", "Sub", "internal", ["project-internal"], "Sub DP_Open_Impl()", "M_Impl"),
    entry("DP_Get", "Function", "supported", ["direct-vba"], "Function DP_Get(ByVal Mode As DP_Mode) As Long"),
    entry("DP_Get_Impl", "Function", "internal", ["project-internal"],
          "Function DP_Get_Impl(ByVal Mode As DP_Mode) As Long", "M_Impl"),
    entry("Ribbon_Open", "Sub", "callback", ["ribbon-onaction"],
          "Sub Ribbon_Open(ByVal control As IRibbonControl)"),
    entry("Ribbon_Open_Impl", "Sub", "internal", ["project-internal"],
          "Sub Ribbon_Open_Impl(ByVal control As IRibbonControl)", "M_Impl"),
    entry("DP_Click", "Sub", "callback", ["onaction", "macro-dialog"], "Sub DP_Click()"),
    entry("DP_Click_Impl", "Sub", "internal", ["project-internal"], "Sub DP_Click_Impl()", "M_Impl"),
    entry("M_Area_Test_Reset", "Sub", "internal", ["test-harness"], "Sub M_Area_Test_Reset()", "M_Impl"),
]}


class PublicSurfaceTests(unittest.TestCase):
    def run_case(self, *, files: dict[str, str] | None = None, manifest: dict[str, Any] | None = None,
                 readme: str = README, ribbon: bytes = RIBBON) -> list[str]:
        project = {FACADE: FACADE_SOURCE, IMPL: IMPL_SOURCE, "src/forms/UF_X.frm": FORM_SOURCE,
                   "test/M_T.bas": TEST_SOURCE}
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
        source = FACADE_SOURCE.replace("Public Sub DP_Open()\n    M_Impl.DP_Open_Impl\nEnd Sub\n", "")
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
                                channels=["direct-vba", "test-harness"])
        readme = README.replace("`DP_Mode`", "`DP_Mode` `M_Area_Test_Reset`")
        self.assertFinding(self.run_case(manifest=manifest, readme=readme), "test seam and must be internal")

    def test_internal_macro_needs_pending_issue(self) -> None:
        manifest = self.changed("DP_Open", classification="internal", channels=["macro-dialog"])
        self.assertFinding(self.run_case(manifest=manifest), "reachable from Alt+F8")
        pending = self.changed("DP_Open", classification="internal", channels=["macro-dialog"], pending="#55")
        self.assertNotIn(True, ["reachable from Alt+F8" in f for f in self.run_case(manifest=pending)])
        stale = self.changed("DP_Open_Impl", pending="#55")
        self.assertFinding(self.run_case(manifest=stale), "'pending' is only for internal symbols")

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
        implementation = 'Attribute VB_Name = "M_Other"\nOption Explicit\nPrivate Sub Work()\nEnd Sub\n'
        findings = self.run_case(files={"src/modules/M_Other.bas": implementation})
        self.assertFinding(findings, "M_Other.bas: implementation modules must use Option Private Module")
        private = implementation.replace("Option Explicit\n", "Option Explicit\nOption Private Module\n")
        self.assertEqual(self.run_case(files={"src/modules/M_Other.bas": private}), [])
        facade = FACADE_SOURCE.replace("Option Explicit\n", "Option Explicit\nOption Private Module\n")
        self.assertFinding(self.run_case(files={FACADE: facade}), "facade must not use Option Private Module")

    def test_private_module_procedure_is_not_a_macro(self) -> None:
        implementation = ('Attribute VB_Name = "M_Other"\nOption Explicit\nOption Private Module\n'
                          "Public Sub M_Other_Run()\nEnd Sub\n")
        manifest = copy.deepcopy(MANIFEST)
        manifest["symbols"].append({"name": "M_Other_Run", "module": "M_Other", "kind": "Sub",
                                    "classification": "internal", "channels": [],
                                    "signatures": ["Sub M_Other_Run()"]})
        self.assertEqual(self.run_case(files={"src/modules/M_Other.bas": implementation}, manifest=manifest), [])

    def test_exposed_object_component_cannot_declare_public_members(self) -> None:
        form = FORM_SOURCE.replace("VB_Exposed = False", "VB_Exposed = True")
        self.assertFinding(self.run_case(files={"src/forms/UF_X.frm": form}), "declares public members: UF_Do")

    def test_string_values_are_not_normalized(self) -> None:
        self.assertEqual(surface.normalize('Const T As String = "Date  /  Time" \'x'),
                         'Const T As String = "Date  /  Time" \'x')
        self.assertEqual(surface.normalize('Sub X( Optional ByVal S As String = "a ,b( c" )'),
                         'Sub X(Optional ByVal S As String = "a ,b( c")')
        source = IMPL_SOURCE + '    Public Const DP_TITLE As String = "A / B" \'Title\n'
        manifest = copy.deepcopy(MANIFEST)
        manifest["symbols"].append(entry("DP_TITLE", "Const", "internal", [], 'Const DP_TITLE As String = "A / B"',
                                         "M_Impl"))
        self.assertEqual(self.run_case(files={IMPL: source}, manifest=manifest), [])
        changed = source.replace('"A / B"', '"A  /  B"')
        self.assertFinding(self.run_case(files={IMPL: changed}, manifest=manifest), "DP_TITLE signature differs")

    def test_conditional_declarations_merge(self) -> None:
        source = FACADE_SOURCE + ("#If VBA7 Then\nPublic Function DP_Hwnd() As LongPtr\nEnd Function\n"
                                  "#Else\nPublic Function DP_Hwnd() As Long\nEnd Function\n#End If\n")
        component = surface.parse_component(FACADE, source)
        hwnd = [s for s in component.symbols if s.name == "DP_Hwnd"]
        self.assertEqual(len(hwnd), 1)
        self.assertEqual(hwnd[0].signatures, ["Function DP_Hwnd() As Long", "Function DP_Hwnd() As LongPtr"])

    def test_conditional_headers_with_shared_body_are_all_compared(self) -> None:
        header = "Public Function DP_Get( _\n    ByVal Mode As DP_Mode) As Long\n"
        alternate = ("#If Win64 Then\n" + header + "#Else\nPublic Function DP_Get(ByVal Mode As Long) As Long\n"
                     "#End If\n")
        source = FACADE_SOURCE.replace(header, alternate)
        self.assertFinding(self.run_case(files={FACADE: source}), "DP_Get signature differs")
        component = surface.parse_component(FACADE, source)
        get = [s for s in component.symbols if s.name == "DP_Get"][0]
        self.assertEqual(get.signatures, ["Function DP_Get(ByVal Mode As DP_Mode) As Long",
                                          "Function DP_Get(ByVal Mode As Long) As Long"])
        body = IMPL_SOURCE.replace("Public Sub DP_Open_Impl()\nEnd Sub\n",
                                   "Public Sub DP_Open_Impl()\n    Dim X As Long\nEnd Sub\n")
        self.assertEqual(self.run_case(files={IMPL: body}), [])

    def test_conditional_enum_and_type_members_are_all_compared(self) -> None:
        enum = "    Public Enum DP_Mode\n        DP_Mode_A = 0 'First\n    End Enum\n"
        alternate = ("#If Win64 Then\n" + enum + "#Else\n" + enum.replace("= 0", "= 99") + "#End If\n")
        self.assertFinding(self.run_case(files={FACADE: FACADE_SOURCE.replace(enum, alternate)}),
                           "DP_Mode signature differs")
        same = "#If Win64 Then\n" + enum + "#Else\n" + enum + "#End If\n"
        self.assertEqual(self.run_case(files={FACADE: FACADE_SOURCE.replace(enum, same)}), [])

    def test_conditional_member_order_is_compared(self) -> None:
        enum = "    Public Enum DP_Mode\n        DP_Mode_A = 0 'First\n    End Enum\n"
        two = "    Public Enum DP_Mode\n        DP_Mode_A\n        DP_Mode_B\n    End Enum\n"
        reordered = two.replace("DP_Mode_A\n        DP_Mode_B", "DP_Mode_B\n        DP_Mode_A")
        manifest = self.changed("DP_Mode", members=["DP_Mode_A", "DP_Mode_B"])
        plain = FACADE_SOURCE.replace(enum, two)
        self.assertEqual(self.run_case(files={FACADE: plain}, manifest=manifest), [])
        source = FACADE_SOURCE.replace(enum, "#If Win64 Then\n" + reordered + "#Else\n" + two + "#End If\n")
        self.assertFinding(self.run_case(files={FACADE: source}, manifest=manifest), "DP_Mode signature differs")
        recorded = self.changed("DP_Mode", members=["DP_Mode_A", "DP_Mode_B"],
                                member_variants=[["DP_Mode_B", "DP_Mode_A"]])
        self.assertEqual(self.run_case(files={FACADE: source}, manifest=recorded), [])

    def test_unmodeled_conditional_symbol_is_reported(self) -> None:
        source = FACADE_SOURCE + "#If DEBUG_BUILD Then\nPublic Sub DP_Debug()\nEnd Sub\n#End If\n"
        findings = self.run_case(files={FACADE: source})
        self.assertTrue(any("M_Fac.bas:" in f and "DEBUG_BUILD" in f for f in findings), findings)

    def test_dispatched_names_resolve_constants_and_literals(self) -> None:
        source = ('Const TICK As String = "M_Tick"\nx = M_GetQualifiedMacroName(TICK)\n'
                  "Application.OnKey \"^+d\", \"'Book.xlsm'!DP_Key\"\n.OnAction = \"DP_Menu\"\n")
        self.assertEqual(surface.dispatched_names(source), {"M_Tick", "DP_Key", "DP_Menu"})

    def test_facade_holds_no_implementation(self) -> None:
        for declaration, expected in (
                ("    Private Declare PtrSafe Function GetTick Lib \"kernel32\" () As Long\n", "Declare PtrSafe"),
                ("    Private mCache As Object 'Cache\n", "Private mCache"),
                ("    Public Const DP_EXTRA As Long = 1 'Extra\n", "Public Const DP_EXTRA"),
                ("    Private Type DP_Hidden\n        A As Long\n    End Type\n", "move DP_Hidden"),
                ("Private Sub Helper()\nEnd Sub\n", "move Helper")):
            source = FACADE_SOURCE.replace("Option Explicit\n", "Option Explicit\n" + declaration)
            self.assertFinding(self.run_case(files={FACADE: source}), expected)

    def test_facade_wrapper_must_delegate_once(self) -> None:
        cases = (
            ("    M_Impl.DP_Open_Impl\n", "    M_Impl.DP_Open_Impl\n    M_Impl.DP_Open_Impl\n"),
            ("    M_Impl.DP_Open_Impl\n", ""),
            ("    M_Impl.DP_Open_Impl\n", "    DP_Open_Impl\n"),
            ("    M_Impl.DP_Open_Impl\n", "    M_Impl.DP_Click_Impl\n"),
            ("    M_Impl.DP_Open_Impl\n", "    M_Fac.DP_Open_Impl\n"),
        )
        for old, new in cases:
            source = FACADE_SOURCE.replace(old, new)
            self.assertFinding(self.run_case(files={FACADE: source}), "DP_Open must delegate in one statement")

    def test_internal_code_never_calls_the_facade(self) -> None:
        source = IMPL_SOURCE.replace("Public Sub DP_Click_Impl()\n", "Public Sub DP_Click_Impl()\n    DP_Open\n")
        self.assertFinding(self.run_case(files={IMPL: source}), "M_Impl.bas: calls the facade wrapper DP_Open")
        form = FORM_SOURCE.replace("Debug.Print DP_LIMIT", "Debug.Print DP_LIMIT\n    M_Fac.DP_Get DP_Mode_A")
        self.assertFinding(self.run_case(files={"src/forms/UF_X.frm": form}), "calls the facade wrapper DP_Get")
        host = ('Attribute VB_Name = "ThisWorkbook"\nOption Explicit\nPrivate Sub Workbook_Open()\n'
                "    DP_Open\nEnd Sub\n")
        manifest = self.changed("DP_Open", channels=["direct-vba", "macro-dialog", "project-internal"])
        self.assertEqual(self.run_case(files={"src/addin/ThisWorkbook.cls": host}, manifest=manifest), [])


if __name__ == "__main__":
    unittest.main()
