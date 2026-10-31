"""Fixtures for check_vba_calls.py: every supported defect next to a valid twin."""
import unittest
from typing import Any

import check_vba_calls as calls

PATTERNS = ["^DP_", "^M_"]


def mod(name: str, body: str, private_module: bool = False) -> str:
    header = f'Attribute VB_Name = "{name}"\nOption Explicit\n'
    if private_module:
        header += "Option Private Module\n"
    return header + body


def cls(name: str, body: str) -> str:
    return ('VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  \'True\nEND\n'
            f'Attribute VB_Name = "{name}"\nAttribute VB_PredeclaredId = False\nOption Explicit\n' + body)


def frm(name: str, body: str) -> str:
    return ('VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ' + name + '\n   Caption = "x"\nEnd\n'
            f'Attribute VB_Name = "{name}"\nAttribute VB_PredeclaredId = True\nOption Explicit\n' + body)


def analyze(files: dict[str, str], gating: bool = True, listed: list[str] | None = None,
            ribbon: bytes | None = None) -> dict[str, Any]:
    config = {"modules": listed if listed is not None else sorted(files), "gating": gating}
    ribbons = None
    if ribbon is not None:
        config["ribbon"] = "ribbon.xml"
        ribbons = {"ribbon.xml": ribbon}
    manifest = {"project_name_patterns": PATTERNS, "macro_name_functions": ["M_Qualify"],
                "configurations": {"fixture": config}}
    return calls.analyze_project(files, manifest, None, ribbons)


def codes(report: dict[str, Any], severity: str = "error") -> list[str]:
    return sorted(f["code"] for f in report["configurations"]["fixture"]["findings"] if f["severity"] == severity)


class ResolutionTests(unittest.TestCase):
    def test_valid_local_and_cross_module_calls(self) -> None:
        report = analyze({
            "a.bas": mod("M_A", "Public Sub DP_Run()\n    DP_Helper 1\n    M_B.DP_Remote\n    Call DP_Remote\nEnd Sub\n"
                                "Private Sub DP_Helper(ByVal x As Long)\nEnd Sub\n"),
            "b.bas": mod("M_B", "Public Sub DP_Remote()\nEnd Sub\n", private_module=True),
        })
        self.assertEqual(report["status"], "pass")
        self.assertEqual(codes(report), [])
        self.assertGreaterEqual(report["configurations"]["fixture"]["stats"]["resolved-calls"], 3)

    def test_missing_unqualified_target_only_for_project_names(self) -> None:
        missing = analyze({"a.bas": mod("M_A", "Public Sub DP_Run()\n    DP_Gone 1\nEnd Sub\n")})
        self.assertEqual(codes(missing), ["VBA-CALL-001"])
        host = analyze({"a.bas": mod("M_A", "Public Sub DP_Run()\n    Beep\n    HostThing 1\n"
                                            "    x = Len(\"a\")\nEnd Sub\n")})
        self.assertEqual(codes(host), [])
        self.assertEqual(host["status"], "pass")

    def test_missing_qualified_member_class_and_form(self) -> None:
        files = {
            "a.bas": mod("M_A", "Public Sub DP_Run()\n    Dim c As cThing\n    Set c = New cThing\n    c.Missing\n"
                                "    c.Present 1\n    M_B.DP_Nope\n    UF_X.lblUnknown.Caption = \"x\"\nEnd Sub\n"),
            "b.bas": mod("M_B", "Public Sub DP_Yes()\nEnd Sub\n"),
            "c.cls": cls("cThing", "Public Sub Present(ByVal x As Long)\nEnd Sub\n"),
            "f.frm": frm("UF_X", "Private Sub UserForm_Initialize()\nEnd Sub\n"),
        }
        report = analyze(files)
        self.assertEqual(codes(report), ["VBA-CALL-002", "VBA-CALL-002"])
        unknown = report["configurations"]["fixture"]["unknown"]
        self.assertIn("member not declared in form code (control or UserForm member)", unknown)

    def test_private_access_and_same_module_private(self) -> None:
        report = analyze({
            "a.bas": mod("M_A", "Public Sub DP_Run()\n    DP_Secret\n    M_B.DP_Secret\nEnd Sub\n"),
            "b.bas": mod("M_B", "Private Sub DP_Secret()\nEnd Sub\nPublic Sub DP_Own()\n    DP_Secret\nEnd Sub\n"),
        })
        self.assertEqual(codes(report), ["VBA-CALL-003", "VBA-CALL-003"])

    def test_option_private_module_stays_project_visible(self) -> None:
        report = analyze({
            "a.bas": mod("M_A", "Public Sub DP_Run()\n    DP_Inner\nEnd Sub\n"),
            "b.bas": mod("M_B", "Public Sub DP_Inner()\nEnd Sub\n", private_module=True),
        })
        self.assertEqual(codes(report), [])

    def test_ambiguity_versus_qualified_and_local_precedence(self) -> None:
        files = {
            "a.bas": mod("M_A", "Public Sub DP_Same()\nEnd Sub\nPublic Sub DP_Local()\n    DP_Same\nEnd Sub\n"),
            "b.bas": mod("M_B", "Public Sub DP_Same()\nEnd Sub\n"),
            "c.bas": mod("M_C", "Public Sub DP_Caller()\n    M_A.DP_Same\n    M_B.DP_Same\n    DP_Same\nEnd Sub\n"),
        }
        report = analyze(files)
        findings = report["configurations"]["fixture"]["findings"]
        ambiguous = [f for f in findings if f["code"] == "VBA-CALL-004"]
        self.assertEqual(len(ambiguous), 1)
        self.assertEqual(ambiguous[0]["module"], "M_C")
        self.assertEqual({c["module"] for c in ambiguous[0]["candidates"]}, {"M_A", "M_B"})
        self.assertIn("VBA-CALL-040", codes(report, "warning"))

    def test_local_shadowing_is_not_a_call(self) -> None:
        report = analyze({
            "a.bas": mod("M_A", "Public Sub DP_Run(ByVal DP_Param As Long)\n    Dim DP_Target(3) As Long\n"
                                "    DP_Target(1) = DP_Param\n    x = DP_Target(2) + DP_Param\nEnd Sub\n"),
            "b.bas": mod("M_B", "Public Sub DP_Target(ByVal a As Long, ByVal b As Long)\nEnd Sub\n"),
        })
        self.assertEqual(codes(report), [])

    def test_comparison_inside_argument_is_not_an_assignment(self) -> None:
        report = analyze({"a.bas": mod("M_A", "Public Sub DP_Assert(ByVal t As String, ByVal c As Boolean)\nEnd Sub\n"
                                              "Public Sub DP_Run()\n    DP_Assert \"x\", 1 = 1\n"
                                              "    DP_Assert \"y\", _\n        2 = 2\nEnd Sub\n")})
        self.assertEqual(codes(report), [])


class ArgumentTests(unittest.TestCase):
    TARGET = ("Public Sub DP_T(ByVal A As Long, Optional ByVal B As Long = 1, Optional C As Variant)\nEnd Sub\n"
              "Public Sub DP_P(ByVal A As Long, ParamArray Rest() As Variant)\nEnd Sub\n"
              "Public Function DP_F(ByVal A As Long) As Long\nEnd Function\n")

    def check(self, body: str) -> list[str]:
        return codes(analyze({"a.bas": mod("M_A", self.TARGET + "Public Sub DP_Run()\n" + body + "End Sub\n")}))

    def test_valid_shapes(self) -> None:
        self.assertEqual(self.check("    DP_T 1\n    DP_T 1, , 3\n    DP_T C:=3, A:=1\n    DP_T 1, C:=3\n"
                                    "    Call DP_T(1, 2)\n    DP_P 1, 2, 3, 4\n    x = DP_F(1) + DP_F(A:=2)\n"
                                    "    DP_T (1)\n    DP_T (1), 2\n"), [])

    def test_too_few(self) -> None:
        self.assertEqual(self.check("    DP_T\n    x = DP_F()\n    x = DP_F\n"), ["VBA-CALL-010"] * 3)

    def test_too_many_without_paramarray(self) -> None:
        self.assertEqual(self.check("    DP_T 1, 2, 3, 4\n    x = DP_F(1, 2)\n"), ["VBA-CALL-011"] * 2)

    def test_omitted_required_versus_omitted_optional(self) -> None:
        self.assertEqual(self.check("    DP_T , 2\n"), ["VBA-CALL-010", "VBA-CALL-012"])
        self.assertEqual(self.check("    DP_T 1, , 3\n"), [])

    def test_named_arguments(self) -> None:
        self.assertEqual(self.check("    DP_T 1, Z:=2\n"), ["VBA-CALL-013"])
        self.assertEqual(self.check("    DP_T 1, A:=2\n"), ["VBA-CALL-014"])
        self.assertEqual(self.check("    DP_T A:=1, A:=2\n"), ["VBA-CALL-014"])
        self.assertEqual(self.check("    DP_P 1, Rest:=2\n"), ["VBA-CALL-015"])
        self.assertEqual(self.check("    DP_T A:=1, 2\n"), ["VBA-CALL-016"])


class PropertyTests(unittest.TestCase):
    CLASS = ("Private m As Long\nPublic Property Get Value(ByVal i As Long) As Long\nEnd Property\n"
             "Public Property Let Value(ByVal i As Long, ByVal v As Long)\nEnd Property\n"
             "Public Property Get ReadOnly() As Long\nEnd Property\n"
             "Public Property Set Target(ByVal o As Object)\nEnd Property\n")

    def check(self, body: str, class_body: str | None = None) -> list[str]:
        return codes(analyze({"c.cls": cls("cBox", class_body or self.CLASS),
                              "a.bas": mod("M_A", "Public Sub DP_Run()\n    Dim b As cBox\n" + body + "End Sub\n")}))

    def test_valid_family(self) -> None:
        self.assertEqual(self.check("    b.Value(1) = 2\n    x = b.Value(1)\n    x = b.ReadOnly\n"
                                    "    Set b.Target = Nothing\n"), [])

    def test_accessor_semantics(self) -> None:
        self.assertEqual(self.check("    b.ReadOnly = 1\n"), ["VBA-CALL-017"])
        self.assertEqual(self.check("    b.Target = 1\n"), ["VBA-CALL-017"])
        self.assertEqual(self.check("    x = b.Value(1, 2)\n"), ["VBA-CALL-011"])
        self.assertEqual(self.check("    b.Value = 2\n"), ["VBA-CALL-010"])

    def test_duplicate_accessor_is_a_conflict(self) -> None:
        duplicate = self.CLASS + "Public Property Get ReadOnly() As Long\nEnd Property\n"
        self.assertEqual(self.check("", duplicate), ["VBA-CALL-020"])
        clash = self.CLASS + "Public Sub ReadOnly()\nEnd Sub\n"
        self.assertEqual(self.check("", clash), ["VBA-CALL-020"])


class DynamicTests(unittest.TestCase):
    def report(self, body: str, extra: str = "", ribbon: bytes | None = None) -> dict[str, Any]:
        return analyze({"a.bas": mod("M_A", "Public Sub DP_Tick()\nEnd Sub\nPrivate Sub DP_Hidden()\nEnd Sub\n"
                                            "Public Function M_Qualify(ByVal n As String) As String\nEnd Function\n"
                                            "Public Sub DP_Run()\n    Dim v As String\n    Dim o As Object\n"
                                            "    Dim c As cHook\n"
                                            + body + "End Sub\n"),
                        "c.cls": cls("cHook", "Public Sub Fire()\nEnd Sub\n" + extra)}, ribbon=ribbon)

    def entries(self, report: dict[str, Any]) -> list[tuple[str, str]]:
        return [(e["channel"], e["status"]) for e in report["configurations"]["fixture"]["entry_points"]]

    def test_literal_targets(self) -> None:
        good = self.report('    Application.Run "DP_Tick"\n    Application.OnTime Now, "M_A.DP_Tick"\n'
                           '    Application.OnKey "^d", M_Qualify("DP_Tick")\n'
                           '    x = Application.Run("\'Book.xlsm\'!DP_Tick")\n'
                           '    CallByName c, "Fire", VbMethod\n')
        self.assertEqual(codes(good), [])
        self.assertIn(("Application.Run", "resolved"), self.entries(good))
        self.assertIn(("CallByName", "resolved"), self.entries(good))
        bad = self.report('    Application.Run "DP_Gone"\n'
                          '    Excel.Application.OnTime EarliestTime:=Now, Procedure:="DP_Gone2"\n'
                          '    CallByName c, "Missing", VbMethod\n    v = M_Qualify("DP_Gone3")\n')
        self.assertEqual(codes(bad), ["VBA-CALL-030"] * 4)
        private = self.report('    Application.Run "DP_Hidden"\n')
        self.assertEqual(codes(private, "warning"), ["VBA-CALL-031"])

    def test_dynamic_targets_are_unknown_not_defects(self) -> None:
        report = self.report('    Application.Run v\n    CallByName o, "Anything", VbMethod\n    .OnAction = v\n')
        self.assertEqual(codes(report), [])
        statuses = [s for _, s in self.entries(report)]
        self.assertGreaterEqual(statuses.count("unknown-dynamic"), 3)

    def test_event_handlers_and_ribbon_callbacks_are_entry_points(self) -> None:
        report = self.report("", "Private WithEvents mApp As Excel.Application\n"
                                 "Private Sub mApp_SheetChange(ByVal S As Object, ByVal T As Object)\nEnd Sub\n"
                                 "Private Sub Class_Initialize()\nEnd Sub\n",
                             ribbon=b'<customUI><button onAction="DP_Tick"/>'
                                    b'<button onAction="DP_NoSuchCallback"/></customUI>')
        entries = self.entries(report)
        self.assertIn(("event-handler (WithEvents)", "external-entry"), entries)
        self.assertIn(("event-handler (class)", "external-entry"), entries)
        self.assertIn(("Ribbon callback", "resolved"), entries)
        self.assertEqual(codes(report), ["VBA-CALL-030"])


    def test_addressof_targets(self) -> None:
        files = {"b.bas": mod("M_B", "Private Function DP_Proc(ByVal h As Long) As Long\nEnd Function\n"
                                     "Public DP_Value As Long\n")}
        good = analyze({**files, "a.bas": mod("M_A", "Private Function DP_Own(ByVal h As Long) As Long\nEnd Function\n"
                                                      "Public Sub DP_Run()\n    x = AddressOf DP_Own\nEnd Sub\n")})
        self.assertEqual(codes(good), [])
        private = analyze({**files, "a.bas": mod("M_A", "Public Sub DP_Run()\n    x = AddressOf DP_Proc\nEnd Sub\n")})
        self.assertEqual(codes(private), ["VBA-CALL-003"])
        value = analyze({**files, "a.bas": mod("M_A", "Public Sub DP_Run()\n    x = AddressOf DP_Value\nEnd Sub\n")})
        self.assertEqual(codes(value), ["VBA-CALL-030"])
        missing = analyze({**files, "a.bas": mod("M_A", "Public Sub DP_Run()\n    x = AddressOf DP_Gone\nEnd Sub\n")})
        self.assertEqual(codes(missing), ["VBA-CALL-030"])

    def test_ribbon_callback_is_checked_in_every_environment(self) -> None:
        report = analyze({"a.bas": mod("M_A", "#If Win64 Then\nPublic Sub DP_Only64()\nEnd Sub\n#End If\n"
                                              "Public Sub DP_Always()\nEnd Sub\n")},
                         ribbon=b'<customUI><button onAction="DP_Only64"/><button onAction="DP_Always"/></customUI>')
        self.assertEqual(report["status"], "pass")
        self.assertEqual(codes(report), [])
        dependent = [f for f in report["configurations"]["fixture"]["findings"] if f["code"] == "VBA-CALL-030"]
        self.assertEqual(len(dependent), 1)
        self.assertEqual(dependent[0]["target"], "DP_Only64")
        self.assertTrue(dependent[0]["configuration_dependent"])
        self.assertEqual(sorted(dependent[0]["environments"]), ["vba6-win32", "vba7-win32"])

    def test_ribbon_callback_in_non_gating_configuration_is_a_warning(self) -> None:
        report = analyze({"a.bas": mod("M_A", "Public Sub DP_Run()\nEnd Sub\n")}, gating=False,
                         ribbon=b'<customUI><button onAction="DP_Gone"/></customUI>')
        self.assertEqual(report["status"], "pass")
        self.assertEqual(codes(report, "warning"), ["VBA-CALL-030"])


class ConfigurationTests(unittest.TestCase):
    def test_alternative_declarations_are_not_duplicates(self) -> None:
        report = analyze({"a.bas": mod("M_A", "#If VBA7 Then\nPublic Sub DP_X(ByVal p As LongPtr)\n#Else\n"
                                              "Public Sub DP_X(ByVal p As Long)\n#End If\nEnd Sub\n"
                                              "Public Sub DP_Run()\n    DP_X 1\nEnd Sub\n")})
        self.assertEqual(report["status"], "pass")
        self.assertEqual(codes(report), [])

    def test_environment_specific_defect_is_configuration_dependent(self) -> None:
        report = analyze({"a.bas": mod("M_A", "#If Win64 Then\nPublic Sub DP_Only64()\nEnd Sub\n#End If\n"
                                              "Public Sub DP_Run()\n    DP_Only64\nEnd Sub\n")})
        findings = report["configurations"]["fixture"]["findings"]
        self.assertEqual(codes(report), [])
        dependent = [f for f in findings if f["code"] == "VBA-CALL-001"]
        self.assertEqual(len(dependent), 1)
        self.assertEqual(dependent[0]["severity"], "unknown")
        self.assertTrue(dependent[0]["configuration_dependent"])
        self.assertEqual(sorted(dependent[0]["environments"]), ["vba6-win32", "vba7-win32"])

    def test_indeterminate_conditional_is_an_analysis_failure(self) -> None:
        report = analyze({"a.bas": mod("M_A",
                                       "#If FEATURE Then\nPublic Sub DP_Run()\n    DP_Gone\nEnd Sub\n#End If\n")})
        self.assertEqual(report["status"], "incomplete")
        self.assertIn("VBA-CALL-090", codes(report, "failure"))
        self.assertEqual(codes(report), [])

    def test_incomplete_coverage_never_passes(self) -> None:
        missing_module = analyze({"a.bas": mod("M_A", "Public Sub DP_Run()\n    DP_Elsewhere\nEnd Sub\n")},
                                 listed=["a.bas", "b.bas"])
        self.assertEqual(missing_module["status"], "incomplete")
        self.assertEqual(codes(missing_module), [])
        self.assertIn("VBA-CALL-001", codes(missing_module, "unknown"))
        unlisted = analyze({"a.bas": mod("M_A", "Public Sub DP_Run()\nEnd Sub\n"),
                            "b.bas": mod("M_B", "Public Sub DP_Other()\nEnd Sub\n")}, listed=["a.bas"])
        self.assertEqual(unlisted["status"], "incomplete")

    def test_parser_failure_never_passes(self) -> None:
        report = analyze({"a.bas": mod("M_A", "Public Sub DP_Run()\n    DP_Gone\n")})
        self.assertEqual(report["status"], "incomplete")
        self.assertIn("VBA-CALL-090", codes(report, "failure"))

    def test_non_gating_configuration_reports_warnings(self) -> None:
        report = analyze({"a.bas": mod("M_A", "Public Sub DP_Run()\n    DP_Gone\nEnd Sub\n")}, gating=False)
        self.assertEqual(report["status"], "pass")
        self.assertEqual(codes(report, "warning"), ["VBA-CALL-001"])


if __name__ == "__main__":
    unittest.main()


class Issue134Tests(unittest.TestCase):
    """False negatives reported by a downstream adopter (#134), each next to the valid case it must keep."""

    def test_project_names_match_case_insensitively(self) -> None:
        for name in ("DP_GONE", "dp_gone", "Dp_Gone"):
            with self.subTest(name=name):
                report = analyze({"a.bas": mod("M_A", f"Public Sub DP_Run()\n    {name} 1\nEnd Sub\n")})
                self.assertEqual(codes(report), ["VBA-CALL-001"])
        host = analyze({"a.bas": mod("M_A", "Public Sub DP_Run()\n    beep\n    hostthing 1\nEnd Sub\n")})
        self.assertEqual(codes(host), [])

    def test_literal_target_in_a_known_module_is_checked_whatever_its_name(self) -> None:
        body = ("Public Sub DP_Run()\n    Application.Run \"M_A.Missing\"\nEnd Sub\n"
                "Public Sub Helper()\nEnd Sub\n")
        self.assertEqual(codes(analyze({"a.bas": mod("M_A", body)})), ["VBA-CALL-030"])
        valid = ("Public Sub DP_Run()\n    Application.Run \"M_A.Helper\"\n"
                 "    Application.Run \"'Other.xlsm'!M_A.Missing\"\n    Application.Run \"Elsewhere.Missing\"\n"
                 "    Application.Run \"cThing.Missing\"\nEnd Sub\nPublic Sub Helper()\nEnd Sub\n")
        report = analyze({"a.bas": mod("M_A", valid), "c.cls": cls("cThing", "Public Sub Present()\nEnd Sub\n")})
        self.assertEqual(codes(report), [])
        statuses = [e["status"] for e in report["configurations"]["fixture"]["entry_points"]
                    if e["channel"] == "Application.Run"]
        self.assertEqual(statuses, ["resolved", "unknown-external", "unknown-external", "unknown-external"])

    def test_recursive_calls_are_argument_checked(self) -> None:
        def check(body: str) -> list[str]:
            return codes(analyze({"a.bas": mod("M_A", body)}))

        self.assertEqual(check("Public Sub DP_Recur(ByVal n As Long)\n    DP_Recur\nEnd Sub\n"), ["VBA-CALL-010"])
        self.assertEqual(check("Public Sub DP_Recur(ByVal n As Long)\n    Call DP_Recur\nEnd Sub\n"),
                         ["VBA-CALL-010"])
        self.assertEqual(check("Public Sub DP_Recur(ByVal n As Long)\n    DP_Recur n - 1\n"
                               "    Call DP_Recur(n - 1)\nEnd Sub\n"), [])
        self.assertEqual(check("Public Function DP_F(ByVal n As Long) As Long\n    DP_F = DP_F()\nEnd Function\n"),
                         ["VBA-CALL-010"])
        self.assertEqual(check("Public Function DP_F(ByVal n As Long) As Long\n    DP_F = n\n"
                               "    If n > 0 Then DP_F = DP_F + DP_F(n - 1)\nEnd Function\n"), [])

    def test_return_value_of_array_function_is_not_a_call(self) -> None:
        body = "Public Function DP_Arr() As Long()\n    Dim x As Long\n    x = DP_Arr(0)\nEnd Function\n"
        self.assertEqual(codes(analyze({"a.bas": mod("M_A", body)})), [])

    def test_class_without_default_instance_is_not_a_receiver(self) -> None:
        predeclared = cls("cPre", "Public Sub Present(ByVal x As Long)\nEnd Sub\n").replace(
            "VB_PredeclaredId = False", "VB_PredeclaredId = True")
        files = {
            "a.bas": mod("M_A", "Public Sub DP_Run()\n    cThing.Present 1\nEnd Sub\n"),
            "c.cls": cls("cThing", "Public Sub Present(ByVal x As Long)\nEnd Sub\n"
                                   "Public Sub Again()\n    Me.Present 1\nEnd Sub\n"),
            "p.cls": predeclared,
        }
        self.assertEqual(codes(analyze(files)), ["VBA-CALL-005"])
        files["a.bas"] = mod("M_A", "Public Sub DP_Run()\n    Dim c As cThing\n    Set c = New cThing\n"
                                    "    c.Present 1\n    cPre.Present 1\n    M_B.DP_Yes\n"
                                    "    CallByName c, \"Present\", VbMethod, 1\nEnd Sub\n")
        files["b.bas"] = mod("M_B", "Public Sub DP_Yes()\nEnd Sub\n")
        self.assertEqual(codes(analyze(files)), [])
        files["a.bas"] = mod("M_A", "Public Sub DP_Run()\n    CallByName cThing, \"Present\", VbMethod, 1\nEnd Sub\n")
        # The class name passed as a value is itself invalid; CallByName's target stays unresolved.
        report = analyze(files)
        self.assertEqual(codes(report), ["VBA-CALL-005"])
        self.assertIn("dynamic CallByName target", report["configurations"]["fixture"]["unknown"])
