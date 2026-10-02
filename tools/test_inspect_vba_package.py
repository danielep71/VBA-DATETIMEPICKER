"""Fixtures for inspect_vba_package.py: built packages with each startup and comparison case."""
import io
import struct
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path

import inspect_vba_package as inspector

FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF
CLASS_BASE = 'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"\n'
WORKBOOK_BASE = 'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\n'
WORKSHEET_BASE = 'Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\n'
MANIFEST = {"project_name_patterns": ["^DP_", "^M_"],
            "configurations": {"fixture": {"modules": ["src/M_DatePicker.bas", "src/cHook.cls"]}}}
STANDARD = ('Attribute VB_Name = "M_DatePicker"\nOption Explicit\n'
            "Public Sub DP_Start()\nEnd Sub\nPublic Sub DP_Init()\n    DP_Start\nEnd Sub\n")
CLASS_STREAM = ('Attribute VB_Name = "cHook"\n' + CLASS_BASE + "Attribute VB_GlobalNameSpace = False\n"
                "Attribute VB_TemplateDerived = False\nAttribute VB_Customizable = False\nOption Explicit\n"
                "Public Sub Fire()\nEnd Sub\n")
CLASS_EXPORT = ("VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n"
                'Attribute VB_Name = "cHook"\nAttribute VB_GlobalNameSpace = False\nOption Explicit\n'
                "Public Sub Fire()\nEnd Sub\n")


def workbook(body: str, base: str = WORKBOOK_BASE, name: str = "ThisWorkbook") -> str:
    return f'Attribute VB_Name = "{name}"\n' + base + "Option Explicit\n" + body


def compress(data: bytes) -> bytes:
    """A valid [MS-OVBA] container: full chunks stored raw, the remainder as literal-only tokens."""
    out = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        piece = data[start:start + 4096]
        if len(piece) == 4096:
            out += struct.pack("<H", 0x3000 | 4095) + piece
            continue
        assert len(piece) <= 3640, "fixture remainder too long for a literal-only chunk"
        body = bytearray()
        for index in range(0, len(piece), 8):
            body.append(0)
            body += piece[index:index + 8]
        out += struct.pack("<H", 0xB000 | (len(body) - 1)) + body
    return bytes(out)


def record(rid: int, body: bytes) -> bytes:
    return struct.pack("<HI", rid, len(body)) + body


def dir_stream(modules: list[tuple[str, bool]], declared: int | None = None) -> bytes:
    """A decompressed dir stream with the records [MS-OVBA] 2.3.4.2 requires, in order."""
    data = record(0x0001, struct.pack("<I", 3)) + record(0x0002, struct.pack("<I", 0x409))
    data += record(0x0014, struct.pack("<I", 0x409)) + record(0x0003, struct.pack("<H", 1252))
    data += record(0x0004, b"VBAProject") + record(0x0005, b"") + record(0x0040, b"")
    data += record(0x0006, b"") + record(0x003D, b"") + record(0x0007, struct.pack("<I", 0))
    data += record(0x0008, struct.pack("<I", 0)) + struct.pack("<HIIH", 0x0009, 4, 1, 0)
    data += record(0x000C, b"") + record(0x003C, b"")
    data += record(0x000F, struct.pack("<H", len(modules) if declared is None else declared))
    data += record(0x0013, b"\xff\xff")
    for name, procedural in modules:
        data += record(0x0019, name.encode("cp1252")) + record(0x0047, name.encode("utf-16-le"))
        data += record(0x001A, name.encode("cp1252")) + record(0x0032, name.encode("utf-16-le"))
        data += record(0x001C, b"") + record(0x0048, b"") + record(0x0031, struct.pack("<I", 0))
        data += record(0x001E, struct.pack("<I", 0)) + record(0x002C, b"\xff\xff")
        data += record(0x0021 if procedural else 0x0022, b"") + record(0x002B, b"")
    return data + record(0x0010, b"")


def compound(streams: dict[tuple[str, ...], bytes]) -> bytes:
    """Minimal version-3 compound file: Root, PROJECT and the VBA storage with its streams."""
    names = [("Root Entry", 5), ("PROJECT", 2), ("VBA", 1)] + [(p[1], 2) for p in streams if p[0] == "VBA"]
    data = {1: streams[("PROJECT",)]}
    data.update({i: streams[("VBA", n)] for i, (n, _) in enumerate(names) if i >= 3})
    mini, minifat, starts = bytearray(), [], {}
    for index, blob in data.items():
        if len(blob) < 4096:
            count = max(1, -(-len(blob) // 64))
            starts[index] = len(minifat)
            minifat += [len(minifat) + k + 1 for k in range(count - 1)] + [ENDOFCHAIN]
            mini += blob.ljust(count * 64, b"\0")
    blobs = [b"", struct.pack(f"<{len(minifat)}I", *minifat), bytes(mini)]
    large = [i for i, blob in data.items() if len(blob) >= 4096]
    blobs += [data[i] for i in large]
    sizes = [max(1, -(-len(b) // 512)) for b in blobs]
    sizes[0] = -(-len(names) // 4)
    fat_count = 1
    while fat_count * 128 < fat_count + sum(sizes):
        fat_count += 1
    fat, first = [FATSECT] * fat_count, []
    for size in sizes:
        first.append(len(fat))
        fat += [len(fat) + k + 1 for k in range(size - 1)] + [ENDOFCHAIN]
    fat += [FREESECT] * (fat_count * 128 - len(fat))
    entries = bytearray()
    for index, (name, kind) in enumerate(names):
        left = right = child = NOSTREAM
        if index == 0:
            child = 1
        elif index == 1:
            right = 2
        elif index == 2:
            child = 3 if len(names) > 3 else NOSTREAM
        elif index + 1 < len(names):
            right = index + 1
        if index == 0:
            start, size = first[2], len(mini)
        elif index in starts:
            start, size = starts[index], len(data[index])
        elif index in data:
            start, size = first[3 + large.index(index)], len(data[index])
        else:
            start, size = 0, 0
        encoded = (name + "\0").encode("utf-16-le")
        entries += encoded.ljust(64, b"\0") + struct.pack("<HBB3I", len(encoded), kind, 1, left, right, child)
        entries += b"\0" * 36 + struct.pack("<IQ", start, size)
    entries = entries.ljust(sizes[0] * 512, b"\0")
    blobs[0] = bytes(entries)
    header = bytearray(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\0" * 16)
    header += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6
    header += struct.pack("<9I", 0, fat_count, first[0], 0, 4096, first[1] if minifat else ENDOFCHAIN,
                          sizes[1] if minifat else 0, ENDOFCHAIN, 0)
    header += struct.pack("<109I", *(list(range(fat_count)) + [FREESECT] * (109 - fat_count)))
    body = struct.pack(f"<{len(fat)}I", *fat)
    for blob, size in zip(blobs, sizes):
        body += blob.ljust(size * 512, b"\0")
    return bytes(header) + body


def package(modules: dict[str, tuple[str, str]], ribbon: bytes | None = None, declared: int | None = None,
            corrupt: bool = False) -> bytes:
    """modules: name -> (kind, text); kind is standard, class, form or document."""
    project = 'ID="{00000000-0000-0000-0000-000000000000}"\r\n'
    keys = {"standard": "Module", "class": "Class", "form": "BaseClass", "document": "Document"}
    for name, (kind, _text) in modules.items():
        project += f"{keys[kind]}={name}" + ("/&H00000000" if kind == "document" else "") + "\r\n"
    project += 'Name="VBAProject"\r\n\r\n[Host Extender Info]\r\n&H00000001={3832D640}\r\n'
    streams = {("PROJECT",): project.encode("cp1252")}
    directory = compress(dir_stream([(n, k == "standard") for n, (k, _) in modules.items()], declared))
    streams[("VBA", "dir")] = b"\x02" + directory[1:] if corrupt else directory
    streams[("VBA", "_VBA_PROJECT")] = b"\xCC\x61\xFF\xFF\x00\x00\x00"  # version-independent, no p-code cache
    for name, (_kind, text) in modules.items():
        streams[("VBA", name)] = compress(text.replace("\n", "\r\n").encode("cp1252"))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("xl/vbaProject.bin", compound(streams))
        if ribbon is not None:
            archive.writestr("customUI/customUI.xml", ribbon)
    return buffer.getvalue()


class Fixture(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.repo = self.root / "repo"
        (self.repo / "src").mkdir(parents=True)
        (self.repo / "src" / "M_DatePicker.bas").write_bytes(STANDARD.replace("\n", "\r\n").encode("cp1252"))
        (self.repo / "src" / "cHook.cls").write_bytes(CLASS_EXPORT.replace("\n", "\r\n").encode("cp1252"))
        for command in (["init", "-q"], ["add", "."],
                        ["-c", "user.name=fixture", "-c", "user.email=fixture@example.invalid",
                         "commit", "-q", "-m", "fixture"]):
            subprocess.run(["git", "-C", str(self.repo), *command], check=True, capture_output=True)

    def tearDown(self):
        self.temp.cleanup()

    def inspect(self, hook: str | None = None, extra: dict | None = None, raw: bytes | None = None,
                **options) -> dict:
        modules = {"ThisWorkbook": ("document", workbook(hook or "")),
                   "M_DatePicker": ("standard", STANDARD), "cHook": ("class", CLASS_STREAM)}
        modules.update(extra or {})
        modules = {k: v for k, v in modules.items() if v is not None}
        path = self.root / "fixture.xlsm"
        path.write_bytes(raw if raw is not None else package(modules, **options))
        return inspector.inspect(path, None, self.repo, "HEAD", "fixture", MANIFEST)

    @staticmethod
    def codes(report: dict, severity: str | None = None) -> list[str]:
        return sorted(f["code"] for f in report["findings"] + report["inspection"]["failures"]
                      if severity is None or f["severity"] == severity)


class ExtractionTests(Fixture):
    def test_decompression_example_from_the_specification(self):
        compressed = bytes.fromhex("012FB0002361616162636465826600706167686" "96A013808616B6C00306D6E6F70067102"
                                   "700410727374757610777879 7A003C".replace(" ", ""))
        self.assertEqual(inspector.decompress(compressed), b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa")

    def test_inventory_records_identity_kinds_and_raw_hashes(self):
        long_module = STANDARD + "'" + "x" * 5000 + "\n"
        report = self.inspect("Private Sub Workbook_Open()\n    DP_Start\nEnd Sub\n",
                              {"M_DatePicker": ("standard", long_module),
                               "Sheet1": ("document", workbook("", WORKSHEET_BASE, "Sheet1"))})
        self.assertEqual(report["inspection"]["status"], "complete")
        self.assertEqual(report["package"]["filename"], "fixture.xlsm")
        self.assertEqual(len(report["package"]["sha256"]), 64)
        rows = {r["name"]: r for r in report["modules"]}
        self.assertEqual((rows["ThisWorkbook"]["kind"], rows["ThisWorkbook"]["document_object"]), ("document", "workbook"))
        self.assertEqual(rows["Sheet1"]["document_object"], "worksheet")
        self.assertEqual(rows["cHook"]["kind"], "class")
        self.assertEqual(rows["cHook"]["comparison"], "identical")
        self.assertEqual(rows["M_DatePicker"]["comparison"], "comments-only")
        self.assertEqual(len(rows["M_DatePicker"]["raw_sha256"]), 64)
        self.assertEqual(report["source"]["revision"], subprocess.run(
            ["git", "-C", str(self.repo), "rev-parse", "HEAD"], capture_output=True, text=True).stdout.strip())

    def test_extraction_failures_never_give_a_clean_verdict(self):
        cases = {
            "not a zip": b"not a package",
            "no VBA project": self._zip_without_vba(),
            "corrupt dir stream": package({"M_DatePicker": ("standard", STANDARD)}, corrupt=True),
            "module count mismatch": package({"M_DatePicker": ("standard", STANDARD)}, declared=2),
        }
        for label, raw in cases.items():
            with self.subTest(label):
                report = self.inspect(raw=raw)
                self.assertEqual(report["inspection"]["status"], "failed")
                self.assertEqual(report["status"], "incomplete")
                self.assertEqual(report["startup"]["verdict"], "not-established")
                self.assertEqual(report["outcome"]["status"], "not-established")
                self.assertIn("VBA-PKG-090", self.codes(report))
                self.assertNotIn("VBA-PKG-001", self.codes(report))

    @staticmethod
    def _zip_without_vba() -> bytes:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types/>")
        return buffer.getvalue()


class StartupTests(Fixture):
    def test_valid_hook_with_direct_call(self):
        report = self.inspect("Private Sub Workbook_Open()\n    On Error Resume Next\n    DP_Start\nEnd Sub\n")
        self.assertEqual(report["status"], "pass")
        self.assertEqual(report["startup"]["verdict"], "wired")
        hook = report["startup"]["environments"]["vba7-win64"]["hooks"][0]
        self.assertEqual((hook["placement"], hook["signature_check"]), ("correct", "correct"))
        self.assertEqual(hook["reaches_target"][0]["path"], ["ThisWorkbook.Workbook_Open:6 -> DP_Start"])
        self.assertEqual(self.codes(report, "error"), [])

    def test_intermediate_call_and_auto_open(self):
        report = self.inspect(extra={"M_Boot": ("standard", 'Attribute VB_Name = "M_Boot"\n'
                                                "Sub Auto_Open()\n    M_DatePicker.DP_Init\nEnd Sub\n")})
        self.assertEqual(report["startup"]["verdict"], "wired")
        hook = report["startup"]["environments"]["vba7-win64"]["hooks"][0]
        self.assertEqual(hook["hook"], "Auto_Open")
        self.assertEqual(len(hook["reaches_target"][0]["path"]), 2)

    def test_absent_hook_is_reported(self):
        report = self.inspect("Private Sub Workbook_BeforeClose(Cancel As Boolean)\nEnd Sub\n")
        self.assertEqual(report["status"], "fail")
        self.assertEqual(report["startup"]["verdict"], "not-wired")
        self.assertIn("VBA-PKG-001", self.codes(report, "error"))

    def test_misplaced_and_mismatched_hooks(self):
        misplaced = self.inspect(extra={"M_Boot": ("standard", 'Attribute VB_Name = "M_Boot"\n'
                                                   "Public Sub Workbook_Open()\n    DP_Start\nEnd Sub\n"),
                                        "Sheet1": ("document", workbook("Private Sub Auto_Open()\n    DP_Start\nEnd Sub\n",
                                                                         WORKSHEET_BASE, "Sheet1"))})
        self.assertEqual(self.codes(misplaced, "error"), ["VBA-PKG-001", "VBA-PKG-002", "VBA-PKG-002"])
        self.assertEqual(misplaced["startup"]["verdict"], "not-wired")
        signature = self.inspect("Private Sub Workbook_Open(ByVal x As Long)\n    DP_Start\nEnd Sub\n")
        self.assertEqual(self.codes(signature, "error"), ["VBA-PKG-001", "VBA-PKG-003"])

    def test_conditional_path_and_addin_install_only(self):
        conditional = self.inspect("Private Sub Workbook_Open()\n    If Application.EnableEvents Then DP_Start\nEnd Sub\n")
        self.assertEqual(conditional["startup"]["verdict"], "wired-conditionally")
        self.assertIn("VBA-PKG-004", self.codes(conditional, "warning"))
        after_exit = self.inspect("Private Sub Workbook_Open()\n    Exit Sub\n    DP_Start\nEnd Sub\n")
        self.assertEqual(after_exit["startup"]["verdict"], "wired-conditionally")
        install = self.inspect("Private Sub Workbook_AddinInstall()\n    DP_Start\nEnd Sub\n")
        self.assertIn("VBA-PKG-007", self.codes(install, "warning"))

    def test_dynamic_and_unresolved_calls_are_unknown(self):
        dynamic = self.inspect('Private Sub Workbook_Open()\n    Application.Run "DP_Start"\nEnd Sub\n')
        self.assertEqual(dynamic["startup"]["verdict"], "not-established")
        self.assertEqual(self.codes(dynamic, "error"), [])
        self.assertEqual(self.codes(dynamic, "unknown"), ["VBA-PKG-001", "VBA-PKG-005"])
        unresolved = self.inspect("Private Sub Workbook_Open()\n    DP_Missing\nEnd Sub\n")
        self.assertEqual(unresolved["startup"]["verdict"], "not-established")
        self.assertIn("VBA-PKG-005", self.codes(unresolved, "unknown"))

    def test_error_diversion_before_target(self):
        report = self.inspect("Private Sub Workbook_Open()\n    On Error GoTo Fail\n    DP_Prepare\n    DP_Start\n"
                              "    Exit Sub\nFail:\n    Debug.Print Err.Description\nEnd Sub\n",
                              {"M_Prep": ("standard", 'Attribute VB_Name = "M_Prep"\n'
                                          "Public Sub DP_Prepare()\n    Dim x As Long\n    x = 1\nEnd Sub\n")})
        self.assertEqual(report["startup"]["verdict"], "wired")
        diversion = [f for f in report["findings"] if f["code"] == "VBA-PKG-009"]
        self.assertEqual(len(diversion), 1)
        self.assertIsNone(diversion[0]["callee_handler"])

    def test_comment_only_hook_and_conditional_compilation(self):
        commented = self.inspect("' Private Sub Workbook_Open()\n'     DP_Start\n' End Sub\n")
        self.assertIn("VBA-PKG-006", self.codes(commented, "info"))
        self.assertEqual(commented["startup"]["verdict"], "not-wired")
        compiled = self.inspect("Private Sub Workbook_Open()\n#If Win64 Then\n    DP_Start\n#End If\nEnd Sub\n")
        self.assertEqual(compiled["startup"]["verdict"], "configuration-dependent")
        dependent = [f for f in compiled["findings"] if f["code"] == "VBA-PKG-001"]
        self.assertTrue(dependent[0]["configuration_dependent"])
        self.assertEqual(dependent[0]["severity"], "unknown")

    def test_ribbon_onload_is_a_startup_path(self):
        report = self.inspect(extra={"M_Ribbon": ("standard", 'Attribute VB_Name = "M_Ribbon"\n'
                                                  "Public Sub DP_RibbonLoad(ribbon As Object)\n    DP_Start\nEnd Sub\n")},
                              ribbon=b'<customUI onLoad="DP_RibbonLoad"><ribbon/></customUI>')
        self.assertEqual(report["package"]["ribbon"]["onLoad"], "DP_RibbonLoad")
        self.assertEqual(report["startup"]["verdict"], "wired")

    def test_incomplete_analysis_never_reports_a_missing_hook(self):
        report = self.inspect("Private Sub Workbook_BeforeClose(Cancel As Boolean)\n",
                              {"M_Other": ("standard", 'Attribute VB_Name = "M_Other"\n#If FEATURE Then\n#End If\n')})
        self.assertEqual(report["inspection"]["status"], "incomplete")
        self.assertEqual(report["status"], "incomplete")
        self.assertEqual(report["startup"]["verdict"], "not-established")
        self.assertIn("VBA-PKG-091", self.codes(report, "failure"))
        self.assertEqual(self.codes(report, "error"), [])


class ComparisonTests(Fixture):
    def test_changed_missing_unexpected_and_comment_only_modules(self):
        hook = "Private Sub Workbook_Open()\n    DP_Start\nEnd Sub\n"
        changed = self.inspect(hook, {"M_DatePicker": ("standard", STANDARD.replace("    DP_Start\n", "    DP_Init\n"))})
        self.assertIn("VBA-PKG-012", self.codes(changed, "error"))
        missing = self.inspect(hook, {"cHook": None})
        self.assertIn("VBA-PKG-010", self.codes(missing, "error"))
        unexpected = self.inspect(hook, {"M_Extra": ("standard", 'Attribute VB_Name = "M_Extra"\n')})
        self.assertIn("VBA-PKG-011", self.codes(unexpected, "warning"))
        comments = self.inspect(hook, {"M_DatePicker": ("standard", STANDARD + "' trailing note\n")})
        self.assertEqual(self.codes(comments, "error"), [])
        self.assertIn("VBA-PKG-015", self.codes(comments, "info"))
        rows = {r["name"]: r for r in comments["modules"]}
        self.assertNotEqual(rows["M_DatePicker"]["raw_sha256"], rows["M_DatePicker"]["source"]["raw_sha256"])
        self.assertIn("VBA-PKG-013", self.codes(comments, "info"))

    def test_manual_export_is_labelled_and_never_clean(self):
        export = self.root / "export"
        export.mkdir()
        (export / "ThisWorkbook.cls").write_text(
            "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n"
            + workbook("Private Sub Workbook_Open()\n    DP_Start\nEnd Sub\n"), encoding="cp1252")
        (export / "M_DatePicker.bas").write_text(STANDARD, encoding="cp1252")
        (export / "cHook.cls").write_text(CLASS_EXPORT, encoding="cp1252")
        report = inspector.inspect(None, export, self.repo, "HEAD", "fixture", MANIFEST)
        self.assertIsNone(report["package"])
        self.assertEqual(report["inspection"]["evidence"], "manual-export")
        self.assertEqual(report["status"], "incomplete")
        self.assertEqual(report["startup"]["verdict"], "wired")
        self.assertIn("VBA-PKG-092", self.codes(report, "unknown"))

    def test_unresolvable_source_revision_is_incomplete(self):
        path = self.root / "fixture.xlsm"
        path.write_bytes(package({"M_DatePicker": ("standard", STANDARD)}))
        report = inspector.inspect(path, None, self.repo, "no-such-revision", "fixture", MANIFEST)
        self.assertEqual(report["status"], "incomplete")
        self.assertIn("VBA-PKG-091", self.codes(report, "failure"))


if __name__ == "__main__":
    unittest.main()
