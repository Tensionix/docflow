"""Cover the style transfer engine taken from the docx-restyle-by-template skill.

The engine itself is the skill's, proven on real volumes; what needs guarding
here is the seam this project added: folder in, report out, and the two
switches that decide whether a reference and a heading map take part.
"""
from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

PYTHON = ROOT / "runtime" / "python.exe"
SCRIPT = ROOT / "system_core" / "docx_restyle_by_template.py"
PROBE = ROOT / "system_core" / "docx_style_probe.py"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


RSID_ATTRS = ("rsidR", "rsidRDefault", "rsidP", "rsidRPr", "rsidTr", "rsidDel")


def build_docx(path: Path, paragraphs: list[str]) -> None:
    """A minimal document carrying revision attributes and an unused style."""
    body = []
    for index, text in enumerate(paragraphs):
        noise = ' '.join('w:%s="00%02X0000"' % (name, index + n)
                         for n, name in enumerate(RSID_ATTRS))
        body.append(
            '<w:p %s><w:pPr><w:pStyle w:val="Normal"/></w:pPr>'
            '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (noise, text)
        )
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="%s"><w:body>%s'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>'
        % (W, "".join(body))
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="%s">'
        '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="NeverUsed"><w:name w:val="Never Used"/></w:style>'
        '</w:styles>' % W
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/styles.xml", styles)


def document_text(path: Path) -> str:
    from lxml import etree

    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    return "".join(node.text or "" for node in root.iter("{%s}t" % W))


def run_script(script: Path, *args: str) -> subprocess.CompletedProcess[str]:
    if not PYTHON.is_file():
        raise unittest.SkipTest("project runtime is unavailable")
    command = [str(PYTHON), str(script), *args]
    return subprocess.run(command, capture_output=True, text=True, encoding="utf-8")


class CleanupTests(unittest.TestCase):
    def setUp(self) -> None:
        import tempfile

        self.tmp = Path(tempfile.mkdtemp(prefix="restyle_test_"))
        self.addCleanup(lambda: __import__("shutil").rmtree(self.tmp, ignore_errors=True))
        self.source = self.tmp / "in" / "doc.docx"
        build_docx(self.source, ["Общие положения", "Текст раздела.", "Ещё абзац."])

    def test_cleanup_keeps_the_text_and_drops_the_noise(self) -> None:
        out = self.tmp / "out"
        report = self.tmp / "report.md"
        result = run_script(
            SCRIPT, "--input", str(self.source.parent), "--outdir", str(out),
            "--report", str(report), "--clean-only",
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

        produced = out / "doc.docx"
        self.assertTrue(produced.is_file())
        self.assertEqual(document_text(produced), document_text(self.source))

        with zipfile.ZipFile(produced) as z:
            body = z.read("word/document.xml").decode("utf-8")
            styles = z.read("word/styles.xml").decode("utf-8")
        self.assertNotIn("w:rsidR", body)
        self.assertNotIn("NeverUsed", styles)

        text = report.read_text(encoding="utf-8")
        self.assertIn("ЧИСТКА", text)
        self.assertIn("атрибутов ревизий удалено", text)

    def test_without_a_map_nothing_is_re_marked(self) -> None:
        out = self.tmp / "out"
        report = self.tmp / "report.md"
        result = run_script(
            SCRIPT, "--input", str(self.source.parent), "--outdir", str(out), "--report", str(report),
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        text = report.read_text(encoding="utf-8")
        self.assertIn("разметка не выполнялась", text)
        self.assertNotIn("РАЗМЕТКА", text)

    def test_reference_is_never_processed_as_a_target(self) -> None:
        reference = self.tmp / "in" / "Эталон.docx"
        build_docx(reference, ["Эталон"])
        out = self.tmp / "out"
        result = run_script(
            SCRIPT, "--input", str(self.source.parent), "--outdir", str(out),
            "--report", str(self.tmp / "report.md"), "--template", str(reference),
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertTrue((out / "doc.docx").is_file())
        self.assertFalse((out / "Эталон.docx").exists())

    def test_a_missing_reference_stops_before_writing(self) -> None:
        out = self.tmp / "out"
        result = run_script(
            SCRIPT, "--input", str(self.source.parent), "--outdir", str(out),
            "--report", str(self.tmp / "report.md"), "--template", str(self.tmp / "nope.docx"),
        )
        self.assertEqual(result.returncode, 2)
        self.assertFalse(out.exists())

    def test_probe_dumps_the_document_for_the_heading_map(self) -> None:
        out = self.tmp / "probe"
        report = self.tmp / "probe.md"
        result = run_script(
            PROBE, "--input", str(self.source.parent), "--outdir", str(out), "--report", str(report),
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        flat = (out / "doc" / "flat.txt").read_text(encoding="utf-8")
        self.assertIn("Общие положения", flat)
        self.assertIn("ПРИЁМНИК", (out / "doc" / "summary.txt").read_text(encoding="utf-8"))


class ServiceWiringTests(unittest.TestCase):
    """The GUI switches must reach the script as the flags it understands."""

    def setUp(self) -> None:
        import tempfile

        from system_core.core.jobs import JobContext
        from system_core.core.manifest import Operation
        from system_core.core.paths import get_project_paths
        from system_core.services import office_service

        self.office_service = office_service
        self.tmp = Path(tempfile.mkdtemp(prefix="restyle_service_"))
        self.addCleanup(lambda: __import__("shutil").rmtree(self.tmp, ignore_errors=True))
        for name in ("config", "input", "output", "logs", "report"):
            (self.tmp / name).mkdir(parents=True, exist_ok=True)
        (self.tmp / "system_core").mkdir(parents=True, exist_ok=True)
        (self.tmp / "system_core" / "docx_restyle_by_template.py").write_text("", encoding="utf-8")
        (self.tmp / "system_core" / "docx_style_probe.py").write_text("", encoding="utf-8")

        self.JobContext = JobContext
        self.Operation = Operation
        self.get_project_paths = get_project_paths

    def call(self, service_name: str, **parameters):
        recorded: dict[str, list[str]] = {}

        def fake_run_command(context, command, **kwargs):
            recorded["command"] = [str(part) for part in command]
            return {"exit_code": 0}

        original = self.office_service._run_command
        self.office_service._run_command = fake_run_command
        try:
            operation = self.Operation(
                id="probe",
                title="probe",
                description="",
                service=f"system_core.services.office_service:{service_name}",
                parameters=parameters,
            )
            context = self.JobContext(
                paths=self.get_project_paths(self.tmp),
                operation=operation,
                log_file=self.tmp / "logs" / "probe.log",
                report_dir=self.tmp / "report",
            )
            getattr(self.office_service, service_name)(context)
        finally:
            self.office_service._run_command = original
        return recorded["command"]

    def test_cleanup_always_asks_for_cleanup_only(self) -> None:
        command = self.call("docx_xml_cleanup")
        self.assertIn("--clean-only", command)
        self.assertNotIn("--template", command)
        self.assertNotIn("--config", command)

    def test_reference_travels_only_when_the_box_is_ticked(self) -> None:
        reference = self.tmp / "input" / "ref.docx"
        reference.write_bytes(b"PK")
        without = self.call("docx_style_fix", use_reference=False, reference_docx="ref.docx")
        self.assertNotIn("--template", without)

        with_reference = self.call("docx_style_fix", use_reference=True, reference_docx="ref.docx")
        self.assertIn("--template", with_reference)
        self.assertIn(str(reference), with_reference)

    def test_heading_map_is_passed_and_checked(self) -> None:
        config = self.tmp / "input" / "map.json"
        config.write_text(json.dumps({"styles": {}}), encoding="utf-8")
        command = self.call("docx_style_fix", style_config="map.json")
        self.assertIn("--config", command)
        self.assertIn(str(config), command)

        with self.assertRaises(FileNotFoundError):
            self.call("docx_style_fix", style_config="missing.json")


if __name__ == "__main__":
    unittest.main()
