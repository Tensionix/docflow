"""Guard the formatting pass: step order, the chain between steps, the reference.

The steps themselves are proven tools with their own tests. What can break here
is the seam: a step that renames its files or skips a document must not cut
the chain, and the reference must never be formatted as a target.
"""
from __future__ import annotations

from pathlib import Path
import shutil
import sys
import tempfile
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from system_core.core.jobs import JobContext
from system_core.core.manifest import Operation
from system_core.core.paths import get_project_paths
from system_core.services import office_service

SCRIPTS = (
    "docx_strict_to_docx.py",
    "docx_accept_changes_simple.py",
    "docx_strip_comments.py",
    "docx_nonprinting_clean.py",
    "docx_anomaly_corrector.py",
    "docx_restyle_by_template.py",
    "docx_apply_standard.py",
    "docx_table_width_optimizer.py",
    "docx_text_hygiene_fix.py",
    "docx_audit_processor.py",
    "docx_finalize_black_clean.py",
)


def _value_after(parts: list[str], *flags: str) -> str:
    for flag in flags:
        if flag in parts:
            return parts[parts.index(flag) + 1]
    raise AssertionError(f"none of {flags} in {parts}")


class FormattingPassTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="format_pass_"))
        self.addCleanup(shutil.rmtree, self.tmp, True)
        for name in ("config", "input", "output", "logs", "report", "workspace"):
            (self.tmp / name).mkdir(parents=True, exist_ok=True)
        (self.tmp / "system_core").mkdir()
        for script in SCRIPTS:
            (self.tmp / "system_core" / script).write_text("", encoding="utf-8")
        (self.tmp / "input" / "doc.docx").write_bytes(b"document")
        (self.tmp / "input" / "sub").mkdir()
        (self.tmp / "input" / "sub" / "second.docx").write_bytes(b"second")
        self.calls: list[tuple[str, list[str]]] = []

    def fake_run_command(self, context, command, **kwargs):
        parts = [str(part) for part in command]
        script = Path(parts[4]).name
        self.calls.append((script, parts))
        source = Path(_value_after(parts, "--input", "--input-dir"))
        target = Path(_value_after(parts, "--outdir"))
        for path in source.rglob("*.docx"):
            if script == "docx_text_hygiene_fix.py" and path.name == "second.docx":
                continue  # a tool with nothing to change may write nothing
            relative = path.relative_to(source)
            if script == "docx_table_width_optimizer.py":
                relative = relative.with_name(f"{relative.stem}{office_service.FORMAT_TABLE_SUFFIX}.docx")
            destination = target / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, destination)
        return {"exit_code": 0}

    def call_of(self, script: str) -> list[str]:
        matches = [parts for name, parts in self.calls if name == script]
        self.assertEqual(len(matches), 1, f"{script} called {len(matches)} times")
        return matches[0]

    def run_pass(self, **parameters) -> dict:
        original = office_service._run_command
        office_service._run_command = self.fake_run_command
        try:
            context = JobContext(
                paths=get_project_paths(self.tmp),
                operation=Operation(
                    id="docx_format_documents",
                    title="format",
                    description="",
                    service="system_core.services.office_service:docx_format_documents",
                    parameters=parameters,
                ),
                log_file=self.tmp / "logs" / "format.log",
                report_dir=self.tmp / "report",
            )
            return office_service.docx_format_documents(context)
        finally:
            office_service._run_command = original

    def test_defaults_clean_then_fit_tables(self) -> None:
        result = self.run_pass()
        self.assertEqual(
            [script for script, _ in self.calls],
            [
                "docx_nonprinting_clean.py",
                "docx_restyle_by_template.py",
                "docx_anomaly_corrector.py",
                "docx_apply_standard.py",
                "docx_table_width_optimizer.py",
            ],
        )
        self.assertIn("--clean-only", self.call_of("docx_restyle_by_template.py"))
        self.assertIn("--fit-overflowing", self.call_of("docx_table_width_optimizer.py"))
        formatted = Path(result["outdir"])
        self.assertTrue((formatted / "doc.docx").is_file())
        self.assertTrue((formatted / "sub" / "second.docx").is_file())
        self.assertEqual(list(formatted.rglob(f"*{office_service.FORMAT_TABLE_SUFFIX}.docx")), [])
        self.assertFalse((self.tmp / "workspace" / "format_pass").exists())
        self.assertTrue(Path(result["report"]).is_file())

    def test_hygiene_wraps_styles_and_the_reference_stays_out(self) -> None:
        (self.tmp / "input" / "ref.docx").write_bytes(b"reference")
        result = self.run_pass(
            before_accept_changes=True,
            before_strip_comments=True,
            use_reference=True,
            reference_docx="ref.docx",
            after_text_hygiene=True,
            after_text_hygiene_dot=True,
            after_remove_strikethrough=True,
            after_audit_fix=True,
            after_finalize_black=True,
        )
        self.assertEqual(
            [script for script, _ in self.calls],
            [
                "docx_accept_changes_simple.py",
                "docx_strip_comments.py",
                "docx_nonprinting_clean.py",
                "docx_anomaly_corrector.py",
                "docx_restyle_by_template.py",
                "docx_apply_standard.py",
                "docx_table_width_optimizer.py",
                "docx_text_hygiene_fix.py",
                "docx_audit_processor.py",
                "docx_finalize_black_clean.py",
            ],
        )
        styles = self.call_of("docx_restyle_by_template.py")
        self.assertIn("--template", styles)
        self.assertNotIn("--clean-only", styles)
        hygiene = self.call_of("docx_text_hygiene_fix.py")
        self.assertIn("--fix-dot", hygiene)
        self.assertIn("--remove-strikethrough", hygiene)
        self.assertIn("--json-out", hygiene)
        audit = self.call_of("docx_audit_processor.py")
        self.assertIn("--fix", audit)
        self.assertNotIn("--annotate", audit)
        self.assertEqual(_value_after(audit, "--norm"), "rules")
        formatted = Path(result["outdir"])
        self.assertFalse((formatted / "ref.docx").exists())
        self.assertTrue((formatted / "sub" / "second.docx").is_file())
        self.assertIn("внутри переноса стилей", Path(result["report"]).read_text(encoding="utf-8"))

    def test_audit_step_takes_the_chosen_norm(self) -> None:
        self.run_pass(after_audit_fix=True, after_audit_norm="gost")
        self.assertEqual(_value_after(self.call_of("docx_audit_processor.py"), "--norm"), "gost")
        with self.assertRaises(RuntimeError):
            self.run_pass(after_audit_fix=True, after_audit_norm="iso")

    def test_text_hygiene_options_need_text_hygiene(self) -> None:
        self.run_pass(after_text_hygiene_dot=True, after_remove_strikethrough=True)
        self.assertNotIn("docx_text_hygiene_fix.py", [script for script, _ in self.calls])

    def test_strict_documents_go_to_word_first_and_leave_if_not_resaved(self) -> None:
        with zipfile.ZipFile(self.tmp / "input" / "strict.docx", "w") as archive:
            archive.writestr("word/document.xml", '<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"/>')
        result = self.run_pass(before_nonprinting=False, before_xml_cleanup=False, before_anomalies=False,
                               apply_standard=False, tables_fit_to_margins=False)
        self.assertEqual([script for script, _ in self.calls], ["docx_strict_to_docx.py"])
        self.assertEqual(result["skipped"], ["strict.docx"])
        formatted = Path(result["outdir"])
        self.assertFalse((formatted / "strict.docx").exists())
        self.assertTrue((formatted / "doc.docx").is_file())
        self.assertIn("strict.docx", Path(result["report"]).read_text(encoding="utf-8"))

    def test_nothing_selected_copies_documents_as_they_are(self) -> None:
        result = self.run_pass(before_nonprinting=False, before_xml_cleanup=False, before_anomalies=False,
                               apply_standard=False, tables_fit_to_margins=False)
        self.assertEqual(self.calls, [])
        self.assertEqual((Path(result["outdir"]) / "doc.docx").read_bytes(), b"document")


if __name__ == "__main__":
    unittest.main()
