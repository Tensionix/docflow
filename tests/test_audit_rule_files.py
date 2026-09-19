"""Audit rules as data: the rule files, their examples, the norm switch.

Every rule in config/rules carries fix and keep examples; loading a file runs
them. These tests hold the loader to that promise and check the processor end
to end: the chosen files and the chosen norm reach the document.
"""
from __future__ import annotations

from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest import mock
import zipfile

from docx import Document
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
SYSTEM_CORE = ROOT / "system_core"
if str(SYSTEM_CORE) not in sys.path:
    sys.path.insert(0, str(SYSTEM_CORE))

import audit_rule_files as rules  # noqa: E402

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from system_core.core.jobs import JobContext  # noqa: E402
from system_core.core.manifest import Operation  # noqa: E402
from system_core.core.paths import get_project_paths  # noqa: E402
from system_core.services import office_service  # noqa: E402

PYTHON = ROOT / "runtime" / "python.exe"
PROCESSOR = SYSTEM_CORE / "docx_audit_processor.py"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NBSP = chr(0xA0)
CYRILLIC_ES = chr(0x421)


def fixed(text: str, norm: str, sets: tuple[str, ...] = rules.RULE_SETS) -> str:
    book = rules.load_rule_book(sets, norm)
    return rules.apply_rules(book.rules, text, fix=True)[0]


class RuleFilesTests(unittest.TestCase):
    def test_every_rule_file_passes_its_examples_in_both_norms(self) -> None:
        for norm in rules.NORMS:
            book = rules.load_rule_book(rules.RULE_SETS, norm)
            self.assertEqual(book.warnings, [], norm)
            self.assertEqual([path.stem for path in book.files], list(rules.RULE_SETS))
            self.assertTrue(all(rule.fix_examples or rule.keep_examples for rule in book.rules))

    def test_every_rule_names_its_source(self) -> None:
        for norm in rules.NORMS:
            book = rules.load_rule_book(rules.RULE_SETS, norm)
            self.assertEqual([rule.code for rule in book.rules if not rule.sources], [], norm)

    def test_the_proven_rules_keep_their_codes_files_and_modes(self) -> None:
        book = rules.load_rule_book(rules.RULE_SETS, "rules")
        by_code = {rule.code: rule for rule in book.rules}
        expected = {
            "AUDIT-UNIT-SQM": ("units", "FIX"),
            "AUDIT-UNIT-CUBIC": ("units", "FIX"),
            "AUDIT-UNIT-SPACE": ("units", "FIX"),
            "AUDIT-PERCENT": ("units", "FIX"),
            "AUDIT-DEGREE": ("units", "FIX"),
            "AUDIT-NUMBER-SIGN": ("house", "FIX"),
            "AUDIT-DATE-YEAR-SUFFIX": ("house", "FIX"),
            "AUDIT-CAPTION-TABLE-TYPO": ("house", "FIX"),
            "AUDIT-RF-SCAN": ("acronyms", "SUGGEST"),
        }
        for code, (rule_set, mode) in expected.items():
            self.assertEqual((by_code[code].rule_set, by_code[code].mode), (rule_set, mode), code)
        candidates = [rule for rule in book.rules if rule.status == "candidate"]
        self.assertTrue(candidates)
        self.assertTrue(all(rule.mode == "SUGGEST" for rule in candidates))

    def test_norm_switch_writes_percent_and_degrees(self) -> None:
        self.assertEqual(fixed(f"износ 52{NBSP}%, график 73/64{NBSP}°{CYRILLIC_ES}", "rules"), "износ 52%, график 73/64°C")
        gost = fixed("износ 52%, график 73/64°C, угол 20 °", "gost")
        self.assertEqual(gost, f"износ 52{NBSP}%, график 73/64{NBSP}°C, угол 20°")
        self.assertEqual(fixed(f"52{NBSP}% и 64{NBSP}°{CYRILLIC_ES}", "gost"), f"52{NBSP}% и 64{NBSP}°{CYRILLIC_ES}")

    def test_other_rules_keep_non_breaking_spaces(self) -> None:
        text = f"№{NBSP}80, 0,5{NBSP}куб.{NBSP}м, 10{NBSP}кв.м, Таблица.{NBSP}3"
        expected = f"№{NBSP}80, 0,5{NBSP}куб.{NBSP}м, 10{NBSP}кв. м, Таблица{NBSP}3"
        for norm in rules.NORMS:
            self.assertEqual(fixed(text, norm), expected, norm)

    def test_candidates_only_report(self) -> None:
        book = rules.load_rule_book(rules.RULE_SETS, "rules")
        text, hits = rules.apply_rules(book.rules, "т.е. 5 млн. руб. на ул.Ленина", fix=True)
        self.assertEqual(text, "т.е. 5 млн. руб. на ул.Ленина")
        self.assertEqual(
            {(hit.rule.code, hit.action) for hit in hits},
            {("AUDIT-ABBR-COMPOUND", "SUGGEST"), ("AUDIT-ABBR-MILLION", "SUGGEST"), ("AUDIT-ADDR-TYPE-SPACE", "SUGGEST")},
        )

    def test_carry_spaces_never_turns_a_non_breaking_space_ordinary(self) -> None:
        self.assertEqual(rules.carry_spaces(f"№{NBSP}1", "№ 1"), f"№{NBSP}1")
        self.assertEqual(rules.carry_spaces("№1", "№ 1"), "№ 1")
        self.assertEqual(rules.carry_spaces(f"10{NBSP}м2", "10 кв. м"), f"10{NBSP}кв. м")
        self.assertEqual(rules.carry_spaces(f"1{NBSP}000 %", f"1{NBSP}000%"), f"1{NBSP}000%")


class BrokenRuleFilesTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="audit_rules_"))
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def write(self, name: str, body: str) -> None:
        (self.tmp / f"{name}.yaml").write_text(body, encoding="utf-8")

    def test_a_rule_that_fails_its_example_only_reports(self) -> None:
        self.write("house", """
rules:
  - code: TEST-WRONG
    action: fix
    match: 'abc'
    replace: 'abd'
    fix:
      "abc": "xyz"
  - code: TEST-UNTESTED
    action: fix
    match: 'klm'
    replace: 'kln'
""")
        book = rules.load_rule_book(["house"], "rules", rules_dir=self.tmp)
        self.assertEqual([rule.mode for rule in book.rules], ["SUGGEST", "SUGGEST"])
        self.assertEqual(len(book.warnings), 2)
        text, hits = rules.apply_rules(book.rules, "abc klm", fix=True)
        self.assertEqual(text, "abc klm")
        self.assertEqual([hit.action for hit in hits], ["SUGGEST", "SUGGEST"])

    def test_a_broken_rule_is_skipped_and_the_rest_still_run(self) -> None:
        self.write("house", """
rules:
  - code: TEST-BAD-REGEX
    match: '(unclosed'
    replace: 'x'
  - code: TEST-BAD-GROUP
    match: 'abc'
    replace: '{missing}'
  - code: TEST-GOOD
    match: '№{SP}*(?P<num>\\d+)'
    replace: '№ {num}'
    fix:
      "№1": "№ 1"
""")
        self.write("units", "rules: [unclosed")
        book = rules.load_rule_book(["units", "house"], "rules", rules_dir=self.tmp)
        self.assertEqual([rule.code for rule in book.rules], ["TEST-GOOD"])
        self.assertEqual(book.rules[0].mode, "FIX")
        joined = "\n".join(book.warnings)
        for part in ("units.yaml", "TEST-BAD-REGEX", "TEST-BAD-GROUP"):
            self.assertIn(part, joined)

    def test_unknown_rule_file_is_refused(self) -> None:
        with self.assertRaises(rules.RuleError):
            rules.parse_rule_sets("units,spelling")
        self.assertEqual(rules.parse_rule_sets("acronyms, units"), ("units", "acronyms"))
        self.assertEqual(rules.parse_rule_sets(""), ())


class ProcessorRuleFilesTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="audit_processor_"))
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.source = self.tmp / "note.docx"
        document = Document()
        document.add_paragraph("Износ 52 %, температура 64°C, котельная №1.")
        document.save(str(self.source))

    def run_processor(self, *extra: str) -> tuple[str, str]:
        out = self.tmp / "out.docx"
        result = subprocess.run(
            [str(PYTHON), str(PROCESSOR), "--file", str(self.source), "--fix", "--out", str(out),
             "--report", str(self.tmp / "report.md"), "--no-docx-report", *extra],
            capture_output=True, text=True, encoding="utf-8",
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        with zipfile.ZipFile(out) as archive:
            root = etree.fromstring(archive.read("word/document.xml"))
        return "".join(node.text or "" for node in root.iter(W + "t")), (self.tmp / "report.md").read_text(encoding="utf-8")

    def test_default_files_and_house_norm(self) -> None:
        text, report = self.run_processor()
        self.assertEqual(text, "Износ 52%, температура 64°C, котельная № 1.")
        self.assertIn("ПРАВИЛА", report)
        self.assertIn("units.yaml, house.yaml, abbreviations.yaml, acronyms.yaml", report)

    def test_gost_norm_and_only_the_chosen_file(self) -> None:
        text, report = self.run_processor("--norm", "gost", "--rule-sets", "units")
        self.assertEqual(text, f"Износ 52{NBSP}%, температура 64{NBSP}°C, котельная №1.")
        self.assertIn("ГОСТ", report)
        self.assertNotIn("AUDIT-NUMBER-SIGN", report)

    def test_no_rule_files_leaves_the_text(self) -> None:
        text, _ = self.run_processor("--rule-sets=")
        self.assertEqual(text, "Износ 52 %, температура 64°C, котельная №1.")

    def test_unknown_rule_file_stops_with_an_error(self) -> None:
        result = subprocess.run(
            [str(PYTHON), str(PROCESSOR), "--file", str(self.source), "--report", str(self.tmp / "r.md"),
             "--no-docx-report", "--rule-sets", "spelling"],
            capture_output=True, text=True, encoding="utf-8",
        )
        self.assertEqual(result.returncode, 2)
        self.assertIn("spelling", result.stdout)


class AuditWindowTests(unittest.TestCase):
    """The checkboxes and the radio of the audit window reach the processor."""

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="audit_window_"))
        self.addCleanup(shutil.rmtree, self.tmp, True)
        for name in ("config", "input", "output", "logs", "report", "workspace", "system_core"):
            (self.tmp / name).mkdir(parents=True, exist_ok=True)
        (self.tmp / "system_core" / "docx_audit_processor.py").write_text("", encoding="utf-8")
        self.commands: list[list[str]] = []

    def run_window(self, **parameters) -> list[str]:
        def fake_run_command(context, command, **kwargs):
            self.commands.append([str(part) for part in command])
            return {"exit_code": 0}

        original = office_service._run_command
        office_service._run_command = fake_run_command
        try:
            context = JobContext(
                paths=get_project_paths(self.tmp),
                operation=Operation(
                    id="docx_audit_processor",
                    title="audit",
                    description="",
                    service="system_core.services.office_service:docx_audit_processor",
                    parameters=parameters,
                ),
                log_file=self.tmp / "logs" / "audit.log",
                report_dir=self.tmp / "report",
            )
            office_service.docx_audit_processor(context)
        finally:
            office_service._run_command = original
        return self.commands[-1]

    def test_defaults_take_every_file_but_addresses_and_the_house_norm(self) -> None:
        command = self.run_window()
        self.assertIn("--rule-sets=units,house,abbreviations,acronyms", command)
        self.assertEqual(command[command.index("--norm") + 1], "rules")

    def test_checkboxes_and_radio(self) -> None:
        command = self.run_window(rules_units=False, rules_addresses=True, audit_norm="gost")
        self.assertIn("--rule-sets=house,abbreviations,acronyms,addresses", command)
        self.assertEqual(command[command.index("--norm") + 1], "gost")
        command = self.run_window(rules_units=False, rules_house=False, rules_abbreviations=False, rules_acronyms=False)
        self.assertIn("--rule-sets=", command)

    def test_manifest_fields_match_the_service(self) -> None:
        manifest = (ROOT / "config" / "tool_manifest.yaml").read_text(encoding="utf-8")
        for _, field, _ in office_service.AUDIT_RULE_SET_FIELDS:
            self.assertEqual(manifest.count(f"id: {field}\n"), 1, field)
        self.assertEqual(manifest.count("id: audit_norm\n"), 1)
        self.assertEqual(manifest.count("id: after_audit_norm\n"), 1)


class AddressWritingTests(unittest.TestCase):
    """Addresses are written by the Audion Address Processor engine, abbreviations and spaces only."""

    ENGINE_SOURCE = ROOT.parent / "Audion Address Processor" / "system_core" / "address_engine"
    ENGINE_FILES = ("audion_address_core.py", "address_slots.py", "models.py")

    def test_the_copy_holds_only_the_naming_modules(self) -> None:
        names = sorted(path.name for path in (SYSTEM_CORE / "address_engine").glob("*.py"))
        self.assertEqual(names, sorted(("__init__.py", *self.ENGINE_FILES)))

    @unittest.skipUnless(
        (Path(__file__).resolve().parents[2] / "Audion Address Processor" / "system_core" / "address_engine").is_dir(),
        "Audion Address Processor is not next to DocFlow",
    )
    def test_the_copy_matches_address_processor(self) -> None:
        for name in self.ENGINE_FILES:
            self.assertEqual(
                (SYSTEM_CORE / "address_engine" / name).read_bytes(),
                (self.ENGINE_SOURCE / name).read_bytes(),
                f"{name} differs from Audion Address Processor: copy it over unchanged",
            )

    def test_only_abbreviations_and_spaces_change(self) -> None:
        written = rules.address_writing
        self.assertEqual(written("п.Московский, ул.Новая, 1б"), "п. Московский, ул. Новая, д. 1б")
        self.assertEqual(written("станица Каневская, улица Горького, дом 5"), "ст-ца Каневская, ул. Горького, д. 5")
        self.assertEqual(written("посёлок Московский, ул. Новая, 1 б"), "п. Московский, ул. Новая, д. 1б")
        for kept in (
            "сл.Никольская, пер.Садовый, 3",  # the engine loses the name
            "ул. Мира, д. 2, кв. 5",  # the engine loses the flat
            "пер. Солнечный, 3/1",  # a slash is not a corpus
        ):
            self.assertEqual(written(kept), kept)

    def test_rule_without_the_engine_is_skipped_and_the_rest_run(self) -> None:
        with mock.patch.object(rules, "_address_engine", side_effect=rules.RuleError("нет движка")):
            book = rules.load_rule_book(["addresses"], "rules")
        codes = [rule.code for rule in book.rules]
        self.assertNotIn("AUDIT-ADDR-WRITING", codes)
        self.assertIn("AUDIT-ADDR-TYPE-SPACE", codes)
        self.assertTrue(any("AUDIT-ADDR-WRITING" in warning for warning in book.warnings))


if __name__ == "__main__":
    unittest.main()
