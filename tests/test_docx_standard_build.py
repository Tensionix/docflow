"""Formatting standards as data: the standard files and what is built from them.

config/standards holds the page, the corporate style set and the numbering of each
standard, every value with its basis. The builder turns a file into a Word template,
a sample document and the review description; these tests hold it to the files.
"""
from __future__ import annotations

from pathlib import Path
import shutil
import sys
import tempfile
import unittest
import zipfile

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
SYSTEM_CORE = ROOT / "system_core"
if str(SYSTEM_CORE) not in sys.path:
    sys.path.insert(0, str(SYSTEM_CORE))

import docx_standard_build as build  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
STANDARDS = ROOT / "config" / "standards"

MINIMAL = """\
id: test
title: Тест
template: test.dotx
sample: test.docx
page:
  size: A4
  margins_mm: 20 20 20 15
defaults:
  font: Arial
styles:
  - name: Обычный
    font: Arial
"""


class TempDirMixin:
    def make_tmp(self) -> Path:
        tmp = Path(tempfile.mkdtemp(prefix="standard_build_"))
        self.addCleanup(shutil.rmtree, tmp, True)
        return tmp


class StandardFilesTests(TempDirMixin, unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = self.make_tmp()

    def load(self, name: str) -> build.Standard:
        return build.load_standard(STANDARDS / name)

    def test_both_standards_hold_a_corporate_style_set(self) -> None:
        low, high = build.CORPORATE_STYLE_RANGE
        for name in ("notes.yaml", "nir_gost_7_32.yaml"):
            standard = self.load(name)
            self.assertTrue(low <= standard.style_count <= high, (name, standard.style_count))

    def test_house_standard_values_keep_their_basis(self) -> None:
        notes = self.load("notes.yaml")
        self.assertEqual(notes.page["margins_mm"].value, (20, 20, 20, 15))
        self.assertEqual(notes.page["margins_mm"].basis, "standard")
        base = notes.by_name["Обычный"].props
        self.assertEqual((base["font"].value, base["size"].value, base["line"].value), ("Tahoma", 12, 1))
        self.assertEqual({base[key].basis for key in ("font", "size", "line")}, {"standard"})
        table_size = notes.by_name["Табличный_центр"].props["size"]
        self.assertEqual((table_size.value, table_size.basis), (10, "standard"))
        self.assertEqual(build.basis_text(notes, "standard"), "Стандарт записок")

    def test_gost_values_cite_their_clauses(self) -> None:
        nir = self.load("nir_gost_7_32.yaml")
        self.assertEqual(nir.page["margins_mm"].value, (20, 20, 30, 15))
        text = nir.by_name["Абзац"].props
        self.assertEqual((text["line"].value, text["first_line"].value), (1.5, 1.25))
        self.assertEqual(text["line"].basis, "gost 6.1.1")
        self.assertEqual(build.basis_text(nir, "gost 6.2.3, 6.4.1"), "ГОСТ 7.32, пп. 6.2.3, 6.4.1")

    def test_template_is_a_word_template_with_the_style_set(self) -> None:
        nir = self.load("nir_gost_7_32.yaml")
        path = build.build_template(nir, self.tmp / nir.template)
        with zipfile.ZipFile(path) as archive:
            types = archive.read("[Content_Types].xml").decode("utf-8")
            styles = etree.fromstring(archive.read("word/styles.xml"))
            numbering = etree.fromstring(archive.read("word/numbering.xml"))
            document = etree.fromstring(archive.read("word/document.xml"))
            footer = etree.fromstring(archive.read("word/footer1.xml"))
            settings = etree.fromstring(archive.read("word/settings.xml"))
        self.assertIn(build.TEMPLATE_CONTENT_TYPE, types)
        defined = {style.get(W + "styleId"): style for style in styles.findall(W + "style")}
        self.assertEqual(len(defined), nir.style_count)
        self.assertEqual(defined["Heading1"].find(W + "name").get(W + "val"), "heading 1")
        paragraph = defined[build.style_id_for("Абзац")]
        self.assertEqual(paragraph.find(f"{W}pPr/{W}spacing").get(W + "line"), "360")
        self.assertEqual(paragraph.find(f"{W}pPr/{W}ind").get(W + "firstLine"), "709")
        level_texts = {level.find(W + "lvlText").get(W + "val") for level in numbering.iter(W + "lvl")}
        self.assertLessEqual({"%1", "%1.%2", "%1.%2.%3", "–", "%1)", "%1."}, level_texts)
        margins = document.find(f"{W}body/{W}sectPr/{W}pgMar")
        self.assertEqual((margins.get(W + "left"), margins.get(W + "right")), ("1701", "850"))
        self.assertIsNotNone(document.find(f"{W}body/{W}sectPr/{W}titlePg"))
        self.assertIn("PAGE", "".join(node.text or "" for node in footer.iter(W + "instrText")))
        self.assertEqual(footer.find(f"{W}p/{W}pPr/{W}jc").get(W + "val"), "center")
        modes = [node.get(W + "val") for node in settings.iter(W + "compatSetting")
                 if node.get(W + "name") == "compatibilityMode"]
        self.assertEqual(modes, ["15"])

    def test_sample_shows_headings_lists_and_captions(self) -> None:
        notes = self.load("notes.yaml")
        path = build.build_sample(notes, self.tmp / notes.sample)
        with zipfile.ZipFile(path) as archive:
            types = archive.read("[Content_Types].xml").decode("utf-8")
            document = etree.fromstring(archive.read("word/document.xml"))
        self.assertIn(build.DOCUMENT_CONTENT_TYPE, types)
        used = {node.get(W + "val") for node in document.iter(W + "pStyle")}
        for name in ("Абзац", "Заголовок 1", "Заголовок 4", "Маркированный список", "Нумерованный список 2",
                     "Название таблицы", "Табличный_заголовки", "Название объекта", "Заголовок приложения"):
            self.assertIn(notes.by_name[name].style_id, used, name)

    def test_description_marks_every_proposal_for_review(self) -> None:
        standards = [self.load("notes.yaml"), self.load("nir_gost_7_32.yaml")]
        path = self.tmp / "description.docx"
        rows, proposals = build.build_description(standards, path, "17.09.2026")
        self.assertGreater(proposals, 0)
        self.assertGreater(rows, proposals)
        with zipfile.ZipFile(path) as archive:
            document = archive.read("word/document.xml").decode("utf-8")
        self.assertEqual(document.count('w:fill="FFF2CC"'), proposals)
        self.assertIn("ГОСТ 7.32, п. 6.1.1", document)
        self.assertIn("Стандарт записок", document)


class StandardValidationTests(TempDirMixin, unittest.TestCase):
    def load_text(self, text: str) -> build.Standard:
        path = self.make_tmp() / "test.yaml"
        path.write_bytes(text.encode("utf-8"))
        return build.load_standard(path)

    def test_minimal_standard_loads(self) -> None:
        standard = self.load_text(MINIMAL)
        self.assertEqual(standard.style_count, 1 + len(build.WORD_DEFAULT_STYLES))

    def test_unknown_parameter_is_named(self) -> None:
        with self.assertRaises(build.StandardError) as caught:
            self.load_text(MINIMAL + "    bolt: yes\n")
        self.assertIn("bolt", str(caught.exception))

    def test_missing_base_style_is_named(self) -> None:
        with self.assertRaises(build.StandardError) as caught:
            self.load_text(MINIMAL + "  - name: Абзац\n    based_on: Основной\n")
        self.assertIn("Основной", str(caught.exception))

    def test_unknown_basis_is_rejected(self) -> None:
        with self.assertRaises(build.StandardError) as caught:
            self.load_text(MINIMAL.replace("font: Arial\nstyles", "font: Arial @gst 6.1\nstyles"))
        self.assertIn("gst", str(caught.exception))

    def test_numbering_must_exist(self) -> None:
        with self.assertRaises(build.StandardError) as caught:
            self.load_text(MINIMAL + "  - name: Заголовок 1\n    numbering: headings 1\n")
        self.assertIn("headings", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
