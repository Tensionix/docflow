"""The house standard: margins, font, body size, table size and the overflow size."""
from __future__ import annotations

from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
PYTHON = ROOT / "runtime" / "python.exe"
SCRIPT = ROOT / "system_core" / "docx_apply_standard.py"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS


def run(text: str, font: str = "Times New Roman", size: int = 28) -> str:
    return ('<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/><w:sz w:val="%d"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r>' % (font, font, size, text))


def cell(text: str) -> str:
    return "<w:tc><w:p>%s</w:p></w:tc>" % run(text, size=22)


def table(columns: int, word: str, declared: str = "") -> str:
    grid = "".join('<w:gridCol w:w="700"/>' for _ in range(columns))
    rows = "".join("<w:tr>%s</w:tr>" % "".join(cell(word) for _ in range(columns)) for _ in range(2))
    return "<w:tbl><w:tblPr>%s</w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>" % (declared, grid, rows)


def build(path: Path) -> None:
    body = (
        '<w:p><w:pPr><w:pStyle w:val="1"/></w:pPr>%s</w:p>' % run("Глава", size=32)
        + "<w:p>%s</w:p>" % run("Основной текст")
        + "<w:p>%s</w:p>" % run("", font="Symbol")
        + '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr></w:p>'
        + table(2, "коротко", '<w:tblW w:w="5411" w:type="pct"/>')
        + '<w:p><w:pPr><w:sectPr><w:pgSz w:w="23811" w:h="16838" w:orient="landscape"/>'
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr></w:p>'
        + table(20, "длинноесловобезпереносовдлянагрузки")
        + '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
          '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
    )
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="%s"><w:body>%s</w:body></w:document>' % (W_NS, body))
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="%s">'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
              '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="20"/></w:rPr></w:style>'
              '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/>'
              '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>' % W_NS)
    parts = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        "</Types>",
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>",
        "word/_rels/document.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>",
        "word/document.xml": document,
        "word/styles.xml": styles,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


class ApplyStandardTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not PYTHON.is_file():
            raise unittest.SkipTest("project runtime is unavailable")
        cls.tmp = Path(tempfile.mkdtemp(prefix="standard_test_"))
        build(cls.tmp / "in" / "doc.docx")
        cls.report = cls.tmp / "report.md"
        cls.result = subprocess.run(
            [str(PYTHON), str(SCRIPT), "--input", str(cls.tmp / "in"), "--outdir", str(cls.tmp / "out"),
             "--report", str(cls.report)],
            capture_output=True, text=True, encoding="utf-8",
        )
        with zipfile.ZipFile(cls.tmp / "out" / "doc.docx") as archive:
            cls.document = etree.fromstring(archive.read("word/document.xml"))
            cls.styles = etree.fromstring(archive.read("word/styles.xml"))

    @classmethod
    def tearDownClass(cls) -> None:
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def sizes(self, element) -> set[str]:
        return {sz.get(W + "val") for sz in element.iter(W + "sz")}

    def test_runs_cleanly(self) -> None:
        self.assertEqual(self.result.returncode, 0, self.result.stdout + self.result.stderr)

    def test_a4_landscape_turns_the_narrow_margin_down(self) -> None:
        margins = self.document.find(f"{W}body/{W}sectPr/{W}pgMar")
        self.assertEqual(
            {key: margins.get(W + key) for key in ("top", "bottom", "left", "right")},
            {"top": "1134", "bottom": "850", "left": "1134", "right": "1134"},
        )
        self.assertEqual(margins.get(W + "header"), "708")

    def test_portrait_and_a3_keep_the_narrow_margin_on_the_right(self) -> None:
        breaks = self.document.findall(f"{W}body/{W}p/{W}pPr/{W}sectPr/{W}pgMar")
        self.assertEqual(len(breaks), 2)
        for margins in breaks:
            self.assertEqual(
                {key: margins.get(W + key) for key in ("top", "bottom", "left", "right")},
                {"top": "1134", "bottom": "1134", "left": "1134", "right": "850"},
            )

    def test_body_text_and_headings(self) -> None:
        paragraphs = self.document.findall(f"{W}body/{W}p")
        self.assertEqual(self.sizes(paragraphs[0]), {"32"})       # heading keeps its size
        self.assertEqual(self.sizes(paragraphs[1]), {"24"})       # body text: 12 pt
        fonts = paragraphs[1].find(f"{W}r/{W}rPr/{W}rFonts")
        self.assertEqual(fonts.get(W + "ascii"), "Tahoma")
        symbol = paragraphs[2].find(f"{W}r/{W}rPr/{W}rFonts")
        self.assertEqual(symbol.get(W + "ascii"), "Symbol")       # bullets keep their font

    def test_tables_take_the_table_size_or_the_overflow_size(self) -> None:
        narrow, wide = self.document.findall(f"{W}body/{W}tbl")
        self.assertEqual(self.sizes(narrow), {"20"})
        self.assertEqual(self.sizes(wide), {"18"})
        self.assertIn("тесных, с уменьшенным кеглем: 1", self.report.read_text(encoding="utf-8"))

    def test_a_table_declared_wider_than_the_text_is_narrowed(self) -> None:
        narrow, _wide = self.document.findall(f"{W}body/{W}tbl")
        declared = narrow.find(f"{W}tblPr/{W}tblW")
        self.assertEqual((declared.get(W + "w"), declared.get(W + "type")), ("5000", "pct"))
        self.assertIn("сужено до поля: 1", self.report.read_text(encoding="utf-8"))

    def test_styles_carry_the_font_and_the_body_size(self) -> None:
        default_fonts = self.styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr/{W}rFonts")
        self.assertEqual(default_fonts.get(W + "ascii"), "Tahoma")
        normal = self.styles.find(f"{W}style[@{W}styleId='a']/{W}rPr")
        self.assertIsNone(normal.find(W + "rFonts").get(W + "asciiTheme"))
        self.assertEqual(normal.find(W + "sz").get(W + "val"), "24")


if __name__ == "__main__":
    unittest.main()
