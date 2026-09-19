"""A table wider than its text is always fitted when the formatting pass asks for it.

Without the flag the optimizer keeps its skip rules: a table on the first two
pages or with three columns or fewer stays as it is, even if it sticks out.
"""
from __future__ import annotations

import json
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
PYTHON = ROOT / "runtime" / "python.exe"
SCRIPT = ROOT / "system_core" / "docx_table_width_optimizer.py"
SUFFIX = "__fit_to_margins_optimized_widths"


def build(path: Path) -> None:
    from docx import Document
    from docx.oxml.ns import qn

    document = Document()
    table = document.add_table(rows=4, cols=2)
    for row_index, row in enumerate(table.rows):
        for column_index, cell in enumerate(row.cells):
            cell.text = f"значение {row_index}.{column_index} с поясняющим текстом"
    for column in table._tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol")):
        column.set(qn("w:w"), "6000")  # 2 x 6000 twips = 21.2 cm, wider than the text of the page
    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(path))


class FitOverflowingTests(unittest.TestCase):
    def setUp(self) -> None:
        if not PYTHON.is_file():
            self.skipTest("project runtime is unavailable")
        self.tmp = Path(tempfile.mkdtemp(prefix="fit_overflow_"))
        self.addCleanup(shutil.rmtree, self.tmp, True)
        build(self.tmp / "in" / "doc.docx")

    def run_optimizer(self, *extra: str) -> dict:
        report_dir = self.tmp / ("report" + "".join(extra).replace("-", "_"))
        result = subprocess.run(
            [str(PYTHON), str(SCRIPT), "--input-dir", str(self.tmp / "in"), "--outdir", str(self.tmp / "out"),
             "--report-dir", str(report_dir), "--all", "--recursive", "--mode", "fit-to-margins",
             "--fit-target", "current-section", *extra],
            capture_output=True, text=True, encoding="utf-8",
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        return json.loads((report_dir / f"doc{SUFFIX}.json").read_text(encoding="utf-8"))

    def test_skip_rules_hold_without_the_flag(self) -> None:
        report = self.run_optimizer()
        self.assertEqual(report["tables"], [])
        self.assertEqual(report["skipped_tables"][0]["reason"], "skip_first_two_pages")

    def test_an_overflowing_table_is_fitted_with_the_flag(self) -> None:
        report = self.run_optimizer("--fit-overflowing")
        self.assertEqual(len(report["tables"]), 1)
        fitted = report["tables"][0]
        self.assertTrue(fitted["forced_overflow"])
        self.assertLess(sum(fitted["optimized_widths_cm"]), 21.0)
        self.assertAlmostEqual(sum(fitted["optimized_widths_cm"]), fitted["target_total_width_cm"], delta=0.1)


if __name__ == "__main__":
    unittest.main()
