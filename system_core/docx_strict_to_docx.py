#!/usr/bin/env python3
"""
Strict Open XML DOCX -> ordinary DOCX through Microsoft Word.

Word can save a document as "Strict Open XML". Its XML lives in other namespaces
(purl.oclc.org/ooxml/...), while the DocFlow tools read the ordinary ones: on such a
file the XML cleanup failed and the text tools silently found nothing. Word resaves
every Strict file as an ordinary DOCX; other files are not written, the next step
takes them as they are. Without Word the Strict files are listed in the report.

Usage:
  python docx_strict_to_docx.py --input input --outdir output/strict_resaved --report report/docx_strict_to_docx.md
"""

from __future__ import annotations

# >>> audion CLI bootstrap >>>
import os as _os, sys as _sys
_HERE = _os.path.dirname(_os.path.abspath(__file__))
if _HERE not in _sys.path:
    _sys.path.insert(0, _HERE)
try:
    _sys.stdout.reconfigure(encoding="utf-8")
    _sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass
# <<< audion CLI bootstrap <<<

import argparse
import shutil
import subprocess
import tempfile
from pathlib import Path

from _office_common import find_docx_files, md_escape, mirrored_output_path, rel_posix, safe_mkdir, write_json_file
from docx_xml_tools import is_strict_ooxml


ROOT = Path(__file__).resolve().parents[1]
CONVERTER = Path(__file__).resolve().parent / "word_com" / "Convert-StrictDocxWord.ps1"
RESAVED = "пересохранён Word"


def _powershell() -> str | None:
    for candidate in (str(ROOT / "system_core" / "powershell" / "pwsh.exe"), "pwsh", "powershell"):
        if Path(candidate).exists():
            return candidate
        found = shutil.which(candidate)
        if found:
            return found
    return None


def convert_with_word(pairs: list[tuple[Path, Path]]) -> dict[Path, str]:
    """Resave each source as its target through one Word instance; a status per source."""
    shell = _powershell()
    if shell is None:
        return {source: "PowerShell не найден" for source, _ in pairs}
    statuses = {source: "Word не сообщил результат" for source, _ in pairs}
    by_target = {str(target): source for source, target in pairs}
    by_source = {str(source): source for source, _ in pairs}
    with tempfile.TemporaryDirectory(prefix="strict_docx_") as work:
        list_file = Path(work) / "files.txt"
        list_file.write_text("\n".join(f"{source}|{target}" for source, target in pairs), encoding="utf-8")
        for _, target in pairs:
            safe_mkdir(target.parent)
        process = subprocess.run(
            [shell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(CONVERTER), "-ListFile", str(list_file)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    for line in process.stdout.splitlines():
        print(line)
        if line.startswith("[WORD-UNAVAILABLE]"):
            reason = line[len("[WORD-UNAVAILABLE]"):].strip()
            return {source: f"Word недоступен: {reason}" for source, _ in pairs}
        if line.startswith("[OK] "):
            source = by_target.get(line[len("[OK] "):].strip())
            if source is not None:
                statuses[source] = RESAVED
        elif line.startswith("[FAILED] "):
            path_text, _, message = line[len("[FAILED] "):].partition(": ")
            source = by_source.get(path_text.strip())
            if source is not None:
                statuses[source] = f"ошибка Word: {message.strip()}"
    return statuses


def write_report(report_path: Path, input_root: Path, total: int, rows: list[dict[str, str]]) -> None:
    safe_mkdir(report_path.parent)
    lines = ["# Strict Open XML -> обычный DOCX", ""]
    lines.append(f"Документов в папке: {total}; в формате Strict Open XML: {len(rows)}.")
    lines.append("")
    if not rows:
        lines.append("Документов в формате Strict Open XML нет - пересохранять нечего.")
    else:
        lines.append("| Файл | Результат |")
        lines.append("|---|---|")
        lines.extend(f"| `{md_escape(row['file'])}` | {md_escape(row['status'])} |" for row in rows)
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Resave Strict Open XML DOCX files as ordinary DOCX through Microsoft Word.")
    parser.add_argument("--input", required=True, help="Input folder with .docx files (recursive)")
    parser.add_argument("--outdir", required=True, help="Output folder for the resaved files")
    parser.add_argument("--report", required=True, help="Markdown report path")
    parser.add_argument("--json-out", default="", help="Optional JSON report path")
    args = parser.parse_args()

    input_root = Path(args.input).resolve()
    out_dir = Path(args.outdir).resolve()
    documents = find_docx_files(input_root)
    pairs = [(path, mirrored_output_path(path, input_root, out_dir)) for path in documents if is_strict_ooxml(path)]
    statuses = convert_with_word(pairs) if pairs else {}
    rows = []
    for source, target in pairs:
        status = statuses[source]
        if status == RESAVED and (not target.is_file() or is_strict_ooxml(target)):
            status = "Word не пересохранил файл в обычный DOCX"
        rows.append({"file": rel_posix(source, input_root), "output": str(target) if status == RESAVED else "", "status": status})

    report_path = Path(args.report).resolve()
    write_report(report_path, input_root, len(documents), rows)
    if args.json_out:
        write_json_file(
            Path(args.json_out).resolve(),
            {
                "tool": "docx_strict_to_docx",
                "version": 1,
                "input": str(input_root),
                "outdir": str(out_dir),
                "summary": {
                    "documents": len(documents),
                    "strict": len(rows),
                    "resaved": sum(1 for row in rows if row["status"] == RESAVED),
                },
                "files": rows,
            },
        )
    print(f"[OK] Documents: {len(documents)}; Strict Open XML: {len(rows)}; resaved: {sum(1 for row in rows if row['status'] == RESAVED)}")
    print(f"[OK] Report: {report_path}")
    # A file Word could not resave is listed in the report; the pass goes on without it.
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
