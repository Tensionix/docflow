#!/usr/bin/env python3
"""Remove audit anchors from annotated DOCX copies."""

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
import re
from dataclasses import dataclass
from pathlib import Path

from _office_common import md_escape, safe_mkdir, write_json_file
from docx_xml_tools import NS, _etree_from_bytes, _etree_to_bytes, list_xml_parts, read_zip_map, write_zip_map


XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ERROR_ANCHOR_RE = re.compile(r"\s*⟦[^⟧]{1,80}⟧")


@dataclass
class StripResult:
    source: Path
    output: Path
    status: str
    removed: int = 0
    error: str = ""


def _set_text_preserve_space(node, value: str) -> None:
    node.text = value
    if value.startswith(" ") or value.endswith(" "):
        node.set(XML_SPACE, "preserve")


def strip_anchors_docx(source: Path, output: Path) -> int:
    files = read_zip_map(source)
    removed = 0
    modified = False
    for part in list_xml_parts(files):
        if part not in files:
            continue
        tree = _etree_from_bytes(files[part])
        part_modified = False
        for text_node in tree.getroot().xpath(".//w:t", namespaces=NS):
            text = text_node.text or ""
            if not text:
                continue
            cleaned, count = ERROR_ANCHOR_RE.subn("", text)
            if count:
                _set_text_preserve_space(text_node, cleaned)
                removed += count
                part_modified = True
        if part_modified:
            files[part] = _etree_to_bytes(tree)
            modified = True
    safe_mkdir(output.parent)
    write_zip_map(output, files)
    return removed if modified else 0


def iter_input_files(path: Path, *, all_docx: bool) -> tuple[Path, list[Path]]:
    resolved = path.resolve()
    if resolved.is_file():
        eligible = resolved.suffix.lower() == ".docx" and not resolved.name.startswith("~$")
        return resolved.parent, [resolved] if eligible else []
    return resolved, sorted(
        file
        for file in resolved.rglob("*.docx")
        if file.is_file()
        and not file.name.startswith("~$")
        and not file.stem.endswith("__unanchored")
        and (all_docx or file.stem.endswith("__annotated"))
    )


def output_path_for(source: Path, input_root: Path, out_dir: Path) -> Path:
    relative = source.name if source.is_file() and input_root == source.parent else source.relative_to(input_root)
    base = out_dir / relative
    stem = base.stem
    if stem.endswith("__annotated"):
        stem = stem[: -len("__annotated")]
    return base.with_name(f"{stem}__unanchored{base.suffix}")


def write_report(out_path: Path, input_root: Path, results: list[StripResult]) -> None:
    safe_mkdir(out_path.parent)
    lines = ["# Снятие audit-якорей DOCX\n"]
    lines.append(f"- Вход: `{md_escape(str(input_root))}`")
    lines.append(f"- Файлов обработано: **{len(results)}**")
    lines.append(f"- Якорей удалено: **{sum(item.removed for item in results)}**")
    lines.append("")
    lines.append("| Файл | Статус | Удалено якорей | Выходной файл | Ошибка |")
    lines.append("|---|---|---:|---|---|")
    for item in results:
        lines.append(
            f"| `{md_escape(str(item.source))}` | `{item.status}` | {item.removed} | `{md_escape(str(item.output))}` | {md_escape(item.error)} |"
        )
    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def build_json(input_root: Path, results: list[StripResult]) -> dict[str, object]:
    return {
        "tool": "docx_audit_anchors",
        "input_root": str(input_root),
        "summary": {
            "files": len(results),
            "ok": sum(1 for item in results if item.status == "OK"),
            "failed": sum(1 for item in results if item.status != "OK"),
            "removed": sum(item.removed for item in results),
        },
        "files": [
            {
                "source": str(item.source),
                "output": str(item.output),
                "status": item.status,
                "removed": item.removed,
                "error": item.error,
            }
            for item in results
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Remove service anchors from __annotated DOCX files.")
    parser.add_argument("--input", default="input", help="Input DOCX file or folder")
    parser.add_argument("--outdir", default="output/audit_unanchored", help="Output folder")
    parser.add_argument("--report", default="report/docx_audit_anchors.md", help="Markdown report path")
    parser.add_argument("--json-out", default="", help="Optional JSON report path")
    parser.add_argument("--all-docx", action="store_true", help="Process all DOCX files, not only *__annotated.docx")
    args = parser.parse_args()

    input_root, files = iter_input_files(Path(args.input), all_docx=args.all_docx)
    report_path = Path(args.report).resolve()
    out_dir = Path(args.outdir).resolve()
    results: list[StripResult] = []

    if not input_root.exists():
        print(f"[ERROR] Input does not exist: {input_root}")
        return 2
    if not files:
        safe_mkdir(report_path.parent)
        report_path.write_text("# Снятие audit-якорей DOCX\n\nФайлы `*__annotated.docx` не найдены.\n", encoding="utf-8")
        print(f"[WARN] No annotated DOCX files found: {input_root}")
        print(f"[OK] Report: {report_path}")
        return 0

    for source in files:
        output = output_path_for(source, input_root, out_dir)
        try:
            removed = strip_anchors_docx(source, output)
            results.append(StripResult(source=source, output=output, status="OK", removed=removed))
            print(f"[OK] {source} -> {output} removed={removed}")
        except Exception as exc:
            results.append(StripResult(source=source, output=output, status="FAILED", error=str(exc)))
            print(f"[FAILED] {source}: {exc}")

    write_report(report_path, input_root, results)
    if args.json_out:
        json_path = Path(args.json_out).resolve()
        write_json_file(json_path, build_json(input_root, results))
        print(f"[OK] JSON: {json_path}")
    print(f"[OK] Report: {report_path}")
    return 1 if any(item.status != "OK" for item in results) else 0


if __name__ == "__main__":
    raise SystemExit(main())
