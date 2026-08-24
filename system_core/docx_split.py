#!/usr/bin/env python3
"""
Split DOCX files into smaller DOCX packages.

The splitter can keep top-level logical sections together or split the raw
document body into a requested number of chunks. It copies the original DOCX
package and replaces only word/document.xml in each output piece.
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
from copy import deepcopy
from dataclasses import dataclass
import json
from pathlib import Path
import re
import sys
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _office_common import safe_mkdir
from docx_pair_diff import Section, build_section, extract_content_blocks, split_sections
from docx_word_compare import _clean_docx_copy


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_BODY = f"{{{W_NS}}}body"
W_P = f"{{{W_NS}}}p"
W_SECT_PR = f"{{{W_NS}}}sectPr"
TOP_SECTION_RE = re.compile(r"^\s*(?:раздел\s+)?(?:\d+|[IVXLCDM]+)(?:(?:[.)](?!\d))|\s+)\s+\S", re.IGNORECASE)
BAD_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')


@dataclass
class DocxPackage:
    document_info: ZipInfo
    other_entries: list[tuple[ZipInfo, bytes]]
    root_template: etree._Element
    body_children: list[etree._Element]
    sect_pr: etree._Element | None


def _docx_files(input_dir: Path) -> list[Path]:
    if input_dir.is_file():
        return [input_dir] if input_dir.suffix.lower() == ".docx" and not input_dir.name.startswith("~$") else []
    return sorted(
        path
        for path in input_dir.rglob("*.docx")
        if path.is_file() and not path.name.startswith("~$")
    )


def _parse_chunk_count(value: str, *, required: bool) -> int | None:
    text = (value or "").strip()
    if not text:
        if required:
            raise RuntimeError("Для механической резки укажите количество кусков или включите резку по верхним разделам.")
        return None
    try:
        parsed = int(float(text.replace(",", ".")))
    except ValueError as exc:
        raise RuntimeError("Количество кусков должно быть целым числом или пустым полем.") from exc
    if parsed <= 0:
        raise RuntimeError("Количество кусков должно быть больше нуля или пустым полем.")
    return parsed


def _safe_name(value: str, fallback: str) -> str:
    text = BAD_FILENAME_RE.sub("_", (value or "").strip())
    text = re.sub(r"\s+", " ", text).strip(" ._")
    if not text:
        text = fallback
    return text[:90].rstrip(" ._") or fallback


def _body_ranges(sections: list[Section]) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    for section in sections:
        indices = [block.body_index for block in section.blocks if block.body_index >= 0]
        if indices:
            ranges.append((min(indices), max(indices)))
    ranges.sort()
    merged: list[tuple[int, int]] = []
    for start, end in ranges:
        if not merged or start > merged[-1][1] + 1:
            merged.append((start, end))
        else:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
    return merged


def _contains_index(ranges: list[tuple[int, int]], index: int) -> bool:
    return any(start <= index <= end for start, end in ranges)


def _load_docx_package(src: Path) -> DocxPackage:
    with ZipFile(src, "r") as zin:
        document_info: ZipInfo | None = None
        document_xml: bytes | None = None
        other_entries: list[tuple[ZipInfo, bytes]] = []
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                document_info = info
                document_xml = data
            else:
                other_entries.append((info, data))

    if document_info is None or document_xml is None:
        raise RuntimeError("DOCX body not found: word/document.xml is missing")

    root = etree.fromstring(document_xml)
    body = root.find(W_BODY)
    if body is None:
        raise RuntimeError("DOCX body not found in word/document.xml")

    body_children = list(body)
    sect_pr = next((deepcopy(child) for child in reversed(body_children) if child.tag == W_SECT_PR), None)
    for child in body_children:
        body.remove(child)
    return DocxPackage(
        document_info=document_info,
        other_entries=other_entries,
        root_template=root,
        body_children=body_children,
        sect_pr=sect_pr,
    )


def _package_document_xml(package: DocxPackage, ranges: list[tuple[int, int]]) -> bytes:
    root = deepcopy(package.root_template)
    body = root.find(W_BODY)
    if body is None:
        raise RuntimeError("DOCX body not found in word/document.xml")

    selected = [
        deepcopy(child)
        for index, child in enumerate(package.body_children)
        if child.tag != W_SECT_PR and _contains_index(ranges, index)
    ]
    if not selected:
        selected = [etree.Element(W_P)]
    sect_pr = deepcopy(package.sect_pr) if package.sect_pr is not None else etree.Element(W_SECT_PR)

    for child in selected:
        body.append(child)
    body.append(sect_pr)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=False)


def _write_package_ranges(package: DocxPackage, dst: Path, ranges: list[tuple[int, int]]) -> None:
    safe_mkdir(dst.parent)
    with ZipFile(dst, "w", ZIP_DEFLATED, compresslevel=1) as zout:
        for info, data in package.other_entries:
            zout.writestr(info, data)
        zout.writestr(package.document_info, _package_document_xml(package, ranges))


def _replace_document_body(xml_bytes: bytes, ranges: list[tuple[int, int]]) -> bytes:
    root = etree.fromstring(xml_bytes)
    body = root.find(W_BODY)
    if body is None:
        raise RuntimeError("DOCX body not found in word/document.xml")

    original_children = list(body)
    sect_pr = next((deepcopy(child) for child in reversed(original_children) if child.tag == W_SECT_PR), None)
    selected = [
        deepcopy(child)
        for index, child in enumerate(original_children)
        if child.tag != W_SECT_PR and _contains_index(ranges, index)
    ]
    if not selected:
        selected = [etree.Element(W_P)]
    if sect_pr is None:
        sect_pr = etree.Element(W_SECT_PR)

    for child in original_children:
        body.remove(child)
    for child in selected:
        body.append(child)
    body.append(sect_pr)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=False)


def _write_docx_ranges(src: Path, dst: Path, ranges: list[tuple[int, int]]) -> None:
    safe_mkdir(dst.parent)
    with ZipFile(src, "r") as zin, ZipFile(dst, "w", ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                data = _replace_document_body(data, ranges)
            zout.writestr(info, data)


def _body_content_ranges(package: DocxPackage, chunks: int) -> list[list[tuple[int, int]]]:
    indices = [index for index, child in enumerate(package.body_children) if child.tag != W_SECT_PR]
    if not indices:
        return [[(0, 0)]]
    chunks = max(1, min(chunks, len(indices)))
    groups: list[list[tuple[int, int]]] = []
    for chunk_index in range(chunks):
        start_pos = round(chunk_index * len(indices) / chunks)
        end_pos = round((chunk_index + 1) * len(indices) / chunks)
        piece = indices[start_pos:end_pos]
        if piece:
            groups.append([(piece[0], piece[-1])])
    return groups


def _numbered_top_sections(docx_path: Path) -> tuple[list[Section], str]:
    blocks = extract_content_blocks(docx_path)
    style_sections = split_sections(blocks, depth=1)
    if len(style_sections) > 1 and any(section.key for section in style_sections):
        return style_sections, "heading-style"

    sections: list[Section] = []
    current_key: tuple[str, ...] = ()
    current_blocks = []
    numbered_hits = 0
    for block in blocks:
        if block.kind in {"heading", "paragraph"} and TOP_SECTION_RE.match(block.text or ""):
            if current_blocks:
                sections.append(build_section(len(sections) + 1, current_key, current_blocks))
            numbered_hits += 1
            current_key = (block.text,)
            current_blocks = [block]
            continue
        current_blocks.append(block)

    if current_blocks:
        sections.append(build_section(len(sections) + 1, current_key, current_blocks))

    if numbered_hits >= 2 and len(sections) > 1:
        return sections, "numbered-text"
    return style_sections or sections, "single-section"


def _group_sections(sections: list[Section], chunks: int | None) -> list[list[Section]]:
    if not sections:
        return []
    if chunks is None:
        return [[section] for section in sections]

    chunk_count = max(1, min(chunks, len(sections)))
    total_weight = sum(max(section.word_count, 1) for section in sections)
    target_weight = max(1, total_weight // chunk_count)
    groups: list[list[Section]] = []
    current: list[Section] = []
    current_weight = 0
    for index, section in enumerate(sections):
        remaining_sections = len(sections) - index
        remaining_groups = chunk_count - len(groups) - 1
        if current and (current_weight >= target_weight or remaining_sections <= remaining_groups):
            groups.append(current)
            current = []
            current_weight = 0
        current.append(section)
        current_weight += max(section.word_count, 1)
    if current:
        groups.append(current)
    while len(groups) > chunk_count:
        tail = groups.pop()
        groups[-1].extend(tail)
    return groups


def _piece_title(sections: list[Section], ordinal: int) -> str:
    titles = [section.title for section in sections if section.title and section.title != "Без раздела"]
    if not titles:
        return f"Кусок {ordinal}"
    if len(titles) == 1:
        return titles[0]
    return f"{titles[0]} +{len(titles) - 1}"


def _split_by_sections(docx_path: Path, target_dir: Path, chunks: int | None) -> tuple[list[dict[str, object]], str, int]:
    package = _load_docx_package(docx_path)
    sections, detector = _numbered_top_sections(docx_path)
    groups = _group_sections(sections, chunks)
    pieces: list[dict[str, object]] = []
    for ordinal, group in enumerate(groups, start=1):
        title = _piece_title(group, ordinal)
        filename = f"{ordinal:03d}_{_safe_name(title, f'chunk_{ordinal:03d}')}.docx"
        out_path = target_dir / filename
        _write_package_ranges(package, out_path, _body_ranges(group))
        pieces.append(
            {
                "ordinal": ordinal,
                "title": title,
                "path": str(out_path),
                "sections": [section.title for section in group],
            }
        )
    return pieces, detector, len(sections)


def _split_by_body(docx_path: Path, target_dir: Path, chunks: int) -> list[dict[str, object]]:
    package = _load_docx_package(docx_path)
    groups = _body_content_ranges(package, chunks)
    pieces: list[dict[str, object]] = []
    for ordinal, ranges in enumerate(groups, start=1):
        out_path = target_dir / f"{ordinal:03d}_body_chunk.docx"
        _write_package_ranges(package, out_path, ranges)
        pieces.append(
            {
                "ordinal": ordinal,
                "title": f"Кусок {ordinal}",
                "path": str(out_path),
                "ranges": ranges,
            }
        )
    return pieces


def _write_reports(report_path: Path, json_path: Path, rows: list[dict[str, object]], outdir: Path) -> None:
    safe_mkdir(report_path.parent)
    safe_mkdir(json_path.parent)
    ok_count = sum(1 for row in rows if row.get("status") == "OK")
    fail_count = len(rows) - ok_count
    lines = [
        "# Разрезка DOCX",
        "",
        f"- Выходная папка: `{outdir}`",
        f"- Документов успешно: {ok_count}",
        f"- Ошибок: {fail_count}",
        "",
        "## Режимы",
        "",
        "- По верхним разделам: режет по Heading 1 или по текстовым заголовкам вида `1. ...`, `2. ...`, `3. ...`.",
        "- По числу кусков: группирует разделы примерно в указанное число частей.",
        "- Без верхних разделов: механически делит тело DOCX на указанное число частей.",
        "- Перед анализом делается техническая cleaned-copy, чтобы убрать битые ссылки на отсутствующие comment-парты.",
        "",
        "## Документы",
        "",
        "| Статус | Файл | Режим | Разделов | Кусков | Папка |",
        "|---|---|---|---:|---:|---|",
    ]
    for row in rows:
        source = str(row.get("source", "")).replace("|", "\\|")
        target = str(row.get("target_dir", row.get("error", ""))).replace("|", "\\|")
        lines.append(
            "| {status} | `{source}` | {mode} | {sections} | {pieces} | `{target}` |".format(
                status=row.get("status", "FAILED"),
                source=source,
                mode=row.get("mode", ""),
                sections=row.get("sections", 0),
                pieces=row.get("pieces", 0),
                target=target,
            )
        )
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    json_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Split DOCX files from input into smaller DOCX packages.")
    parser.add_argument("--input", default="input", help="Input folder")
    parser.add_argument("--outdir", default="output/docx_split", help="Output folder")
    parser.add_argument("--report", default="report/docx_split.md", help="Markdown report path")
    parser.add_argument("--json-out", default="report/docx_split.json", help="JSON report path")
    parser.add_argument("--chunks", default="", help="Optional target chunk count")
    parser.add_argument("--top-level-sections", action="store_true", help="Split or group by top-level sections")
    args = parser.parse_args()

    input_dir = Path(args.input).resolve()
    outdir = Path(args.outdir).resolve()
    report_path = Path(args.report).resolve()
    json_path = Path(args.json_out).resolve()
    safe_mkdir(outdir)

    chunks = _parse_chunk_count(args.chunks, required=not args.top_level_sections)
    files = _docx_files(input_dir)
    if not files:
        raise FileNotFoundError(f"В input нет DOCX файлов: {input_dir}")

    rows: list[dict[str, object]] = []
    clean_dir = report_path.parent / "docx_split_cleaned"
    input_base = input_dir.parent if input_dir.is_file() else input_dir
    for source in files:
        relative = source.relative_to(input_base)
        target_dir = outdir / relative.parent / source.stem
        try:
            clean_source = clean_dir / relative
            clean_stats = _clean_docx_copy(source, clean_source)
            if args.top_level_sections:
                pieces, detector, section_count = _split_by_sections(clean_source, target_dir, chunks)
                mode = f"top-level-sections:{detector}"
            else:
                pieces = _split_by_body(clean_source, target_dir, int(chunks or 1))
                section_count = 0
                mode = "body"
            rows.append(
                {
                    "status": "OK",
                    "source": str(source),
                    "clean_source": str(clean_source),
                    "cleaned": clean_stats,
                    "target_dir": str(target_dir),
                    "mode": mode,
                    "requested_chunks": chunks,
                    "sections": section_count,
                    "pieces": len(pieces),
                    "files": pieces,
                }
            )
            print(f"[OK] {source} -> {target_dir} ({len(pieces)} chunks)")
        except Exception as exc:
            rows.append(
                {
                    "status": "FAILED",
                    "source": str(source),
                    "target_dir": str(target_dir),
                    "error": str(exc),
                }
            )
            print(f"[FAILED] {source}: {exc}")

    _write_reports(report_path, json_path, rows, outdir)
    print(f"[OK] Report: {report_path}")
    print(f"[OK] JSON: {json_path}")
    return 1 if any(row.get("status") != "OK" for row in rows) else 0


if __name__ == "__main__":
    raise SystemExit(main())
