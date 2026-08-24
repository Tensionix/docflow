from __future__ import annotations

import argparse
import copy
import json
import os
import sys
import tempfile
from collections import Counter
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from lxml import etree

CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from _office_common import md_escape, rel_posix, safe_mkdir, write_json_file
from docx_table_unifier_model import build_table_model, normalize_text
from docx_table_width_optimizer import (
    _apply_widths,
    _body_table_page_map,
    _body_table_section_map,
    _content_balanced_widths,
    _fallback_total_width_cm,
    _read_table_total_width_cm,
    _section_available_width_cm,
)


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W_NS = NS["w"]
TWIPS_PER_CM = 1440 / 2.54
SUPPORTED_EXTENSIONS = {".docx"}


def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def w_attr(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def cm_to_twips(value: float) -> int:
    return int(round(value * TWIPS_PER_CM))


def pt_to_half_points(value: float) -> str:
    return str(int(round(value * 2)))


def parse_xml(data: bytes) -> etree._Element:
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    return etree.fromstring(data, parser=parser)


def serialize_xml(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def get_or_add_child(parent: etree._Element, name: str, index: int | None = None) -> etree._Element:
    child = parent.find(f"w:{name}", namespaces=NS)
    if child is not None:
        return child
    child = etree.Element(w_tag(name))
    if index is None:
        parent.append(child)
    else:
        parent.insert(index, child)
    return child


def remove_children(parent: etree._Element, names: set[str]) -> int:
    removed = 0
    for child in list(parent):
        local = etree.QName(child).localname
        if child.tag.startswith(f"{{{W_NS}}}") and local in names:
            parent.remove(child)
            removed += 1
    return removed


def set_table_borders(tbl_pr: etree._Element, border_size: int, border_color: str) -> None:
    for existing in tbl_pr.findall("w:tblBorders", namespaces=NS):
        tbl_pr.remove(existing)
    borders = etree.Element(w_tag("tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = etree.SubElement(borders, w_tag(side))
        el.set(w_attr("val"), "single")
        el.set(w_attr("sz"), str(border_size))
        el.set(w_attr("space"), "0")
        el.set(w_attr("color"), border_color)
    tbl_pr.append(borders)


def set_table_cell_margins(parent: etree._Element, name: str, margin_twips: int) -> None:
    for existing in parent.findall(f"w:{name}", namespaces=NS):
        parent.remove(existing)
    mar = etree.Element(w_tag(name))
    for side in ("top", "start", "bottom", "end", "left", "right"):
        el = etree.SubElement(mar, w_tag(side))
        el.set(w_attr("w"), str(margin_twips))
        el.set(w_attr("type"), "dxa")
    parent.append(mar)


def set_child_attrs(parent: etree._Element, name: str, attrs: dict[str, str]) -> etree._Element:
    child = get_or_add_child(parent, name)
    for key, value in attrs.items():
        child.set(w_attr(key), value)
    return child


def set_run_font(rpr: etree._Element, font: str, half_points: str) -> None:
    rfonts = get_or_add_child(rpr, "rFonts", 0)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(w_attr(attr), font)

    sz = get_or_add_child(rpr, "sz")
    sz.set(w_attr("val"), half_points)
    szcs = get_or_add_child(rpr, "szCs")
    szcs.set(w_attr("val"), half_points)


def clean_run_emphasis(rpr: etree._Element) -> int:
    return remove_children(
        rpr,
        {
            "b",
            "bCs",
            "i",
            "iCs",
            "u",
            "strike",
            "dstrike",
            "caps",
            "smallCaps",
            "outline",
            "shadow",
            "emboss",
            "imprint",
        },
    )


def get_or_add_rpr(run: etree._Element) -> etree._Element:
    rpr = run.find("w:rPr", namespaces=NS)
    if rpr is not None:
        return rpr
    rpr = etree.Element(w_tag("rPr"))
    run.insert(0, rpr)
    return rpr


def get_or_add_tbl_pr(tbl: etree._Element) -> etree._Element:
    tbl_pr = tbl.find("w:tblPr", namespaces=NS)
    if tbl_pr is not None:
        return tbl_pr
    tbl_pr = etree.Element(w_tag("tblPr"))
    tbl.insert(0, tbl_pr)
    return tbl_pr


def get_or_add_tc_pr(tc: etree._Element) -> etree._Element:
    tc_pr = tc.find("w:tcPr", namespaces=NS)
    if tc_pr is not None:
        return tc_pr
    tc_pr = etree.Element(w_tag("tcPr"))
    tc.insert(0, tc_pr)
    return tc_pr


def normalize_cell_paragraphs(tc: etree._Element, font: str, half_points: str, line_twips: int) -> int:
    cleaned = 0
    for paragraph in tc.findall(".//w:p", namespaces=NS):
        ppr = paragraph.find("w:pPr", namespaces=NS)
        if ppr is None:
            ppr = etree.Element(w_tag("pPr"))
            paragraph.insert(0, ppr)
        spacing = get_or_add_child(ppr, "spacing")
        spacing.set(w_attr("before"), "0")
        spacing.set(w_attr("after"), "0")
        spacing.set(w_attr("line"), str(line_twips))
        spacing.set(w_attr("lineRule"), "auto")
        rpr = ppr.find("w:rPr", namespaces=NS)
        if rpr is None:
            rpr = etree.Element(w_tag("rPr"))
            ppr.append(rpr)
        cleaned += clean_run_emphasis(rpr)
        set_run_font(rpr, font, half_points)
    return cleaned


def _font_from_rpr(rpr: etree._Element | None) -> str | None:
    if rpr is None:
        return None
    rfonts = rpr.find("w:rFonts", namespaces=NS)
    if rfonts is None:
        return None
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        value = rfonts.get(w_attr(attr))
        if value:
            return value
    return None


def _size_from_rpr(rpr: etree._Element | None) -> float | None:
    if rpr is None:
        return None
    size = rpr.find("w:sz", namespaces=NS)
    if size is None:
        size = rpr.find("w:szCs", namespaces=NS)
    if size is None:
        return None
    value = size.get(w_attr("val"))
    if not value:
        return None
    try:
        return float(value) / 2.0
    except ValueError:
        return None


def _style_maps(styles_xml: bytes | None) -> tuple[dict[str, tuple[str | None, float | None]], tuple[str | None, float | None]]:
    if not styles_xml:
        return {}, (None, None)
    try:
        root = parse_xml(styles_xml)
    except etree.XMLSyntaxError:
        return {}, (None, None)

    defaults_rpr = root.find("w:docDefaults/w:rPrDefault/w:rPr", namespaces=NS)
    default_font = _font_from_rpr(defaults_rpr)
    default_size = _size_from_rpr(defaults_rpr)
    styles: dict[str, tuple[str | None, float | None]] = {}
    for style in root.findall("w:style", namespaces=NS):
        if style.get(w_attr("type")) != "paragraph":
            continue
        style_id = style.get(w_attr("styleId"))
        if not style_id:
            continue
        rpr = style.find("w:rPr", namespaces=NS)
        styles[style_id] = (_font_from_rpr(rpr), _size_from_rpr(rpr))
    return styles, (default_font, default_size)


def _iter_body_paragraphs_outside_tables(root: etree._Element) -> list[etree._Element]:
    body = root.find("w:body", namespaces=NS)
    if body is None:
        return []
    paragraphs: list[etree._Element] = []
    for child in body:
        if child.tag == w_tag("p"):
            paragraphs.append(child)
    return paragraphs


def detect_document_typography(docx_path: Path, fallback_font: str = "Tahoma") -> dict[str, object]:
    try:
        with ZipFile(docx_path, "r") as zin:
            document_xml = zin.read("word/document.xml")
            styles_xml = zin.read("word/styles.xml") if "word/styles.xml" in zin.namelist() else None
    except Exception:
        return {"font": fallback_font, "font_source": "fallback", "body_size_pt": None}

    style_map, defaults = _style_maps(styles_xml)
    root = parse_xml(document_xml)
    font_counts: Counter[str] = Counter()
    size_counts: Counter[float] = Counter()

    for paragraph in _iter_body_paragraphs_outside_tables(root):
        style_id = None
        pstyle = paragraph.find("w:pPr/w:pStyle", namespaces=NS)
        if pstyle is not None:
            style_id = pstyle.get(w_attr("val"))
        style_font, style_size = style_map.get(style_id or "", (None, None))
        default_font, default_size = defaults

        paragraph_text = "".join(node.text or "" for node in paragraph.findall(".//w:t", namespaces=NS))
        if not normalize_text(paragraph_text):
            continue
        weight = min(max(len(normalize_text(paragraph_text)), 1), 240)

        if style_font:
            font_counts[style_font] += weight
        if style_size:
            size_counts[style_size] += weight

        for run in paragraph.findall("w:r", namespaces=NS):
            text = "".join(node.text or "" for node in run.findall(".//w:t", namespaces=NS))
            if not normalize_text(text):
                continue
            rpr = run.find("w:rPr", namespaces=NS)
            run_weight = min(max(len(normalize_text(text)), 1), 120)
            font = _font_from_rpr(rpr) or style_font or default_font
            size = _size_from_rpr(rpr) or style_size or default_size
            if font:
                font_counts[font] += run_weight * 2
            if size:
                size_counts[size] += run_weight * 2

    if not font_counts and defaults[0]:
        font_counts[defaults[0]] += 1
    font = font_counts.most_common(1)[0][0] if font_counts else fallback_font
    size = size_counts.most_common(1)[0][0] if size_counts else None
    return {
        "font": font,
        "font_source": "body-or-style" if font_counts else "fallback",
        "body_size_pt": size,
        "font_candidates": dict(font_counts.most_common(8)),
        "size_candidates": {str(k): v for k, v in size_counts.most_common(8)},
    }


def table_text_stats(model) -> dict[str, object]:
    chars = 0
    filled = 0
    max_cell_chars = 0
    for row in model.display_grid:
        for cell in row:
            text = normalize_text(cell)
            chars += len(text)
            max_cell_chars = max(max_cell_chars, len(text))
            if text:
                filled += 1
    cells = max(1, model.width * model.height)
    return {
        "chars": chars,
        "filled_cells": filled,
        "total_cells": model.width * model.height,
        "avg_chars_per_filled_cell": round(chars / max(1, filled), 4),
        "max_cell_chars": max_cell_chars,
        "fill_ratio": round(filled / cells, 4),
    }


def classify_dense(model, stats: dict[str, object], target_width_cm: float | None = None) -> tuple[bool, list[str]]:
    chars = int(stats.get("chars") or 0)
    filled = int(stats.get("filled_cells") or 0)
    avg_chars = float(stats.get("avg_chars_per_filled_cell") or 0)
    data_rows = max(0, model.height - max(0, model.header_row_count))
    column_density = (model.width / target_width_cm) if target_width_cm and target_width_cm > 0 else 0.0

    reasons: list[str] = []
    if model.width >= 7:
        reasons.append("many_columns")
    if data_rows >= 25:
        reasons.append("many_data_rows")
    if chars >= 3000:
        reasons.append("many_characters")
    if filled >= 20 and avg_chars >= 70:
        reasons.append("dense_cells")
    if column_density >= 0.42:
        reasons.append("columns_per_cm")
    return bool(reasons), reasons


def _target_width_for_table(
    document: Document,
    table,
    model,
    section_map: dict[int, int],
    portrait_avail: float,
    landscape_avail: float,
    fallback_width: float,
) -> tuple[float, str, float]:
    mapped_width_cm = _section_available_width_cm(document, table, section_map)
    existing_width_cm = _read_table_total_width_cm(table, fallback_width)
    target_kind = "mapped-section"
    total_width_cm = mapped_width_cm
    if (
        landscape_avail > portrait_avail + 3.0
        and model.width >= 6
        and existing_width_cm >= portrait_avail + 4.0
        and existing_width_cm <= landscape_avail + 1.0
    ):
        total_width_cm = landscape_avail
        target_kind = "wide-landscape"
    return total_width_cm, target_kind, existing_width_cm


def should_skip_table(
    table_index: int,
    page_estimate: int,
    skip_first_tables: int,
    skip_first_pages: int,
) -> tuple[bool, str]:
    if skip_first_tables > 0 and table_index <= skip_first_tables:
        return True, "skipped-first-tables"
    if skip_first_pages > 0 and page_estimate <= skip_first_pages:
        return True, "skipped-first-pages"
    return False, ""


def apply_width_stage(
    input_path: Path,
    stage_path: Path,
    normal_size: float,
    dense_size: float,
    skip_first_tables: int,
    skip_first_pages: int,
    preheader_mode: str,
) -> list[dict[str, object]]:
    document = Document(str(input_path))
    section_map = _body_table_section_map(document)
    page_map = _body_table_page_map(document)
    section_widths = [
        max(8.0, section.page_width.cm - section.left_margin.cm - section.right_margin.cm)
        for section in document.sections
    ]
    portrait_avail = min(section_widths) if section_widths else 17.5
    landscape_avail = max(section_widths) if section_widths else portrait_avail
    fallback_width = _fallback_total_width_cm(document)
    table_meta: list[dict[str, object]] = []

    for idx, table in enumerate(document.tables, start=1):
        model = build_table_model(table, idx)
        page_estimate = page_map.get(idx, 1)
        skip, skip_status = should_skip_table(idx, page_estimate, skip_first_tables, skip_first_pages)

        target_width_cm = 0.0
        target_kind = ""
        existing_width_cm = 0.0
        if model.width > 0 and model.height > 0:
            target_width_cm, target_kind, existing_width_cm = _target_width_for_table(
                document,
                table,
                model,
                section_map,
                portrait_avail,
                landscape_avail,
                fallback_width,
            )
        stats = table_text_stats(model)
        dense, dense_reasons = classify_dense(model, stats, target_width_cm)

        meta: dict[str, object] = {
            "table_index": idx,
            "page_estimate": page_estimate,
            "rows": model.height,
            "columns": model.width,
            "header_rows": model.header_row_count,
            "preheader_rows": model.preheader_row_count,
            "dense": dense,
            "dense_reasons": dense_reasons,
            "font_pt": dense_size if dense else normal_size,
            "target_total_twips": 0,
            "widths_cm": [],
            "target_width_cm": round(target_width_cm, 4),
            "target_kind": target_kind,
            "existing_width_cm": round(existing_width_cm, 4),
            **stats,
        }

        if skip:
            meta["status"] = skip_status
            table_meta.append(meta)
            continue
        if model.width <= 0 or model.height <= 0:
            meta["status"] = "skipped-empty"
            table_meta.append(meta)
            continue

        widths_cm = _content_balanced_widths(model, target_width_cm, preheader_mode=preheader_mode)
        if widths_cm:
            _apply_widths(table, widths_cm)
            meta["status"] = "widths-applied"
        else:
            meta["status"] = "no-widths"
        meta["target_total_twips"] = sum(cm_to_twips(w) for w in widths_cm)
        meta["widths_cm"] = [round(w, 4) for w in widths_cm]
        table_meta.append(meta)

    document.save(str(stage_path))
    return table_meta


def patch_document_xml(
    xml_bytes: bytes,
    original_xml_bytes: bytes,
    table_meta: list[dict[str, object]],
    font: str,
    margin_twips: int,
    border_size: int,
    border_color: str,
    line_twips: int,
) -> tuple[bytes, dict[str, int]]:
    root = parse_xml(xml_bytes)
    original_root = parse_xml(original_xml_bytes)
    body = root.find("w:body", namespaces=NS)
    if body is None:
        return xml_bytes, {"tables_seen": 0}

    tables = [child for child in body if child.tag == w_tag("tbl")]
    original_body = original_root.find("w:body", namespaces=NS)
    original_tables = [child for child in original_body if child.tag == w_tag("tbl")] if original_body is not None else []
    stats = Counter()
    stats["tables_seen"] = len(tables)

    for index, tbl in enumerate(tables):
        meta = table_meta[index] if index < len(table_meta) else {}
        status = str(meta.get("status") or "")
        if status.startswith("skipped-first"):
            if index < len(original_tables):
                parent = tbl.getparent()
                parent.replace(tbl, copy.deepcopy(original_tables[index]))
                stats["tables_preserved_from_original"] += 1
            else:
                stats["tables_skipped_no_original"] += 1
            continue
        if status == "skipped-empty":
            stats["empty_tables_skipped"] += 1
            continue

        half_points = pt_to_half_points(float(meta.get("font_pt", 10.0)))
        target_twips = int(meta.get("target_total_twips") or 0)

        tbl_pr = get_or_add_tbl_pr(tbl)
        if target_twips > 0:
            set_child_attrs(tbl_pr, "tblW", {"w": str(target_twips), "type": "dxa"})
        set_child_attrs(tbl_pr, "tblInd", {"w": "0", "type": "dxa"})
        set_child_attrs(tbl_pr, "tblLayout", {"type": "fixed"})
        set_table_borders(tbl_pr, border_size, border_color)
        set_table_cell_margins(tbl_pr, "tblCellMar", margin_twips)

        for tr in tbl.findall(".//w:tr", namespaces=NS):
            tr_pr = tr.find("w:trPr", namespaces=NS)
            if tr_pr is not None:
                for tr_height in tr_pr.findall("w:trHeight", namespaces=NS):
                    tr_pr.remove(tr_height)
                    stats["row_heights_removed"] += 1

        for tc in tbl.findall(".//w:tc", namespaces=NS):
            tc_pr = get_or_add_tc_pr(tc)
            for no_wrap in tc_pr.findall("w:noWrap", namespaces=NS):
                tc_pr.remove(no_wrap)
                stats["nowrap_removed"] += 1
            set_child_attrs(tc_pr, "vAlign", {"val": "center"})
            set_table_cell_margins(tc_pr, "tcMar", margin_twips)
            stats["paragraph_emphasis_removed"] += normalize_cell_paragraphs(tc, font, half_points, line_twips)
            stats["cells_normalized"] += 1

        for run in tbl.findall(".//w:r", namespaces=NS):
            rpr = get_or_add_rpr(run)
            stats["emphasis_removed"] += clean_run_emphasis(rpr)
            set_run_font(rpr, font, half_points)
            stats["runs_normalized"] += 1
        stats["tables_normalized"] += 1

    return serialize_xml(root), dict(stats)


def patch_package(
    input_path: Path,
    stage_path: Path,
    output_path: Path,
    table_meta: list[dict[str, object]],
    font: str,
    margin_twips: int,
    border_size: int,
    border_color: str,
    line_twips: int,
) -> dict[str, object]:
    patch_stats: dict[str, object] = {}
    with ZipFile(stage_path, "r") as zin:
        infos = zin.infolist()
        data = {info.filename: zin.read(info.filename) for info in infos}

    if "word/document.xml" not in data:
        raise RuntimeError("word/document.xml not found")

    with ZipFile(input_path, "r") as zin_original:
        original_document_xml = zin_original.read("word/document.xml")

    data["word/document.xml"], doc_stats = patch_document_xml(
        data["word/document.xml"],
        original_document_xml,
        table_meta,
        font,
        margin_twips,
        border_size,
        border_color,
        line_twips,
    )
    patch_stats["document_xml"] = doc_stats

    safe_mkdir(output_path.parent)
    with ZipFile(output_path, "w", ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, data[info.filename])

    return patch_stats


def _relative_stem(source: Path, input_dir: Path) -> Path:
    try:
        return source.relative_to(input_dir).with_suffix("")
    except ValueError:
        return Path(source.stem)


def _suffix_path(base_dir: Path, rel_stem: Path, suffix: str, extension: str) -> Path:
    return (base_dir / rel_stem).with_name(f"{rel_stem.name}{suffix}{extension}")


def iter_sources(input_dir: Path, args: argparse.Namespace) -> list[Path]:
    if args.input:
        return [Path(args.input).resolve()]
    if input_dir.is_file():
        eligible = input_dir.suffix.lower() == ".docx" and not input_dir.name.startswith("~$")
        return [input_dir] if eligible else []
    globber = input_dir.rglob if args.recursive else input_dir.glob
    files = sorted(path for path in globber("*.docx") if path.is_file() and not path.name.startswith("~$"))
    if args.all:
        return files
    return files[:1]


def _resolve_project_path(root: Path, value: str) -> Path:
    path = Path(value)
    if not path.is_absolute():
        path = root / path
    return path.resolve()


def _safe_font(args: argparse.Namespace, source: Path) -> tuple[str, dict[str, object]]:
    detected = detect_document_typography(source, fallback_font=args.fallback_font)
    font_value = str(args.font or "").strip()
    if not font_value or font_value.lower() == "auto":
        return str(detected["font"]), detected
    detected["font_override"] = font_value
    return font_value, detected


def _summarize_report(report: dict[str, object]) -> dict[str, object]:
    tables = list(report.get("tables", []))
    return {
        "source_file": report.get("source_file"),
        "output_docx": report.get("output_docx"),
        "tables": len(tables),
        "tables_normalized": sum(1 for row in tables if row.get("status") == "widths-applied"),
        "dense_tables": sum(1 for row in tables if row.get("dense")),
        "skipped_tables": sum(1 for row in tables if str(row.get("status", "")).startswith("skipped")),
        "font": report.get("font"),
        "normal_size_pt": report.get("normal_size_pt"),
        "dense_size_pt": report.get("dense_size_pt"),
    }


def process_one(input_dir: Path, out_dir: Path, report_dir: Path, source: Path, args: argparse.Namespace) -> dict[str, object]:
    if not source.exists():
        raise FileNotFoundError(f"Input file was not found: {source}")
    if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise RuntimeError(f"Unsupported input type: {source.suffix}. Supported: .docx")

    rel_stem = _relative_stem(source, input_dir)
    suffix = "__document_tables_unified"
    out_docx = _suffix_path(out_dir, rel_stem, suffix, ".docx")
    out_json = _suffix_path(report_dir, rel_stem, suffix, ".json")
    safe_mkdir(out_docx.parent)
    safe_mkdir(out_json.parent)

    font, typography = _safe_font(args, source)
    margin_twips = cm_to_twips(args.cell_margin_cm)
    skip_first_tables = max(0, int(args.skip_first_tables or 0))
    skip_first_pages = max(0, int(args.skip_first_pages or 0))
    line_twips = max(180, int(args.line_twips or 240))

    fd, stage_name = tempfile.mkstemp(prefix="document_tables_unifier_", suffix=".docx", dir=str(out_docx.parent))
    os.close(fd)
    stage_path = Path(stage_name)
    try:
        table_meta = apply_width_stage(
            source,
            stage_path,
            args.normal_size,
            args.dense_size,
            skip_first_tables,
            skip_first_pages,
            args.preheader_mode,
        )
        patch_stats = patch_package(
            source,
            stage_path,
            out_docx,
            table_meta,
            font,
            margin_twips,
            args.border_size,
            args.border_color,
            line_twips,
        )
    finally:
        try:
            stage_path.unlink()
        except FileNotFoundError:
            pass

    report = {
        "source_file": source.name,
        "source_path": str(source),
        "relative_path": rel_posix(source, input_dir) if source.is_relative_to(input_dir) else source.name,
        "output_docx": str(out_docx),
        "output_json": str(out_json),
        "operation": "document-tables-unifier",
        "font": font,
        "typography": typography,
        "normal_size_pt": args.normal_size,
        "dense_size_pt": args.dense_size,
        "cell_margin_cm": args.cell_margin_cm,
        "cell_margin_twips": margin_twips,
        "border": {"val": "single", "size": args.border_size, "color": args.border_color},
        "line_twips": line_twips,
        "preheader_mode": args.preheader_mode,
        "skip_first_tables": skip_first_tables,
        "skip_first_pages": skip_first_pages,
        "tables": table_meta,
        "patch_stats": patch_stats,
    }
    report["summary"] = _summarize_report(report)
    write_json_file(out_json, report)
    return report


def write_summary_markdown(summary_path: Path, reports: list[dict[str, object]], input_dir: Path, out_dir: Path, errors: list[dict[str, object]]) -> None:
    safe_mkdir(summary_path.parent)
    lines = [
        "# Document tables unifier",
        "",
        f"Input: `{input_dir}`",
        f"Output: `{out_dir}`",
        "",
        "| DOCX | Tables | Normalized | Dense | Skipped | Font | Output |",
        "|---|---:|---:|---:|---:|---|---|",
    ]
    for report in reports:
        summary = report.get("summary", {})
        lines.append(
            "| "
            + " | ".join(
                [
                    md_escape(str(summary.get("source_file", ""))),
                    str(summary.get("tables", 0)),
                    str(summary.get("tables_normalized", 0)),
                    str(summary.get("dense_tables", 0)),
                    str(summary.get("skipped_tables", 0)),
                    md_escape(str(summary.get("font", ""))),
                    md_escape(str(summary.get("output_docx", ""))),
                ]
            )
            + " |"
        )
    if errors:
        lines.extend(["", "## Errors", ""])
        for error in errors:
            lines.append(f"- `{md_escape(str(error.get('source', '')))}`: {md_escape(str(error.get('error', '')))}")
    summary_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Unify formatting of existing DOCX tables without rebuilding document text."
    )
    parser.add_argument("--input", help="Path to a single DOCX file")
    parser.add_argument("--input-dir", default="input", help="Input folder with DOCX files")
    parser.add_argument("--outdir", default="output/word_excel_tables/document_tables_unified", help="Output folder")
    parser.add_argument("--report-dir", default="report/word_excel_tables/document_tables_unified", help="Report folder")
    parser.add_argument("--all", action="store_true", help="Process all DOCX files from input directory")
    parser.add_argument("--recursive", action="store_true", help="Process DOCX files recursively")
    parser.add_argument("--font", default="auto", help="Table font; use auto to infer from body text")
    parser.add_argument("--fallback-font", default="Tahoma", help="Fallback font when body font cannot be inferred")
    parser.add_argument("--normal-size", type=float, default=10.0, help="Font size for regular tables")
    parser.add_argument("--dense-size", type=float, default=8.0, help="Font size for dense tables")
    parser.add_argument("--cell-margin-cm", type=float, default=0.2, help="DOCX table cell margin in centimeters")
    parser.add_argument("--border-size", type=int, default=4, help="Word border size in eighths of a point")
    parser.add_argument("--border-color", default="000000", help="Border color as RRGGBB")
    parser.add_argument("--line-twips", type=int, default=240, help="Line spacing inside table cells")
    parser.add_argument("--skip-first-tables", type=int, default=0, help="Preserve the first N body tables unchanged")
    parser.add_argument("--skip-first-pages", type=int, default=2, help="Preserve tables estimated on the first N pages")
    parser.add_argument(
        "--preheader-mode",
        choices=("include", "separate"),
        default="separate",
        help="Include pre-header rows in width weighting or keep them out of header weighting.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    root = Path(__file__).resolve().parents[1]
    input_dir = _resolve_project_path(root, args.input_dir)
    out_dir = _resolve_project_path(root, args.outdir)
    report_dir = _resolve_project_path(root, args.report_dir)
    sources = iter_sources(input_dir, args)
    if not sources:
        print(f"[ERROR] No DOCX files were found in: {input_dir}")
        print("[INFO] Put one or more .docx files into the input folder and run again.")
        return 1

    reports: list[dict[str, object]] = []
    errors: list[dict[str, object]] = []
    for source in sources:
        try:
            report = process_one(input_dir, out_dir, report_dir, source, args)
            reports.append(report)
            summary = report["summary"]
            print(
                "[OK] "
                f"{summary['source_file']}: tables={summary['tables']}, "
                f"normalized={summary['tables_normalized']}, dense={summary['dense_tables']}, "
                f"skipped={summary['skipped_tables']}"
            )
            print(f"[OK] Output DOCX: {summary['output_docx']}")
        except Exception as exc:
            errors.append({"source": str(source), "error": str(exc)})
            print(f"[ERROR] {source}: {exc}")

    summary_json = report_dir / "document_tables_unifier_summary.json"
    summary_md = report_dir / "document_tables_unifier_summary.md"
    write_json_file(
        summary_json,
        {
            "input_dir": str(input_dir),
            "out_dir": str(out_dir),
            "processed": [_summarize_report(report) for report in reports],
            "errors": errors,
        },
    )
    write_summary_markdown(summary_md, reports, input_dir, out_dir, errors)
    print(f"[OK] Summary JSON: {summary_json}")
    print(f"[OK] Summary MD: {summary_md}")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
