from __future__ import annotations

import copy
import argparse
import json
import re
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


@dataclass
class MergeStep:
    chain_index: int
    step_index: int
    anchor_table_index: int
    merged_table_index: int
    column_count: int
    top_overlap_rows_removed: int
    overlap_offset_in_anchor: int
    repeat_header_enabled_from_top: int
    removed_between_paragraphs: int


@dataclass
class ChainReport:
    chain_index: int
    table_indexes_before_merge: List[int]
    column_count: int
    merged_slice_count: int
    top_overlap_rows_removed_total: int
    repeat_header_rows: int


def qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    if prefix != "w":
        raise ValueError(tag)
    return f"{{{W_NS}}}{local}"


def normalize_text(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()


def paragraph_text(p) -> str:
    texts = [t.text or "" for t in p.xpath('.//w:t', namespaces=NS)]
    return normalize_text("".join(texts))


def is_ignorable_paragraph(p) -> bool:
    return paragraph_text(p) == ""


def table_rows(tbl):
    return tbl.xpath('./w:tr', namespaces=NS)


def cell_texts(row) -> list[str]:
    cells = row.xpath('./w:tc', namespaces=NS)
    out = []
    for tc in cells:
        texts = [t.text or "" for t in tc.xpath('.//w:t', namespaces=NS)]
        out.append(normalize_text("".join(texts)))
    return out


def row_signatures(tbl, limit: int = 16) -> list[list[str]]:
    return [cell_texts(r) for r in table_rows(tbl)[:limit]]


def table_column_widths(tbl) -> list[int]:
    grid_cols = tbl.xpath('./w:tblGrid/w:gridCol', namespaces=NS)
    widths = []
    for gc in grid_cols:
        val = gc.get(qn('w:w')) or gc.get('w:w') or gc.get('w')
        try:
            widths.append(int(val))
        except Exception:
            widths.append(0)
    return widths


def table_style(tbl) -> str:
    vals = tbl.xpath('./w:tblPr/w:tblStyle/@w:val', namespaces=NS)
    return vals[0] if vals else ""


def same_table_shape(tbl_a, tbl_b, width_tolerance: int = 5000) -> bool:
    wa = table_column_widths(tbl_a)
    wb = table_column_widths(tbl_b)
    if len(wa) != len(wb) or not wa:
        return False
    if table_style(tbl_a) != table_style(tbl_b):
        return False
    for a, b in zip(wa, wb):
        if abs(a - b) > width_tolerance:
            return False
    return True


def best_top_overlap(anchor_tbl, next_tbl, max_probe: int = 16) -> tuple[int, int]:
    """Return (overlap_count, offset_in_anchor).

    Standard stitcher is intentionally conservative: it only removes overlap
    when offset == 0, i.e. when the next slice repeats the actual top rows of
    the anchor table.
    """
    a_rows = row_signatures(anchor_tbl, limit=max_probe)
    b_rows = row_signatures(next_tbl, limit=max_probe)
    if not a_rows or not b_rows:
        return 0, 0

    best_count = 0
    best_offset = 0
    max_offset = min(len(a_rows) - 1, max_probe - 1)
    for offset in range(0, max_offset + 1):
        count = 0
        limit = min(len(a_rows) - offset, len(b_rows), max_probe)
        for idx in range(limit):
            if a_rows[offset + idx] == b_rows[idx]:
                count += 1
            else:
                break
        if count > best_count:
            best_count = count
            best_offset = offset
    return best_count, best_offset


def ensure_tbl_header_on_first_rows(tbl, count: int) -> None:
    rows = table_rows(tbl)
    for idx, row in enumerate(rows):
        tr_pr = row.find(qn('w:trPr'))
        if tr_pr is None:
            tr_pr = etree.Element(qn('w:trPr'))
            row.insert(0, tr_pr)
        existing = tr_pr.find(qn('w:tblHeader'))
        if idx < count:
            if existing is None:
                existing = etree.SubElement(tr_pr, qn('w:tblHeader'))
            existing.set(qn('w:val'), 'true')
        elif existing is not None:
            tr_pr.remove(existing)


def can_merge(anchor_tbl, next_tbl, between_paragraphs) -> bool:
    if not same_table_shape(anchor_tbl, next_tbl):
        return False
    if any(not is_ignorable_paragraph(p) for p in between_paragraphs):
        return False
    return True


def merge_table_into_anchor(anchor_tbl, next_tbl, rows_to_skip: int) -> None:
    rows_to_copy = table_rows(next_tbl)[rows_to_skip:]
    for row in rows_to_copy:
        anchor_tbl.append(copy.deepcopy(row))


def stitch_docx_tables(src_docx: Path, dst_docx: Path, dst_json: Path, *, preheader_mode: str = "separate") -> dict:
    with zipfile.ZipFile(src_docx, 'r') as zin:
        xml_bytes = zin.read('word/document.xml')
    root = etree.fromstring(xml_bytes)
    body = root.find(qn('w:body'))
    if body is None:
        raise RuntimeError('word/document.xml has no body')

    merge_steps: list[MergeStep] = []
    chain_reports: list[ChainReport] = []

    original_table_count = len(body.xpath('./w:tbl', namespaces=NS))
    chain_index = 0
    i = 0
    while i < len(body):
        child = body[i]
        if child.tag != qn('w:tbl'):
            i += 1
            continue

        anchor = child
        anchor_table_index = sum(1 for elem in body[:i + 1] if elem.tag == qn('w:tbl'))
        anchor_table_indexes = [anchor_table_index]
        chain_overlap_total = 0
        chain_repeat_header_rows = 0
        merged_slice_count = 1
        step_index = 0

        while True:
            j = i + 1
            between = []
            while j < len(body) and body[j].tag == qn('w:p'):
                between.append(body[j])
                j += 1
            if j >= len(body) or body[j].tag != qn('w:tbl'):
                break

            nxt = body[j]
            if not can_merge(anchor, nxt, between):
                break

            overlap_count, overlap_offset = best_top_overlap(anchor, nxt)
            rows_to_remove = overlap_count if overlap_offset == 0 else 0
            repeat_from_top = rows_to_remove if overlap_offset == 0 else 0

            if step_index == 0 and repeat_from_top > 0:
                chain_repeat_header_rows = repeat_from_top

            merge_table_into_anchor(anchor, nxt, rows_to_remove)
            chain_overlap_total += rows_to_remove
            merged_slice_count += 1

            for elem in between:
                body.remove(elem)
            body.remove(nxt)

            step_index += 1
            merge_steps.append(MergeStep(
                chain_index=chain_index,
                step_index=step_index,
                anchor_table_index=anchor_table_indexes[0],
                merged_table_index=anchor_table_indexes[-1] + 1,
                column_count=len(table_column_widths(anchor)),
                top_overlap_rows_removed=rows_to_remove,
                overlap_offset_in_anchor=overlap_offset,
                repeat_header_enabled_from_top=repeat_from_top,
                removed_between_paragraphs=len(between),
            ))
            anchor_table_indexes.append(anchor_table_indexes[-1] + 1)

        if merged_slice_count > 1:
            if chain_repeat_header_rows > 0:
                ensure_tbl_header_on_first_rows(anchor, chain_repeat_header_rows)
            chain_reports.append(ChainReport(
                chain_index=chain_index,
                table_indexes_before_merge=anchor_table_indexes,
                column_count=len(table_column_widths(anchor)),
                merged_slice_count=merged_slice_count,
                top_overlap_rows_removed_total=chain_overlap_total,
                repeat_header_rows=chain_repeat_header_rows,
            ))
            chain_index += 1
        i += 1

    final_table_count = len(body.xpath('./w:tbl', namespaces=NS))

    report = {
        'mode': 'standard_basic_stitcher',
        'preheader_mode': preheader_mode,
        'input_docx': str(src_docx),
        'output_docx': str(dst_docx),
        'tables_before': original_table_count,
        'tables_after': final_table_count,
        'chains_merged': len(chain_reports),
        'merge_steps': [asdict(x) for x in merge_steps],
        'chain_reports': [asdict(x) for x in chain_reports],
    }

    dst_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src_docx, 'r') as zin, zipfile.ZipFile(dst_docx, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml':
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')
            zout.writestr(item, data)

    dst_json.parent.mkdir(parents=True, exist_ok=True)
    dst_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return report


def _resolve_project_path(project_root: Path, value: str) -> Path:
    path = Path(value)
    if not path.is_absolute():
        path = project_root / path
    return path.resolve()


def _iter_docx_sources(input_dir: Path, recursive: bool) -> list[Path]:
    if input_dir.is_file():
        eligible = input_dir.suffix.lower() == '.docx' and not input_dir.name.startswith('~$')
        return [input_dir] if eligible else []
    globber = input_dir.rglob if recursive else input_dir.glob
    return sorted(path for path in globber('*.docx') if path.is_file() and not path.name.startswith('~$'))


def _relative_stem(source: Path, input_dir: Path) -> Path:
    try:
        return source.relative_to(input_dir).with_suffix('')
    except ValueError:
        return Path(source.stem)


def run_batch(input_dir: Path, output_dir: Path, report_dir: Path, *, recursive: bool = False, preheader_mode: str = "separate") -> list[dict]:
    results = []
    for src in _iter_docx_sources(input_dir, recursive):
        rel_stem = _relative_stem(src, input_dir)
        dst_docx = (output_dir / rel_stem).with_name(f'{rel_stem.name}__stitched.docx')
        dst_json = (report_dir / rel_stem).with_name(f'{rel_stem.name}__stitched.json')
        report = stitch_docx_tables(src, dst_docx, dst_json, preheader_mode=preheader_mode)
        results.append(report)
    return results


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='Stitch sliced DOCX tables conservatively.')
    parser.add_argument('--input', default='input', help='Input folder with DOCX files')
    parser.add_argument('--outdir', default='output/word_excel_tables/stitched', help='Output folder for stitched DOCX files')
    parser.add_argument('--report-dir', default='report/word_excel_tables/stitched', help='Output folder for JSON reports')
    parser.add_argument('--recursive', action='store_true', help='Process DOCX files recursively under input folder')
    parser.add_argument('--preheader-mode', choices=['include', 'separate'], default='separate', help='Compatibility flag for table workflows; standard stitcher does not rebuild pre-header rows.')
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(__file__).resolve().parents[1]
    input_dir = _resolve_project_path(project_root, args.input)
    output_dir = _resolve_project_path(project_root, args.outdir)
    report_dir = _resolve_project_path(project_root, args.report_dir)
    results = run_batch(input_dir, output_dir, report_dir, recursive=args.recursive, preheader_mode=args.preheader_mode)
    print(json.dumps(results, ensure_ascii=False, indent=2))
    if not results:
        print(f'[ERROR] No DOCX files were found in: {input_dir}')
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
