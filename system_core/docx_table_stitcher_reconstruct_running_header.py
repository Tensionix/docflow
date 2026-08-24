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
    running_header_rows_removed: int
    running_header_offset_in_first_slice: int
    reconstruction_applied: bool
    preheader_rows_extracted: int
    preheader_mode: str
    running_header_rows: int
    repeat_header_rows: int
    removed_between_paragraphs: int


@dataclass
class ChainReport:
    chain_index: int
    table_indexes_before_merge: List[int]
    column_count: int
    merged_slice_count: int
    reconstruction_applied: bool
    preheader_rows_extracted: int
    preheader_mode: str
    running_header_rows: int
    repeat_header_rows: int
    rows_removed_from_continuations: int


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


def row_signatures(tbl, limit: int = 20) -> list[list[str]]:
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


def best_top_overlap(anchor_tbl, next_tbl, max_probe: int = 20) -> tuple[int, int]:
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


def can_merge(anchor_tbl, next_tbl, between_paragraphs) -> bool:
    if not same_table_shape(anchor_tbl, next_tbl):
        return False
    if any(not is_ignorable_paragraph(p) for p in between_paragraphs):
        return False
    return True


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


def clear_tbl_header_flags(tbl) -> None:
    for row in table_rows(tbl):
        tr_pr = row.find(qn('w:trPr'))
        if tr_pr is None:
            continue
        existing = tr_pr.find(qn('w:tblHeader'))
        if existing is not None:
            tr_pr.remove(existing)


def make_trimmed_table_like(src_tbl, keep_row_count: int):
    new_tbl = copy.deepcopy(src_tbl)
    rows = table_rows(new_tbl)
    for row in rows[keep_row_count:]:
        new_tbl.remove(row)
    clear_tbl_header_flags(new_tbl)
    return new_tbl


def remove_first_rows(tbl, count: int) -> None:
    rows = table_rows(tbl)
    for row in rows[:count]:
        tbl.remove(row)


def make_separator_paragraph():
    p = etree.Element(qn('w:p'))
    p_pr = etree.SubElement(p, qn('w:pPr'))
    spacing = etree.SubElement(p_pr, qn('w:spacing'))
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    spacing.set(qn('w:line'), '20')
    spacing.set(qn('w:lineRule'), 'exact')

    r = etree.SubElement(p, qn('w:r'))
    r_pr = etree.SubElement(r, qn('w:rPr'))
    sz = etree.SubElement(r_pr, qn('w:sz'))
    sz.set(qn('w:val'), '2')
    sz_cs = etree.SubElement(r_pr, qn('w:szCs'))
    sz_cs.set(qn('w:val'), '2')
    t = etree.SubElement(r, qn('w:t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = ' '
    return p




def set_border(el, side: str, val: str = 'single', sz: str = '4', color: str = 'auto', space: str = '0') -> None:
    el.set(qn('w:val'), val)
    el.set(qn('w:sz'), sz)
    el.set(qn('w:space'), space)
    el.set(qn('w:color'), color)


def apply_all_borders_to_row(row, sz: str = '4', color: str = 'auto') -> None:
    for tc in row.xpath('./w:tc', namespaces=NS):
        tc_pr = tc.find(qn('w:tcPr'))
        if tc_pr is None:
            tc_pr = etree.Element(qn('w:tcPr'))
            tc.insert(0, tc_pr)
        tc_borders = tc_pr.find(qn('w:tcBorders'))
        if tc_borders is None:
            tc_borders = etree.SubElement(tc_pr, qn('w:tcBorders'))
        for side in ('top', 'left', 'bottom', 'right'):
            border = tc_borders.find(qn(f'w:{side}'))
            if border is None:
                border = etree.SubElement(tc_borders, qn(f'w:{side}'))
            set_border(border, side, sz=sz, color=color)

def insert_blank_paragraphs(body, index: int, count: int = 1) -> None:
    for _ in range(count):
        body.insert(index, make_separator_paragraph())
        index += 1


def merge_table_into_anchor(anchor_tbl, next_tbl, rows_to_skip: int) -> None:
    rows_to_copy = table_rows(next_tbl)[rows_to_skip:]
    for row in rows_to_copy:
        anchor_tbl.append(copy.deepcopy(row))


def detect_chain_running_header(first_tbl, continuation_tables) -> tuple[int, int]:
    """Return (header_row_count, offset_in_first_tbl).

    Strategy:
    - take the best overlap between first slice and first continuation;
    - require offset > 0 because rows above it become the preheader table;
    - optionally shrink the candidate if later continuations share a shorter
      common prefix.
    """
    if not continuation_tables:
        return 0, 0
    count, offset = best_top_overlap(first_tbl, continuation_tables[0])
    if count <= 0 or offset <= 0:
        return 0, 0

    candidate = row_signatures(first_tbl, limit=offset + count)[offset:offset + count]
    if not candidate:
        return 0, 0

    common = len(candidate)
    for tbl in continuation_tables[1:]:
        top = row_signatures(tbl, limit=common)
        shared = 0
        for idx in range(min(len(candidate), len(top))):
            if candidate[idx] == top[idx]:
                shared += 1
            else:
                break
        if shared == 0:
            # later continuation does not repeat the block; keep the original
            # first-pair decision because many Acrobat exports only repeat on
            # some slices.
            continue
        common = min(common, shared)

    return common, offset


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
        merged_slice_count = 1
        step_index = 0
        reconstruction_applied = False
        preheader_rows_extracted = 0
        running_header_rows = 0
        repeat_header_rows = 0
        rows_removed_total = 0

        # Gather a potential chain first so reconstruction can look ahead.
        chain_positions = []
        chain_tables = []
        probe_i = i
        while True:
            probe_j = probe_i + 1
            probe_between = []
            while probe_j < len(body) and body[probe_j].tag == qn('w:p'):
                probe_between.append(body[probe_j])
                probe_j += 1
            if probe_j >= len(body) or body[probe_j].tag != qn('w:tbl'):
                break
            if not can_merge(body[probe_i], body[probe_j], probe_between):
                break
            chain_positions.append((probe_i, probe_j, list(probe_between)))
            chain_tables.append(body[probe_j])
            probe_i = probe_j

        if chain_tables:
            running_header_rows, preheader_rows_extracted = detect_chain_running_header(anchor, chain_tables)
            if running_header_rows > 0 and preheader_rows_extracted > 0:
                if preheader_mode == "include":
                    repeat_header_rows = preheader_rows_extracted + running_header_rows
                    ensure_tbl_header_on_first_rows(anchor, repeat_header_rows)
                    reconstruction_applied = True
                else:
                    pre_tbl = make_trimmed_table_like(anchor, preheader_rows_extracted)
                    body.insert(i, pre_tbl)
                    i += 1
                    remove_first_rows(anchor, preheader_rows_extracted)
                    repeat_header_rows = running_header_rows
                    ensure_tbl_header_on_first_rows(anchor, repeat_header_rows)
                    pre_rows = table_rows(pre_tbl)
                    anchor_rows = table_rows(anchor)
                    if pre_rows:
                        apply_all_borders_to_row(pre_rows[-1])
                    if anchor_rows:
                        apply_all_borders_to_row(anchor_rows[0])
                    insert_blank_paragraphs(body, i, count=1)
                    i += 1
                    reconstruction_applied = True

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
            if running_header_rows > 0:
                rows_to_remove = running_header_rows if overlap_offset == 0 else 0
            else:
                rows_to_remove = overlap_count if overlap_offset == 0 else 0
                if step_index == 0 and rows_to_remove > 0:
                    running_header_rows = rows_to_remove
                    repeat_header_rows = running_header_rows
                    ensure_tbl_header_on_first_rows(anchor, repeat_header_rows)

            merge_table_into_anchor(anchor, nxt, rows_to_remove)
            rows_removed_total += rows_to_remove
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
                running_header_rows_removed=rows_to_remove,
                running_header_offset_in_first_slice=preheader_rows_extracted if reconstruction_applied else 0,
                reconstruction_applied=reconstruction_applied,
                preheader_rows_extracted=preheader_rows_extracted if reconstruction_applied else 0,
                preheader_mode=preheader_mode,
                running_header_rows=running_header_rows,
                repeat_header_rows=repeat_header_rows or running_header_rows,
                removed_between_paragraphs=len(between),
            ))
            anchor_table_indexes.append(anchor_table_indexes[-1] + 1)

        if merged_slice_count > 1:
            if (repeat_header_rows or running_header_rows) > 0:
                ensure_tbl_header_on_first_rows(anchor, repeat_header_rows or running_header_rows)
            chain_reports.append(ChainReport(
                chain_index=chain_index,
                table_indexes_before_merge=anchor_table_indexes,
                column_count=len(table_column_widths(anchor)),
                merged_slice_count=merged_slice_count,
                reconstruction_applied=reconstruction_applied,
                preheader_rows_extracted=preheader_rows_extracted if reconstruction_applied else 0,
                preheader_mode=preheader_mode,
                running_header_rows=running_header_rows,
                repeat_header_rows=repeat_header_rows or running_header_rows,
                rows_removed_from_continuations=rows_removed_total,
            ))
            chain_index += 1
        i += 1

    final_table_count = len(body.xpath('./w:tbl', namespaces=NS))

    report = {
        'mode': 'reconstruct_running_header_separate_engine',
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
        dst_docx = (output_dir / rel_stem).with_name(f'{rel_stem.name}__stitched_running_header.docx')
        dst_json = (report_dir / rel_stem).with_name(f'{rel_stem.name}__stitched_running_header.json')
        report = stitch_docx_tables(src, dst_docx, dst_json, preheader_mode=preheader_mode)
        results.append(report)
    return results


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='Stitch DOCX tables and reconstruct running headers.')
    parser.add_argument('--input', default='input', help='Input folder with DOCX files')
    parser.add_argument('--outdir', default='output/word_excel_tables/stitched_running_header', help='Output folder for stitched DOCX files')
    parser.add_argument('--report-dir', default='report/word_excel_tables/stitched_running_header', help='Output folder for JSON reports')
    parser.add_argument('--recursive', action='store_true', help='Process DOCX files recursively under input folder')
    parser.add_argument('--preheader-mode', choices=['include', 'separate'], default='separate', help='Keep first-page pre-header as a separate table or include it in repeated header rows.')
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
