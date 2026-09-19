#!/usr/bin/env python3
"""Apply a house formatting standard to DOCX copies: margins, font, text sizes.

Margins go onto every section, A3 sheets included. The font replaces the font
of every run and style except symbol fonts, which carry list bullets and
special marks. Body text outside tables takes the body size; headings, the
title and table-of-contents entries keep their own. Tables take the table
size and drop to the overflow size when even the narrowest balanced layout
cannot fit: the longest word of every column, set at the table size, already
needs more width than the page leaves between the margins.
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
from dataclasses import dataclass, field
from pathlib import Path
import re

from lxml import etree

from _office_common import find_docx_files, mirrored_output_path, safe_mkdir, write_json_file
from docx_xml_tools import read_zip_map, write_zip_map

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS

TWIPS_PER_CM = 1440 / 2.54
CHAR_EM = 0.58              # average Tahoma advance for Cyrillic and digits, in em
CELL_MARGINS_TWIPS = 216    # Word's default left + right cell margins
HAIRLINE_COLUMN_TWIPS = 100  # grid columns narrower than this hold nothing readable
DEFAULT_PAGE_WIDTH = 11906
SYMBOL_FONTS = {"symbol", "wingdings", "wingdings 2", "wingdings 3", "webdings",
                "marlett", "mt extra", "zapfdingbats"}
WORD_BREAK = re.compile(r"[\s\-‐‑–—]+")

# CT_ParaRPr / CT_RPr child order: Word rejects a file that breaks it.
RPR_ORDER = ["ins", "del", "moveFrom", "moveTo", "rStyle", "rFonts", "b", "bCs", "i", "iCs",
             "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
             "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
             "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
             "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
             "rPrChange"]
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
             "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
             "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
             "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
             "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
             "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]


@dataclass
class Standard:
    top: int
    bottom: int
    left: int
    right: int
    font: str
    body_half_points: int
    table_half_points: int
    overflow_half_points: int
    table_pt: float


@dataclass
class DocumentResult:
    source: Path
    output: Path
    status: str = "OK"
    error: str = ""
    sections: int = 0
    font_runs: int = 0
    symbol_runs_kept: int = 0
    body_runs_sized: int = 0
    heading_paragraphs_kept: int = 0
    tables_normal: int = 0
    tables_narrowed: int = 0
    tables_overflow: list[dict] = field(default_factory=list)


def local(tag: object) -> str:
    return etree.QName(tag).localname if isinstance(tag, str) else ""


def child_in_order(parent: etree._Element, tag: str, order: list[str]) -> etree._Element:
    """Return the child `tag` of parent, creating it at its schema position."""
    existing = parent.find(W + tag)
    if existing is not None:
        return existing
    element = etree.Element(W + tag)
    position = order.index(tag)
    for index, child in enumerate(parent):
        name = local(child.tag)
        if name in order and order.index(name) > position:
            parent.insert(index, element)
            return element
    parent.append(element)
    return element


def run_properties(run: etree._Element) -> etree._Element:
    rpr = run.find(W + "rPr")
    if rpr is None:
        rpr = etree.Element(W + "rPr")
        run.insert(0, rpr)
    return rpr


def mark_properties(paragraph: etree._Element) -> etree._Element:
    ppr = paragraph.find(W + "pPr")
    if ppr is None:
        ppr = etree.Element(W + "pPr")
        paragraph.insert(0, ppr)
    return child_in_order(ppr, "rPr", PPR_ORDER)


def is_symbol_font(rpr: etree._Element) -> bool:
    fonts = rpr.find(W + "rFonts")
    if fonts is None:
        return False
    names = {(fonts.get(W + key) or "").strip().lower() for key in ("ascii", "hAnsi", "cs", "eastAsia")}
    return bool(names & SYMBOL_FONTS)


def set_font(rpr: etree._Element, font: str) -> bool:
    """Put the house font on a property block; symbol fonts are left alone."""
    if is_symbol_font(rpr):
        return False
    fonts = child_in_order(rpr, "rFonts", RPR_ORDER)
    for attribute in list(fonts.attrib):
        if local(attribute).endswith("Theme"):
            del fonts.attrib[attribute]
    for key in ("ascii", "hAnsi", "cs", "eastAsia"):
        fonts.set(W + key, font)
    return True


def set_size(rpr: etree._Element, half_points: int) -> None:
    for tag in ("sz", "szCs"):
        child_in_order(rpr, tag, RPR_ORDER).set(W + "val", str(half_points))


def keeps_own_size_resolver(styles_root: etree._Element | None):
    """Headings, the title and TOC entries keep their size; everything else is body text."""
    styles: dict[str, tuple[str, str | None, bool]] = {}
    if styles_root is not None:
        for style in styles_root.iter(W + "style"):
            name = style.find(W + "name")
            based = style.find(W + "basedOn")
            outline = style.find(W + "pPr/" + W + "outlineLvl")
            styles[style.get(W + "styleId") or ""] = (
                (name.get(W + "val") if name is not None else "").strip().lower(),
                based.get(W + "val") if based is not None else None,
                outline is not None and (outline.get(W + "val") or "9").isdigit()
                and int(outline.get(W + "val") or "9") < 9,
            )

    def keeps(style_id: str | None, depth: int = 0) -> bool:
        if not style_id or style_id not in styles or depth > 20:
            return False
        name, based, outline = styles[style_id]
        if outline or name.startswith(("heading", "toc")) or name in ("title", "subtitle") or "заголовок" in name:
            return True
        return keeps(based, depth + 1)

    return keeps


def top_level_blocks(container: etree._Element):
    """Body paragraphs and tables in order, looking through content controls."""
    for child in container:
        name = local(child.tag)
        if name in ("p", "tbl"):
            yield child
        elif name == "sdt":
            content = child.find(W + "sdtContent")
            if content is not None:
                yield from top_level_blocks(content)


def is_landscape(sect_pr: etree._Element | None) -> bool:
    size = sect_pr.find(W + "pgSz") if sect_pr is not None else None
    if size is None:
        return False
    width, height = size.get(W + "w") or "", size.get(W + "h") or ""
    if width.isdigit() and height.isdigit():
        return int(width) > int(height)
    return size.get(W + "orient") == "landscape"


PORTRAIT_SPINE_TWIPS = 16838  # the bound left edge of a portrait A4 sheet, 297 mm


def must_turn_to_bind(sect_pr: etree._Element | None) -> bool:
    """A landscape sheet shorter than the spine cannot be bound by its left edge."""
    if not is_landscape(sect_pr):
        return False
    size = sect_pr.find(W + "pgSz")
    height = size.get(W + "h") or "" if size is not None else ""
    return (int(height) if height.isdigit() else 11906) < PORTRAIT_SPINE_TWIPS * 0.95


def sheet_margins(sect_pr: etree._Element | None, standard: Standard) -> tuple[int, int, int, int]:
    """Margins of one sheet as (top, right, bottom, left).

    The standard describes a portrait A4 sheet bound on its left edge. A sheet
    as tall as that edge - A3 landscape among them - is bound the same way and
    unfolds to the right, so it keeps the portrait margins. A4 landscape is too
    short for the spine: it is turned a quarter clockwise to be bound, and its
    margins turn with it, the narrow right edge becoming the bottom one.
    """
    if must_turn_to_bind(sect_pr):
        return standard.left, standard.top, standard.right, standard.bottom
    return standard.top, standard.right, standard.bottom, standard.left


def usable_width(sect_pr: etree._Element | None, standard: Standard) -> int:
    width = DEFAULT_PAGE_WIDTH
    if sect_pr is not None:
        size = sect_pr.find(W + "pgSz")
        if size is not None and (size.get(W + "w") or "").isdigit():
            width = int(size.get(W + "w"))
    _top, right, _bottom, left = sheet_margins(sect_pr, standard)
    return width - left - right


def tables_with_width(body: etree._Element, standard: Standard) -> list[tuple[etree._Element, int]]:
    """Every top-level table with the width its own section leaves between the margins."""
    result: list[tuple[etree._Element, int]] = []
    pending: list[etree._Element] = []
    for block in top_level_blocks(body):
        if local(block.tag) == "tbl":
            pending.append(block)
            continue
        sect = block.find(W + "pPr/" + W + "sectPr")
        if sect is not None:
            width = usable_width(sect, standard)
            result.extend((table, width) for table in pending)
            pending = []
    width = usable_width(body.find(W + "sectPr"), standard)
    result.extend((table, width) for table in pending)
    return result


def cell_text(cell: etree._Element) -> str:
    # Paragraphs of a cell are joined with a space, or words glue across them.
    return " ".join("".join(t.text or "" for t in p.iter(W + "t")) for p in cell.findall(W + "p"))


def narrowest_width(table: etree._Element, size_pt: float) -> int:
    """Width the table needs at its narrowest: each column as wide as its longest word."""
    grid = [int(col.get(W + "w") or 0) if (col.get(W + "w") or "0").isdigit() else 0
            for col in table.findall(W + "tblGrid/" + W + "gridCol")]
    if not grid:
        return 0
    longest = [0] * len(grid)
    for row in table.findall(W + "tr"):
        column = 0
        before = row.find(W + "trPr/" + W + "gridBefore")
        if before is not None and (before.get(W + "val") or "").isdigit():
            column += int(before.get(W + "val"))
        for cell in row.findall(W + "tc"):
            span_el = cell.find(W + "tcPr/" + W + "gridSpan")
            span = int(span_el.get(W + "val")) if span_el is not None and (span_el.get(W + "val") or "").isdigit() else 1
            if span == 1 and column < len(longest):
                words = [word for word in WORD_BREAK.split(cell_text(cell)) if word]
                longest[column] = max([longest[column], *(len(word) for word in words)])
            column += max(span, 1)
    live = [i for i, width in enumerate(grid) if width >= HAIRLINE_COLUMN_TWIPS] or list(range(len(grid)))
    char_twips = size_pt * 20 * CHAR_EM
    return int(sum(longest[i] * char_twips + CELL_MARGINS_TWIPS for i in live))


def narrow_declared_width(table: etree._Element, width: int) -> bool:
    """A table may not declare more than the text: 100%, or the width between the margins.

    Fitting the column grid leaves the declared width alone, and Word lays a
    fixed table out at the declared width, past the margin if it says so.
    """
    declared = table.find(W + "tblPr/" + W + "tblW")
    if declared is None:
        return False
    kind, value = declared.get(W + "type") or "", (declared.get(W + "w") or "").strip()
    if kind == "pct" and value.isdigit() and int(value) > 5000:
        declared.set(W + "w", "5000")
        return True
    if kind == "dxa" and value.isdigit() and int(value) > width:
        declared.set(W + "w", str(width))
        return True
    return False


def apply_margins(root: etree._Element, standard: Standard) -> int:
    count = 0
    for sect in root.iter(W + "sectPr"):
        if local(sect.getparent().tag) == "sectPrChange":
            continue
        margins = sect.find(W + "pgMar")
        if margins is None:
            margins = etree.Element(W + "pgMar")
            size = sect.find(W + "pgSz")
            if size is not None:
                size.addnext(margins)
            else:
                sect.append(margins)
            margins.set(W + "header", "708")
            margins.set(W + "footer", "708")
            margins.set(W + "gutter", "0")
        top, right, bottom, left = sheet_margins(sect, standard)
        margins.set(W + "top", str(top))
        margins.set(W + "right", str(right))
        margins.set(W + "bottom", str(bottom))
        margins.set(W + "left", str(left))
        count += 1
    return count


def apply_styles(styles_root: etree._Element, standard: Standard) -> None:
    defaults = styles_root.find(W + "docDefaults")
    if defaults is None:
        defaults = etree.Element(W + "docDefaults")
        styles_root.insert(0, defaults)
    run_default = defaults.find(W + "rPrDefault")
    if run_default is None:
        run_default = etree.Element(W + "rPrDefault")
        defaults.insert(0, run_default)
    rpr = run_default.find(W + "rPr")
    if rpr is None:
        rpr = etree.SubElement(run_default, W + "rPr")
    set_font(rpr, standard.font)
    set_size(rpr, standard.body_half_points)

    for style in styles_root.iter(W + "style"):
        style_rpr = style.find(W + "rPr")
        if style_rpr is None:
            continue
        if style_rpr.find(W + "rFonts") is not None:
            set_font(style_rpr, standard.font)
        is_normal = style.get(W + "type") == "paragraph" and style.get(W + "default") in ("1", "true", "on")
        if is_normal and style_rpr.find(W + "sz") is not None:
            set_size(style_rpr, standard.body_half_points)


def apply_standard(source: Path, output: Path, standard: Standard) -> DocumentResult:
    result = DocumentResult(source=source, output=output)
    files = read_zip_map(source)
    root = etree.fromstring(files["word/document.xml"])
    styles_root = etree.fromstring(files["word/styles.xml"]) if "word/styles.xml" in files else None
    body = root.find(W + "body")
    keeps_own_size = keeps_own_size_resolver(styles_root)

    result.sections = apply_margins(root, standard)

    table_paragraphs: set[etree._Element] = set()
    for table, width in tables_with_width(body, standard):
        if narrow_declared_width(table, width):
            result.tables_narrowed += 1
        needed = narrowest_width(table, standard.table_pt)
        overflow = needed > width
        half_points = standard.overflow_half_points if overflow else standard.table_half_points
        if overflow:
            result.tables_overflow.append({
                "index": len(result.tables_overflow) + result.tables_normal + 1,
                "needed_cm": round(needed / TWIPS_PER_CM, 1),
                "available_cm": round(width / TWIPS_PER_CM, 1),
            })
        else:
            result.tables_normal += 1
        for paragraph in table.iter(W + "p"):
            table_paragraphs.add(paragraph)
            mark = mark_properties(paragraph)
            set_font(mark, standard.font)
            set_size(mark, half_points)
            for run in paragraph.findall(W + "r"):
                rpr = run_properties(run)
                if run.find(W + "sym") is not None or not set_font(rpr, standard.font):
                    result.symbol_runs_kept += 1
                else:
                    result.font_runs += 1
                set_size(rpr, half_points)

    for paragraph in body.iter(W + "p"):
        if paragraph in table_paragraphs:
            continue
        ppr = paragraph.find(W + "pPr")
        style = ppr.find(W + "pStyle") if ppr is not None else None
        outline = ppr.find(W + "outlineLvl") if ppr is not None else None
        keep_size = keeps_own_size(style.get(W + "val") if style is not None else None) or (
            outline is not None and (outline.get(W + "val") or "9").isdigit() and int(outline.get(W + "val") or "9") < 9
        )
        if keep_size:
            result.heading_paragraphs_kept += 1
        mark = ppr.find(W + "rPr") if ppr is not None else None
        if mark is not None:
            set_font(mark, standard.font)
            if not keep_size:
                set_size(mark, standard.body_half_points)
        for run in paragraph.iter(W + "r"):
            rpr = run_properties(run)
            if run.find(W + "sym") is not None or not set_font(rpr, standard.font):
                result.symbol_runs_kept += 1
            else:
                result.font_runs += 1
            if not keep_size:
                set_size(rpr, standard.body_half_points)
                result.body_runs_sized += 1

    files["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if styles_root is not None:
        apply_styles(styles_root, standard)
        files["word/styles.xml"] = etree.tostring(styles_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    safe_mkdir(output.parent)
    write_zip_map(output, files)
    return result


def write_report(path: Path, input_root: Path, standard_text: str, results: list[DocumentResult]) -> None:
    safe_mkdir(path.parent)
    lines = ["# Стандарт оформления", "", "Папка: `%s`" % input_root, "", standard_text, ""]
    for r in results:
        lines.append("## %s" % r.source.name)
        lines.append("")
        if r.status != "OK":
            lines += ["**Ошибка:** %s" % r.error, ""]
            continue
        lines.append("- разделов с полями по стандарту: %d" % r.sections)
        lines.append("- фрагментов текста со шрифтом стандарта: %d; символьных шрифтов не тронуто: %d"
                     % (r.font_runs, r.symbol_runs_kept))
        lines.append("- фрагментов основного текста с кеглем стандарта: %d" % r.body_runs_sized)
        lines.append("- абзацев заголовков и оглавления со своим кеглем: %d" % r.heading_paragraphs_kept)
        lines.append("- таблиц с кеглем таблиц: %d; тесных, с уменьшенным кеглем: %d"
                     % (r.tables_normal, len(r.tables_overflow)))
        lines.append("- таблиц, объявлявших ширину больше поля, сужено до поля: %d" % r.tables_narrowed)
        for item in r.tables_overflow:
            lines.append("  - таблица %d: самой узкой раскладке нужно %.1f см, между полями %.1f см"
                         % (item["index"], item["needed_cm"], item["available_cm"]))
        lines.append("")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def parse_number(value: str) -> float:
    return float(str(value).replace(",", "."))


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply a house formatting standard to DOCX copies.")
    parser.add_argument("--input", required=True, help="Input folder with DOCX files")
    parser.add_argument("--outdir", default="output/standard_applied", help="Output folder")
    parser.add_argument("--report", default="report/docx_apply_standard.md", help="Markdown report path")
    parser.add_argument("--json-out", default="", help="Optional JSON report path")
    parser.add_argument("--margins-cm", default="2,2,2,1.5", help="top,bottom,left,right in cm")
    parser.add_argument("--font", default="Tahoma", help="Font for text and styles")
    parser.add_argument("--body-size", default="12", help="Body text size, pt")
    parser.add_argument("--table-size", default="10", help="Table text size, pt")
    parser.add_argument("--table-overflow-size", default="9", help="Size for tables that cannot fit, pt")
    args = parser.parse_args()

    try:
        top, bottom, left, right = (parse_number(part) for part in args.margins_cm.split(","))
        body_pt, table_pt, overflow_pt = (parse_number(v) for v in (args.body_size, args.table_size, args.table_overflow_size))
    except ValueError as exc:
        print("[ERROR] Bad standard values: %s" % exc)
        return 2
    standard = Standard(
        top=round(top * TWIPS_PER_CM), bottom=round(bottom * TWIPS_PER_CM),
        left=round(left * TWIPS_PER_CM), right=round(right * TWIPS_PER_CM),
        font=args.font.strip() or "Tahoma",
        body_half_points=round(body_pt * 2), table_half_points=round(table_pt * 2),
        overflow_half_points=round(overflow_pt * 2), table_pt=table_pt,
    )
    standard_text = ("Поля: сверху %s, снизу %s, слева %s, справа %s см. Шрифт %s, текст %s пт, "
                     "таблицы %s пт, тесные таблицы %s пт. Альбомный A4 повёрнут в переплёт: правое "
                     "поле книжного листа у него снизу. A3 продолжает лист вправо и сохраняет поля книжного."
                     % (top, bottom, left, right, standard.font, body_pt, table_pt, overflow_pt))

    input_root = Path(args.input).resolve()
    if not input_root.exists():
        print("[ERROR] Input folder does not exist: %s" % input_root)
        return 2
    out_dir = Path(args.outdir).resolve()
    report_path = Path(args.report).resolve()

    results: list[DocumentResult] = []
    for source in find_docx_files(input_root):
        output = mirrored_output_path(source, input_root, out_dir)
        try:
            result = apply_standard(source, output, standard)
            print("[OK] %s -> %s tables=%d overflow=%d" % (source, output, result.tables_normal + len(result.tables_overflow), len(result.tables_overflow)))
        except Exception as exc:
            result = DocumentResult(source=source, output=output, status="FAILED", error=str(exc))
            print("[FAILED] %s: %s" % (source, exc))
        results.append(result)

    write_report(report_path, input_root, standard_text, results)
    if args.json_out:
        write_json_file(Path(args.json_out).resolve(), {
            "input": str(input_root),
            "documents": [
                {"source": str(r.source), "output": str(r.output), "status": r.status, "error": r.error,
                 "sections": r.sections, "tables_normal": r.tables_normal, "tables_narrowed": r.tables_narrowed,
                 "tables_overflow": r.tables_overflow}
                for r in results
            ],
        })
    print("[OK] Report: %s" % report_path)
    return 0 if all(r.status == "OK" for r in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
