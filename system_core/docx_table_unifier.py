from __future__ import annotations

import argparse
from pathlib import Path
import sys

from docx import Document
from docx.oxml.ns import qn

CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from docx_table_unifier_model import group_tables, parse_docx_tables
from docx_table_unifier_writers import write_docx, write_json, write_markdown

SUPPORTED_EXTENSIONS = {".docx"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Parse DOCX tables safely, group compatible tables, and export unified results.")
    parser.add_argument("--input", help="Path to a single DOCX file")
    parser.add_argument("--input-dir", default="input", help="Input folder with DOCX files")
    parser.add_argument("--outdir", default="output/word_excel_tables/unified", help="Output folder for generated DOCX files")
    parser.add_argument("--report-dir", default="report/word_excel_tables/unified", help="Output folder for MD/JSON reports")
    parser.add_argument("--all", action="store_true", help="Process all supported files from input directory")
    parser.add_argument("--recursive", action="store_true", help="Process DOCX files recursively under input directory")
    parser.add_argument("--mode", choices=["safe", "balanced", "width-only"], default="safe", help="Grouping mode. Safe is the default.")
    parser.add_argument("--layout", choices=["standard", "merged-sections"], default="standard", help="Output writer layout profile.")
    parser.add_argument("--preheader-mode", choices=["include", "separate"], default="separate", help="Treat first-page pre-header rows as repeat header rows or keep them once before the repeat header.")
    parser.add_argument("--page-size", choices=["A4", "A3"], default="A4", help="Output page size for rebuilt DOCX files.")
    parser.add_argument("--page-orientation", choices=["document", "portrait", "landscape"], default="document", help="Output page orientation.")
    parser.add_argument("--margin-top-mm", type=float, help="Output top margin in millimeters.")
    parser.add_argument("--margin-right-mm", type=float, help="Output right margin in millimeters.")
    parser.add_argument("--margin-bottom-mm", type=float, help="Output bottom margin in millimeters.")
    parser.add_argument("--margin-left-mm", type=float, help="Output left margin in millimeters.")
    return parser.parse_args()


def _resolve_project_path(root: Path, value: str) -> Path:
    path = Path(value)
    if not path.is_absolute():
        path = root / path
    return path.resolve()


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


def _relative_stem(source: Path, input_dir: Path) -> Path:
    try:
        return source.relative_to(input_dir).with_suffix("")
    except ValueError:
        return Path(source.stem)


def _is_landscape(section: object) -> bool:
    return bool(getattr(section, "page_width", 0) > getattr(section, "page_height", 0))


def _source_table_orientation(source: Path) -> str:
    document = Document(str(source))
    sections = list(document.sections)
    if not sections:
        return "portrait"

    current_section = 0
    table_seen = False
    for child in document._body._element.iterchildren():
        if child.tag == qn("w:tbl"):
            table_seen = True
            if _is_landscape(sections[min(current_section, len(sections) - 1)]):
                return "landscape"
        elif child.tag == qn("w:p"):
            ppr = child.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                current_section = min(current_section + 1, len(sections) - 1)

    if not table_seen and any(_is_landscape(section) for section in sections):
        return "landscape"
    return "portrait"


def process_one(input_dir: Path, out_dir: Path, report_dir: Path, source: Path, args: argparse.Namespace) -> int:
    if not source.exists():
        print(f"[ERROR] Input file was not found: {source}")
        return 1
    if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        print(f"[ERROR] Unsupported input type: {source.suffix}. Supported: .docx")
        return 1

    print(f"[INFO] Reading DOCX tables from: {source.name}")
    tables = parse_docx_tables(source)
    if not tables:
        print("[WARN] No top-level tables were found in this DOCX file.")
        return 0

    print(f"[INFO] Parsed {len(tables)} tables.")
    groups = group_tables(tables, mode=args.mode, preheader_mode=args.preheader_mode)
    page_orientation = args.page_orientation
    if page_orientation == "document":
        page_orientation = _source_table_orientation(source)
        print(f"[INFO] Output orientation from source table sections: {page_orientation}")
    out_dir.mkdir(parents=True, exist_ok=True)
    report_dir.mkdir(parents=True, exist_ok=True)
    rel_stem = _relative_stem(source, input_dir)

    suffix = "__merged_sections" if args.layout == "merged-sections" else "__unified"
    out_docx = (out_dir / rel_stem).with_name(f"{rel_stem.name}{suffix}.docx")
    out_md = (report_dir / rel_stem).with_name(f"{rel_stem.name}{suffix}.md")
    out_json = (report_dir / rel_stem).with_name(f"{rel_stem.name}{suffix}.json")
    out_docx.parent.mkdir(parents=True, exist_ok=True)
    out_md.parent.mkdir(parents=True, exist_ok=True)
    out_json.parent.mkdir(parents=True, exist_ok=True)

    write_docx(
        groups,
        out_docx,
        source.name,
        args.mode,
        layout=args.layout,
        preheader_mode=args.preheader_mode,
        page_size=args.page_size,
        page_orientation=page_orientation,
        margin_top_mm=args.margin_top_mm,
        margin_right_mm=args.margin_right_mm,
        margin_bottom_mm=args.margin_bottom_mm,
        margin_left_mm=args.margin_left_mm,
    )
    write_markdown(groups, out_md, source.name, args.mode)
    write_json(groups, out_json, source.name, args.mode, layout=args.layout, preheader_mode=args.preheader_mode, page_size=args.page_size, page_orientation=page_orientation)

    print(f"[OK] Output DOCX : {out_docx}")
    print(f"[OK] Output MD   : {out_md}")
    print(f"[OK] Output JSON : {out_json}")
    return 0


def main() -> int:
    args = parse_args()
    root = Path(__file__).resolve().parents[1]
    input_dir = _resolve_project_path(root, args.input_dir)
    out_dir = _resolve_project_path(root, args.outdir)
    report_dir = _resolve_project_path(root, args.report_dir)
    sources = iter_sources(input_dir, args)
    if not sources:
        print(f"[ERROR] No DOCX files were found in: {input_dir}")
        print("[INFO] Put one or more .docx files into the input folder and run again.")
        return 1

    overall_rc = 0
    for source in sources:
        rc = process_one(input_dir, out_dir, report_dir, source, args)
        if rc != 0:
            overall_rc = rc
    return overall_rc


if __name__ == "__main__":
    raise SystemExit(main())
