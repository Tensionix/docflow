#!/usr/bin/env python3
"""
DOCX Strip Comments

Creates a new DOCX with:
- comment parts removed (word/comments*.xml)
- comment markers removed from document/header/footer XML
- comment relationships removed from *.rels

Usage:
  python docx_strip_comments.py --input input --outdir output/no_comments
  python docx_strip_comments.py --file "x.docx" --out "output/x_no_comments.docx"
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
from pathlib import Path

from _office_common import safe_mkdir, find_docx_files, mirrored_output_path
from docx_xml_tools import (
    read_zip_map, write_zip_map, list_xml_parts, find_comments,
    strip_comment_markers, remove_comment_relationships,
)

def process_one(docx_path: Path, out_path: Path) -> None:
    files = read_zip_map(docx_path)

    # Remove comment parts
    for k in find_comments(files):
        files.pop(k, None)

    # Remove comment references in main parts
    for part in list_xml_parts(files):
        files[part] = strip_comment_markers(files[part])

    # Remove relationships to comments
    for rel_path in [p for p in files.keys() if p.endswith(".rels")]:
        # Typical locations:
        #   word/_rels/document.xml.rels
        #   word/_rels/header1.xml.rels etc.
        files[rel_path] = remove_comment_relationships(files[rel_path])

    write_zip_map(out_path, files)

def main() -> int:
    ap = argparse.ArgumentParser(description="Remove comments from DOCX files.")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--input", help="Input folder (process all .docx recursively)")
    g.add_argument("--file", help="Single DOCX file")
    ap.add_argument("--outdir", default="output/no_comments", help="Output folder for batch mode")
    ap.add_argument("--out", help="Output DOCX path for single-file mode")
    args = ap.parse_args()

    if args.input:
        in_dir = Path(args.input).resolve()
        out_dir = Path(args.outdir).resolve()
        safe_mkdir(out_dir)

        files = find_docx_files(in_dir)
        if not files:
            print(f"[WARN] No .docx files found in: {in_dir}")
            return 0

        for p in files:
            out_path = mirrored_output_path(p, in_dir, out_dir)
            safe_mkdir(out_path.parent)
            process_one(p, out_path)

        print(f"[OK] Processed: {len(files)} file(s)")
        print(f"[OK] Output folder: {out_dir}")
        print("[OK] Output mirrors the input folder structure.")
        return 0

    # Single file
    in_file = Path(args.file).resolve()
    if not in_file.exists():
        print(f"[ERROR] File not found: {in_file}")
        return 2

    if not args.out:
        print("[ERROR] --out is required when using --file")
        return 2

    out_path = Path(args.out).resolve()
    safe_mkdir(out_path.parent)
    process_one(in_file, out_path)
    print(f"[OK] Wrote: {out_path}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
