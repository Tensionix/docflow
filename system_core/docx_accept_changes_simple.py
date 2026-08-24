#!/usr/bin/env python3
"""
DOCX Accept Changes (Simple)

Creates a new DOCX where tracked changes are "accepted" in a pragmatic way:
- <w:del> blocks are removed (deletions rejected)
- <w:ins> blocks are unwrapped (insertions accepted)
- moveFrom removed, moveTo unwrapped
- <w:trackRevisions/> removed from settings.xml if present

This is a best-effort tool. For heavily complex revisions, still review the output in Word.

Usage:
  python docx_accept_changes_simple.py --input input --outdir output/accepted_changes
  python docx_accept_changes_simple.py --file x.docx --out output/x_accepted.docx
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
    read_zip_map, write_zip_map, list_xml_parts,
    accept_changes_simple, disable_track_revisions,
)

def process_one(docx_path: Path, out_path: Path) -> None:
    files = read_zip_map(docx_path)

    # Apply to main content parts
    for part in list_xml_parts(files):
        files[part] = accept_changes_simple(files[part])

    # Also apply to the main document.xml
    if "word/document.xml" in files:
        files["word/document.xml"] = accept_changes_simple(files["word/document.xml"])

    # Disable track revisions in settings
    if "word/settings.xml" in files:
        files["word/settings.xml"] = disable_track_revisions(files["word/settings.xml"])

    write_zip_map(out_path, files)

def main() -> int:
    ap = argparse.ArgumentParser(description="Accept tracked changes in DOCX (simple mode).")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--input", help="Input folder (process all .docx recursively)")
    g.add_argument("--file", help="Single DOCX file")
    ap.add_argument("--outdir", default="output/accepted_changes", help="Output folder for batch mode")
    ap.add_argument("--out", help="Output DOCX path for single-file mode")
    args = ap.parse_args()

    if args.input:
        in_dir = Path(args.input).resolve()
        out_dir = Path(args.outdir).resolve()
        safe_mkdir(out_dir)

        docx_files = find_docx_files(in_dir)
        if not docx_files:
            print(f"[WARN] No .docx files found in: {in_dir}")
            return 0

        for p in docx_files:
            out_path = mirrored_output_path(p, in_dir, out_dir)
            safe_mkdir(out_path.parent)
            process_one(p, out_path)

        print(f"[OK] Processed: {len(docx_files)} file(s)")
        print(f"[OK] Output folder: {out_dir}")
        print("[OK] Output mirrors the input folder structure.")
        return 0

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
