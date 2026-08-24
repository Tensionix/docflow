#!/usr/bin/env python3
"""
DOCX Quality Gate (Hard Fail)

Same checks as docx_quality_gate.py but returns a non-zero exit code if any file fails.
Useful for "regulatory" workflows: do not send document unless gate passes.

Exit codes:
  0  - PASS (no issues)
  1  - FAIL (issues found)
  2+ - runtime/IO error

Usage:
  python docx_quality_gate_hard_fail.py --input input --out report/docx_gate_hard_fail.md
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

from _office_common import safe_mkdir, find_docx_files, rel_posix, write_json_file
from docx_quality_gate import (
    count_nonblack_runs, count_nonblack_styles,
    count_highlights, count_highlights_in_styles,
    count_shading, count_shading_in_styles,
    count_strikethrough, count_strikethrough_in_styles,
    write_quality_docx_report,
    QUALITY_HEADERS_RU, QUALITY_HARD_FAIL_TITLE_RU,
    quality_json_payload,
)
from docx_xml_tools import read_zip_map, find_comments, find_tracked_changes


def main() -> int:
    ap = argparse.ArgumentParser(description="DOCX quality gate with hard-fail exit code.")
    ap.add_argument("--input", default="input", help="Input folder with .docx files (recursive)")
    ap.add_argument("--out", default="report/docx_gate_hard_fail.md", help="Output Markdown report path")
    ap.add_argument("--docx-out", default="", help="Optional DOCX report path")
    ap.add_argument("--no-docx-report", action="store_true", help="Do not write a DOCX report")
    ap.add_argument("--json-out", default="", help="Optional JSON report path")
    args = ap.parse_args()

    in_dir = Path(args.input).resolve()
    out_path = Path(args.out).resolve()
    safe_mkdir(out_path.parent)

    if not in_dir.exists():
        print(f"[ERROR] Input folder does not exist: {in_dir}")
        return 2

    files = find_docx_files(in_dir)
    if not files:
        out_path.write_text(f"# {QUALITY_HARD_FAIL_TITLE_RU}\n\nФайлы DOCX не найдены.\n", encoding="utf-8")
        if not args.no_docx_report:
            docx_out = Path(args.docx_out).resolve() if args.docx_out else out_path.with_suffix(".docx")
            write_quality_docx_report(docx_out, QUALITY_HARD_FAIL_TITLE_RU, in_dir, [], {})
            print(f"[OK] Wrote DOCX report: {docx_out}")
        if args.json_out:
            json_out = Path(args.json_out).resolve()
            write_json_file(json_out, quality_json_payload("docx_quality_gate_hard_fail", in_dir, [], {}))
            print(f"[OK] Wrote JSON report: {json_out}")
        print(f"[WARN] No DOCX files found in: {in_dir}")
        print(f"[OK] Wrote report: {out_path}")
        return 0

    rows = []
    any_fail = False

    for p in files:
        z = read_zip_map(p)
        nonblack = count_nonblack_runs(z)
        style_nonblack = count_nonblack_styles(z)
        highlight = count_highlights(z)
        style_highlight = count_highlights_in_styles(z)
        shading = count_shading(z)
        style_shading = count_shading_in_styles(z)
        comments_parts = find_comments(z)
        comments = 1 if comments_parts else 0
        strike = count_strikethrough(z)
        style_strike = count_strikethrough_in_styles(z)

        tc = find_tracked_changes(z)
        ins = tc.get("<w:ins", 0)
        dele = tc.get("<w:del", 0)
        mvf = tc.get("<w:moveFrom", 0)
        mvt = tc.get("<w:moveTo", 0)

        status = "PASS"
        if any([nonblack, style_nonblack, highlight, style_highlight, shading, style_shading, comments, strike, style_strike, ins, dele, mvf, mvt]):
            status = "FAIL"
            any_fail = True

        rows.append((rel_posix(p, in_dir), status, nonblack, style_nonblack, highlight, style_highlight, shading, style_shading, comments, strike, style_strike, ins, dele, mvf, mvt))

    lines = []
    lines.append(f"# {QUALITY_HARD_FAIL_TITLE_RU}\n")
    lines.append(f"- Папка input: `{in_dir}`")
    lines.append(f"- Файлов проверено: **{len(files)}**\n")

    lines.append("| " + " | ".join(QUALITY_HEADERS_RU) + " |")
    lines.append("|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
    for r in rows:
        lines.append(
            f"| `{r[0]}` | {r[1]} | {r[2]} | {r[3]} | {r[4]} | {r[5]} | {r[6]} | {r[7]} | {r[8]} | {r[9]} | {r[10]} | {r[11]} | {r[12]} | {r[13]} | {r[14]} |"
        )
    lines.append("")

    if any_fail:
        lines.append("## Результат\n- **FAIL**: найдены проблемы качества. Перед отправкой документ нужно исправить.\n")
    else:
        lines.append("## Результат\n- **PASS**: проблем качества не найдено.\n")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    if not args.no_docx_report:
        docx_out = Path(args.docx_out).resolve() if args.docx_out else out_path.with_suffix(".docx")
        write_quality_docx_report(docx_out, QUALITY_HARD_FAIL_TITLE_RU, in_dir, rows, None)
        print(f"[OK] Wrote DOCX report: {docx_out}")
    if args.json_out:
        json_out = Path(args.json_out).resolve()
        write_json_file(json_out, quality_json_payload("docx_quality_gate_hard_fail", in_dir, rows, None))
        print(f"[OK] Wrote JSON report: {json_out}")
    print(f"[OK] Wrote report: {out_path}")
    return 1 if any_fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
