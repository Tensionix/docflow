from __future__ import annotations

from pathlib import Path
import argparse
import json
import sys


def detect_python_mode(root: Path) -> str:
    if (root / "runtime" / "python.exe").exists():
        return "portable-runtime"
    if (root / "runtime" / "python" / "python.exe").exists():
        return "portable-runtime"
    return "system-python"


def print_project_info() -> int:
    root = Path(__file__).resolve().parents[1]
    payload = {
        "project_root": str(root),
        "python_executable": sys.executable,
        "python_version": sys.version.split()[0],
        "python_mode": detect_python_mode(root),
        "folders": {
            "input": str(root / "input"),
            "output": str(root / "output"),
            "logs": str(root / "logs"),
            "config": str(root / "config"),
            "data": str(root / "data"),
            "runtime": str(root / "runtime"),
            "wheelhouse": str(root / "wheelhouse"),
            "release": str(root / "release"),
            "portable_powershell": str(root / "system_core" / "powershell"),
        },
        "message": "Portable template is ready. Replace system_core/main.py with your project logic.",
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Audion DocFlow command line.")
    subparsers = parser.add_subparsers(dest="command")

    from system_core import docx_comma_lowercase, morph_replace

    comma = subparsers.add_parser(
        "comma-lowercase-docx",
        parents=[docx_comma_lowercase.build_parser(add_help=False)],
        help="Lowercase first letters after commas in DOCX files.",
    )
    comma.set_defaults(func=docx_comma_lowercase.run)

    morph = subparsers.add_parser(
        "morph-replace",
        parents=[morph_replace.build_parser(add_help=False)],
        help="Morphological find and replace in DOCX/XLSX files.",
    )
    morph.set_defaults(func=morph_replace.run)
    return parser


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv:
        return print_project_info()
    parser = build_parser()
    args = parser.parse_args(argv)
    func = getattr(args, "func", None)
    if not callable(func):
        parser.print_help()
        return 0
    return int(func(args))


if __name__ == "__main__":
    raise SystemExit(main())
